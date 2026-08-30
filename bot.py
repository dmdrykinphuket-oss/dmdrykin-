"""
Телеграм-бот, который отвечает через модель Claude (claude-sonnet-5).

Что делает:
- отвечает только пользователям, чьи Telegram ID указаны в ALLOWED_USERS (.env),
  остальным — вежливый отказ;
- хранит историю диалога отдельно для каждого пользователя (последние 20 сообщений);
- принимает файлы PDF, DOCX и TXT: извлекает текст и держит документ в контексте,
  пока пользователь задаёт по нему вопросы;
- команда /reset очищает историю, /forget убирает документ из контекста;
- умеет искать в интернете (web_search) и показывает ссылки на источники;
- если Anthropic API вернул ошибку — бот пишет об этом в чат и продолжает работать.

Токены читаются из файла .env и в коде не хранятся.
"""

import io
import logging
import os

from anthropic import AsyncAnthropic
import anthropic
import docx
import pypdf
from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# --- Настройки -------------------------------------------------------------

# Загружаем переменные из файла .env в окружение.
load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

# Модель Claude, через которую отвечает бот.
MODEL = "claude-sonnet-5"

# Сколько последних сообщений хранить в истории каждого пользователя.
# 20 сообщений = примерно 10 реплик пользователя и 10 ответов бота.
MAX_HISTORY_MESSAGES = 20

# Максимальная длина ответа модели (в токенах).
MAX_TOKENS = 2000

# --- Веб-поиск -----------------------------------------------------------
# Модель сама решает, искать ли в интернете или ответить по памяти.
# max_uses ограничивает число поисков на один запрос пользователя —
# это защищает от лишних трат.
MAX_WEB_SEARCHES = 3

WEB_SEARCH_TOOL = {
    "type": "web_search_20260209",
    "name": "web_search",
    "max_uses": MAX_WEB_SEARCHES,
}

# Сколько раз повторно запросить модель, если длинный поиск «встал на паузу»
# (stop_reason == "pause_turn").
MAX_PAUSE_RESTARTS = 3

# Telegram не принимает сообщения длиннее 4096 символов.
TELEGRAM_MAX_LEN = 4000

# --- Документы ----------------------------------------------------------
# Какие форматы принимаем.
SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".txt"}

# Максимальный размер присланного файла. Больше — бот вежливо откажет.
# (Telegram и так не даёт ботам скачивать файлы больше ~20 МБ.)
MAX_FILE_SIZE_MB = 10
MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024

# Максимум символов текста из документа, который уходит модели.
# Защищает от гигантских документов и лишних трат (~30 000 токенов).
MAX_DOC_CHARS = 120_000

SYSTEM_PROMPT = (
    "Ты дружелюбный ассистент в Telegram. Отвечай кратко и по делу, "
    "на том же языке, на котором пишет пользователь. "
    "Если вопрос требует свежих или проверяемых фактов — используй веб-поиск. "
    "Если ответ ты и так знаешь — отвечай сразу, без поиска. "
    "Если пользователь прислал документ — отвечай на вопросы, опираясь на него."
)


def parse_allowed_users(raw: str | None) -> set[int]:
    """Превращает строку '123,456' из .env в множество чисел {123, 456}."""
    if not raw:
        return set()
    result = set()
    for part in raw.split(","):
        part = part.strip()
        if part.isdigit():
            result.add(int(part))
    return result


ALLOWED_USERS = parse_allowed_users(os.getenv("ALLOWED_USERS"))

# --- Логирование ----------------------------------------------------------

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
# У httpx слишком подробные логи — приглушаем.
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpx2").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)

# --- Клиент Anthropic и хранилище истории --------------------------------

claude = AsyncAnthropic(api_key=ANTHROPIC_API_KEY)

# История диалогов: {telegram_id: [ {"role": "user"/"assistant", "content": "..."} ]}
# Хранится в оперативной памяти — после перезапуска бота история очищается.
conversations: dict[int, list[dict]] = {}

# Активный документ каждого пользователя:
# {telegram_id: {"filename": str, "text": str, "size": int (байт)}}
documents: dict[int, dict] = {}


def is_allowed(update: Update) -> bool:
    user = update.effective_user
    return user is not None and user.id in ALLOWED_USERS


def human_size(num_bytes: int) -> str:
    """Размер в человекочитаемом виде: 1234 -> '1.2 КБ'."""
    size = float(num_bytes)
    for unit in ("Б", "КБ", "МБ"):
        if size < 1024 or unit == "МБ":
            return f"{size:.0f} {unit}" if unit == "Б" else f"{size:.1f} {unit}"
        size /= 1024


def extract_text(extension: str, data: bytes) -> str:
    """Извлекает текст из файла PDF / DOCX / TXT."""
    if extension == ".pdf":
        reader = pypdf.PdfReader(io.BytesIO(data))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    if extension == ".docx":
        document = docx.Document(io.BytesIO(data))
        return "\n".join(p.text for p in document.paragraphs)
    if extension == ".txt":
        return data.decode("utf-8", errors="replace")
    raise ValueError(f"Формат {extension} не поддерживается")


def _build_messages(history: list[dict], document: dict | None) -> list[dict]:
    """Собирает список сообщений для API. Если есть документ — подкладывает его в начало."""
    if not document:
        return list(history)
    doc_block = {
        "type": "document",
        "source": {
            "type": "text",
            "media_type": "text/plain",
            "data": document["text"],
        },
        "title": document["filename"],
        # Кэшируем документ: повторные вопросы по нему обходятся заметно дешевле.
        "cache_control": {"type": "ephemeral"},
    }
    preamble = [
        {
            "role": "user",
            "content": [
                doc_block,
                {
                    "type": "text",
                    "text": (
                        f"Пользователь прислал файл «{document['filename']}». "
                        "Отвечай на его вопросы, опираясь на этот документ."
                    ),
                },
            ],
        },
        {"role": "assistant", "content": "Документ получен. Задавайте вопросы по нему."},
    ]
    return preamble + list(history)


async def _create_message(messages: list[dict], use_web_search: bool):
    kwargs = dict(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=messages,
    )
    if use_web_search:
        kwargs["tools"] = [WEB_SEARCH_TOOL]
    return await claude.messages.create(**kwargs)


async def _run_conversation(history: list[dict], use_web_search: bool, document: dict | None):
    """Запрос к модели. Если ответ «встал на паузу» посреди поиска — продолжаем его."""
    messages = _build_messages(history, document)
    response = await _create_message(messages, use_web_search)
    restarts = 0
    while response.stop_reason == "pause_turn" and restarts < MAX_PAUSE_RESTARTS:
        messages = messages + [{"role": "assistant", "content": response.content}]
        response = await _create_message(messages, use_web_search)
        restarts += 1
    return response


def _extract_answer(response) -> str:
    """Собирает текст ответа и список источников (ссылок), на которые сослалась модель."""
    text_parts: list[str] = []
    sources: list[tuple[str, str]] = []
    seen: set[str] = set()

    for block in response.content:
        if block.type != "text":
            continue
        if block.text:
            text_parts.append(block.text)
        # У текстовых блоков после веб-поиска есть citations со ссылками на источники.
        for citation in getattr(block, "citations", None) or []:
            url = getattr(citation, "url", None)
            if url and url not in seen:
                seen.add(url)
                sources.append((getattr(citation, "title", None) or url, url))

    answer = "\n".join(text_parts).strip() or "(пустой ответ от модели)"

    if sources:
        links = "\n".join(
            f"{i}. {title}\n{url}" for i, (title, url) in enumerate(sources, 1)
        )
        answer = f"{answer}\n\nИсточники:\n{links}"

    if len(answer) > TELEGRAM_MAX_LEN:
        answer = answer[:TELEGRAM_MAX_LEN] + "…"
    return answer


async def ask_claude(history: list[dict], document: dict | None = None) -> str:
    """Отправляет историю (и документ, если есть) в Claude и возвращает текст ответа.

    Сначала пробуем с веб-поиском. Если поиск недоступен (например, не включён
    для аккаунта Anthropic) — повторяем запрос без него, чтобы бот всё равно ответил.
    """
    try:
        response = await _run_conversation(history, use_web_search=True, document=document)
    except anthropic.BadRequestError:
        logger.warning("Веб-поиск недоступен — отвечаю без него", exc_info=True)
        response = await _run_conversation(history, use_web_search=False, document=document)
    return _extract_answer(response)


# --- Обработчики команд и сообщений --------------------------------------

DENY_TEXT = "Извините, у вас нет доступа к этому боту."


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    await update.message.reply_text(
        "Привет! Я отвечаю с помощью Claude.\n\n"
        "• просто напишите сообщение — отвечу, при необходимости поищу в интернете;\n"
        "• пришлите файл PDF, DOCX или TXT — и задавайте вопросы по нему;\n"
        "• /reset — очистить историю диалога;\n"
        "• /forget — убрать документ из контекста."
    )


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    conversations.pop(update.effective_user.id, None)
    await update.message.reply_text("История диалога очищена.")


async def forget(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    removed = documents.pop(update.effective_user.id, None)
    if removed:
        await update.message.reply_text(
            f"Документ «{removed['filename']}» убран из контекста."
        )
    else:
        await update.message.reply_text("Сейчас в контексте нет документа.")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return

    tg_doc = update.message.document
    filename = tg_doc.file_name or "файл"
    extension = os.path.splitext(filename)[1].lower()

    if extension not in SUPPORTED_EXTENSIONS:
        await update.message.reply_text(
            f"Формат «{extension or '?'}» не поддерживается. "
            "Пришлите, пожалуйста, файл PDF, DOCX или TXT."
        )
        return

    # Предварительная проверка размера (Telegram сообщает его заранее).
    if tg_doc.file_size and tg_doc.file_size > MAX_FILE_SIZE:
        await update.message.reply_text(
            f"Файл слишком большой ({human_size(tg_doc.file_size)}). "
            f"Максимум — {MAX_FILE_SIZE_MB} МБ."
        )
        return

    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action=ChatAction.TYPING
    )

    try:
        tg_file = await tg_doc.get_file()
        data = bytes(await tg_file.download_as_bytearray())
    except Exception:
        logger.exception("Не удалось скачать файл из Telegram")
        await update.message.reply_text(
            "Не получилось скачать файл. Попробуйте прислать его ещё раз."
        )
        return

    if len(data) > MAX_FILE_SIZE:
        await update.message.reply_text(
            f"Файл слишком большой ({human_size(len(data))}). "
            f"Максимум — {MAX_FILE_SIZE_MB} МБ."
        )
        return

    try:
        text = extract_text(extension, data)
    except Exception:
        logger.exception("Ошибка извлечения текста из %s", filename)
        await update.message.reply_text(
            "Не удалось прочитать содержимое файла. "
            "Возможно, он повреждён или защищён паролем."
        )
        return

    if not text.strip():
        await update.message.reply_text(
            "В файле не нашлось текста. Если это скан или картинки — "
            "распознавание пока не поддерживается."
        )
        return

    note = ""
    if len(text) > MAX_DOC_CHARS:
        text = text[:MAX_DOC_CHARS]
        note = (
            f"\n\n⚠️ Документ большой — модели передана только первая часть "
            f"(~{MAX_DOC_CHARS // 1000} тыс. символов)."
        )

    documents[update.effective_user.id] = {
        "filename": filename,
        "text": text,
        "size": len(data),
    }

    await update.message.reply_text(
        f"Файл «{filename}» принят ({human_size(len(data))}). "
        f"Теперь задавайте вопросы по нему. /forget — убрать документ." + note
    )


async def handle_unsupported(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Фото, аудио, видео и прочее, что бот пока не умеет читать."""
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    await update.message.reply_text(
        "Пока я умею работать только с текстом и файлами PDF, DOCX и TXT."
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        logger.info(
            "Отказано в доступе: id=%s username=%s",
            update.effective_user.id if update.effective_user else "?",
            update.effective_user.username if update.effective_user else "?",
        )
        return

    user_id = update.effective_user.id
    user_text = update.message.text

    # Достаём историю пользователя и добавляем новое сообщение.
    history = conversations.setdefault(user_id, [])
    history.append({"role": "user", "content": user_text})

    # Документ этого пользователя (если он присылал файл) — уйдёт модели вместе с историей.
    document = documents.get(user_id)

    # Показываем «печатает...» пока ждём ответ модели.
    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action=ChatAction.TYPING
    )

    try:
        answer = await ask_claude(history, document)
    except anthropic.APIStatusError as e:
        logger.exception("Ошибка Anthropic API")
        # Убираем последнее сообщение пользователя, чтобы не копить «висящие» реплики.
        history.pop()
        await update.message.reply_text(
            f"Ошибка при обращении к Claude (код {e.status_code}). "
            "Попробуйте ещё раз чуть позже."
        )
        return
    except anthropic.APIConnectionError:
        logger.exception("Проблема сети при обращении к Anthropic")
        history.pop()
        await update.message.reply_text(
            "Не удалось связаться с Claude (проблема сети). Попробуйте ещё раз."
        )
        return
    except Exception:
        logger.exception("Непредвиденная ошибка при обращении к Claude")
        history.pop()
        await update.message.reply_text(
            "Что-то пошло не так при обращении к Claude. Попробуйте ещё раз позже."
        )
        return

    history.append({"role": "assistant", "content": answer})

    # Обрезаем историю до последних MAX_HISTORY_MESSAGES сообщений.
    if len(history) > MAX_HISTORY_MESSAGES:
        del history[:-MAX_HISTORY_MESSAGES]

    await update.message.reply_text(answer)


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Ловит всё, что не поймали обработчики, чтобы бот не падал молча."""
    logger.exception("Ошибка в обработчике", exc_info=context.error)


# --- Точка входа ---------------------------------------------------------

def main() -> None:
    if not TELEGRAM_TOKEN:
        raise SystemExit(
            "Не найден TELEGRAM_TOKEN. Впишите его в файл .env "
            "(см. пример в .env.example)."
        )
    if not ANTHROPIC_API_KEY:
        raise SystemExit(
            "Не найден ANTHROPIC_API_KEY. Впишите его в файл .env "
            "(см. пример в .env.example)."
        )
    if not ALLOWED_USERS:
        raise SystemExit(
            "Список ALLOWED_USERS пуст. Впишите в .env хотя бы один Telegram ID."
        )

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("forget", forget))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(
        MessageHandler(filters.PHOTO | filters.AUDIO | filters.VIDEO | filters.VOICE, handle_unsupported)
    )
    app.add_error_handler(on_error)

    logger.info("Бот запущен. Разрешённые пользователи: %s", sorted(ALLOWED_USERS))
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
