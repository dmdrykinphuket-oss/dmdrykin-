"""
Телеграм-бот, который отвечает через модель Claude (claude-sonnet-5).

Что делает:
- отвечает только пользователям, чьи Telegram ID указаны в ALLOWED_USERS (.env),
  остальным — вежливый отказ;
- хранит всю историю диалога в SQLite (файл history.db рядом с ботом), отдельно
  по каждому пользователю; модели передаёт последние 20 сообщений;
- при перезапуске продолжает диалог с того места, где остановились;
- принимает файлы PDF, DOCX и TXT: извлекает текст и держит документ в контексте,
  пока пользователь задаёт по нему вопросы;
- принимает голосовые и аудиофайлы (m4a, mp3, …), в том числе присланные
  документом: расшифровывает их локально через faster-whisper (модель small,
  русский), затем прогоняет текст через claude-sonnet-5 — чистит орфографию,
  пунктуацию, слова-паразиты, разбивает на абзацы; для длинных записей добавляет
  структурное резюме. Сырая расшифровка не выводится, доступна по /raw;
- команда /reset очищает историю пользователя, /history показывает её размер,
  /raw показывает сырую расшифровку последнего аудио,
  /forget убирает документ из контекста;
- умеет искать в интернете (web_search) и показывает ссылки на источники;
- если Anthropic API вернул ошибку — бот пишет об этом в чат и продолжает работать.

Токены читаются из файла .env и в коде не хранятся.
"""

import asyncio
import io
import json
import logging
import os
import sqlite3
import threading

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

# Сколько последних сообщений передавать модели в контексте.
# 20 сообщений = примерно 10 реплик пользователя и 10 ответов бота.
# (В базе хранится ВСЯ история — это только окно, которое видит модель.)
MAX_HISTORY_MESSAGES = 20

# Файл базы данных с историей — рядом с bot.py, чтобы путь не зависел от того,
# из какой папки запущен бот.
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.db")

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

# --- Голосовые сообщения ----------------------------------------------
# Расшифровка идёт локально через faster-whisper (модель small, русский язык).
# Модель (~460 МБ) скачивается один раз при первом голосовом и кэшируется.
WHISPER_MODEL_SIZE = "small"
WHISPER_LANGUAGE = "ru"

# Максимальная длительность голосового / аудиофайла. Длиннее — бот вежливо
# откажет (расшифровка долгих записей занимает много времени и памяти).
MAX_VOICE_SECONDS = 120

# После расшифровки текст прогоняется через модель: чистится орфография,
# пунктуация, убираются слова-паразиты и повторы, добавляются абзацы.
# Если сырая расшифровка длиннее порога — добавляется структурное резюме.
TRANSCRIPT_SUMMARY_THRESHOLD = 400
# Не отдаём модели совсем гигантские расшифровки (защита от лишних трат).
MAX_TRANSCRIPT_CHARS = 12_000

# --- Логирование диалога ----------------------------------------------
# Если True — весь диалог (сообщения пользователя и ответы Claude) пишется
# в лог целиком, открытым текстом. Поставьте False, чтобы отключить.
LOG_DIALOG = True

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

# --- Клиент Anthropic ---------------------------------------------------

claude = AsyncAnthropic(api_key=ANTHROPIC_API_KEY)

# Активный документ каждого пользователя (в памяти, при перезапуске сбрасывается):
# {telegram_id: {"filename": str, "text": str, "size": int (байт)}}
documents: dict[int, dict] = {}

# Сырая расшифровка последнего аудио каждого пользователя — для команды /raw
# (в памяти; в чат не выводится).
raw_transcripts: dict[int, str] = {}


# --- История диалогов: SQLite -----------------------------------------
# Вся переписка хранится в файле DB_PATH и переживает перезапуск бота.

_db: sqlite3.Connection | None = None


def get_db() -> sqlite3.Connection:
    """Возвращает подключение к базе, создавая файл и таблицу при первом вызове."""
    global _db
    if _db is None:
        _db = sqlite3.connect(DB_PATH, check_same_thread=False)
        _db.row_factory = sqlite3.Row
        _db.execute("PRAGMA journal_mode=WAL")
        _db.execute(
            """
            CREATE TABLE IF NOT EXISTS messages (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    INTEGER NOT NULL,
                role       TEXT    NOT NULL,   -- 'user' | 'assistant'
                content    TEXT    NOT NULL,
                created_at TEXT    NOT NULL DEFAULT (datetime('now'))
            )
            """
        )
        _db.execute(
            "CREATE INDEX IF NOT EXISTS idx_messages_user ON messages(user_id, id)"
        )
        _db.commit()
    return _db


def db_add_message(user_id: int, role: str, content: str) -> None:
    db = get_db()
    db.execute(
        "INSERT INTO messages (user_id, role, content) VALUES (?, ?, ?)",
        (user_id, role, content),
    )
    db.commit()


def db_recent_messages(user_id: int, limit: int) -> list[dict]:
    """Последние `limit` сообщений пользователя в хронологическом порядке."""
    rows = get_db().execute(
        "SELECT role, content FROM messages WHERE user_id = ? "
        "ORDER BY id DESC LIMIT ?",
        (user_id, limit),
    ).fetchall()
    messages = [{"role": r["role"], "content": r["content"]} for r in reversed(rows)]
    # Первым сообщением для модели обязано быть сообщение пользователя.
    while messages and messages[0]["role"] != "user":
        messages.pop(0)
    return messages


def db_count_messages(user_id: int) -> int:
    return get_db().execute(
        "SELECT COUNT(*) AS n FROM messages WHERE user_id = ?", (user_id,)
    ).fetchone()["n"]


def db_clear_user(user_id: int) -> int:
    db = get_db()
    deleted = db.execute(
        "DELETE FROM messages WHERE user_id = ?", (user_id,)
    ).rowcount
    db.commit()
    return deleted


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


# --- Распознавание речи (faster-whisper) --------------------------------
# Модель тяжёлая, поэтому грузим её один раз и лениво — при первом голосовом.
_whisper_model = None
_whisper_lock = threading.Lock()


def _get_whisper_model():
    global _whisper_model
    if _whisper_model is None:
        with _whisper_lock:
            if _whisper_model is None:
                from faster_whisper import WhisperModel

                logger.info(
                    "Загружаю модель распознавания речи faster-whisper (%s)…",
                    WHISPER_MODEL_SIZE,
                )
                try:
                    # Модель уже в кэше — грузим без обращения к сети
                    # (иначе при плохом интернете загрузка подвисает).
                    _whisper_model = WhisperModel(
                        WHISPER_MODEL_SIZE, device="cpu",
                        compute_type="int8", local_files_only=True,
                    )
                except Exception:
                    logger.info("Модель не найдена в кэше — скачиваю (один раз)…")
                    _whisper_model = WhisperModel(
                        WHISPER_MODEL_SIZE, device="cpu", compute_type="int8"
                    )
                logger.info("Модель распознавания речи готова.")
    return _whisper_model


def transcribe_voice(audio_bytes: bytes) -> str:
    """Расшифровывает аудио (bytes) в текст. Блокирующая операция — звать через to_thread."""
    model = _get_whisper_model()
    segments, _info = model.transcribe(
        io.BytesIO(audio_bytes),
        language=WHISPER_LANGUAGE,
        beam_size=5,
        vad_filter=True,  # отсекает тишину и шум по краям
    )
    return " ".join(segment.text.strip() for segment in segments).strip()


# --- Обработка расшифровки моделью -------------------------------------

_CLEAN_SYSTEM = (
    "Ты редактор расшифровок голосовых сообщений. На входе — сырой текст "
    "распознавания речи. Твоя работа:\n"
    "• исправить орфографию и пунктуацию, расставить абзацы;\n"
    "• убрать слова-паразиты, оговорки, самоповторы, «эканье»;\n"
    "• СОХРАНИТЬ смысл и формулировки автора — ничего не перефразировать, "
    "не добавлять и не додумывать;\n"
    "• не отвечать на текст и не комментировать его.\n"
    "Отвечай строго в формате JSON по заданной схеме, на русском языке."
)

_CLEAN_SCHEMA = {
    "type": "object",
    "properties": {
        "cleaned": {
            "type": "string",
            "description": (
                "Вычищенный текст: орфография, пунктуация, абзацы; без слов-"
                "паразитов и повторов. Смысл и формулировки не менять."
            ),
        },
        "gist": {
            "type": "string",
            "description": (
                "Краткая суть одной строкой. Пустая строка, если резюме не нужно."
            ),
        },
        "key_points": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Ключевые пункты. Пустой массив, если их нет или резюме не нужно.",
        },
        "agreements": {
            "type": "array",
            "items": {"type": "string"},
            "description": (
                "Договорённости, которые ЯВНО прозвучали в тексте. "
                "Пустой массив, если их нет."
            ),
        },
        "tasks": {
            "type": "array",
            "items": {"type": "string"},
            "description": (
                "Задачи и поручения из текста. Пустой массив, если их нет."
            ),
        },
    },
    "required": ["cleaned", "gist", "key_points", "agreements", "tasks"],
    "additionalProperties": False,
}


async def process_transcript(raw: str) -> str:
    """Причёсывает расшифровку через модель и возвращает готовый текст для чата.

    Если raw длиннее TRANSCRIPT_SUMMARY_THRESHOLD — добавляет структурное резюме.
    Бросает исключение, если модель недоступна (вызывающий покажет сырой текст).
    """
    need_summary = len(raw) > TRANSCRIPT_SUMMARY_THRESHOLD

    if need_summary:
        task = (
            "Запись длинная — помимо cleaned заполни резюме по её содержанию: "
            "gist (суть одной строкой), key_points (ключевые пункты), "
            "agreements (договорённости) и tasks (задачи). "
            "Блок оставляй пустым, если в тексте этого нет."
        )
    else:
        task = (
            "Запись короткая — резюме не нужно: gist оставь пустой строкой, "
            "key_points, agreements и tasks — пустыми массивами."
        )

    response = await claude.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=_CLEAN_SYSTEM,
        messages=[{
            "role": "user",
            "content": f"{task}\n\nСырая расшифровка:\n---\n{raw}",
        }],
        output_config={"format": {"type": "json_schema", "schema": _CLEAN_SCHEMA}},
    )
    text = next(b.text for b in response.content if b.type == "text")
    data = json.loads(text)

    cleaned = (data.get("cleaned") or "").strip()
    if not cleaned:
        raise ValueError("модель вернула пустой cleaned")

    parts = [cleaned]
    if need_summary:
        blocks = []
        gist = (data.get("gist") or "").strip()
        if gist:
            blocks.append(f"📌 Суть: {gist}")
        if data.get("key_points"):
            blocks.append(
                "🔑 Ключевые пункты:\n"
                + "\n".join(f"• {p}" for p in data["key_points"])
            )
        if data.get("agreements"):
            blocks.append(
                "🤝 Договорённости:\n"
                + "\n".join(f"• {a}" for a in data["agreements"])
            )
        if data.get("tasks"):
            blocks.append(
                "✅ Задачи:\n" + "\n".join(f"• {t}" for t in data["tasks"])
            )
        if blocks:
            parts.append("— — — — —\n" + "\n\n".join(blocks))

    return "\n\n".join(parts)


async def _safe(coro) -> None:
    """Ждёт корутину и глотает сетевые ошибки — для косметических действий
    (правка/удаление статуса), сбой которых не должен ронять обработчик."""
    try:
        await coro
    except Exception:
        logger.debug("Второстепенное действие не удалось", exc_info=True)


async def deliver(message, text: str, retries: int = 3) -> bool:
    """Настойчиво отправляет ответ (важный результат). Возвращает True при успехе."""
    for attempt in range(retries):
        try:
            await send_chunks(message, text)
            return True
        except Exception:
            logger.warning(
                "Не удалось отправить ответ (попытка %s/%s)", attempt + 1, retries,
                exc_info=True,
            )
            await asyncio.sleep(2 * (attempt + 1))
    logger.error("Ответ так и не отправлен, текст: %s", text[:500])
    return False


async def send_chunks(message, text: str) -> None:
    """Отправляет длинный текст, разбивая по абзацам на части ≤ 4096 символов."""
    limit = 4096
    if len(text) <= limit:
        await message.reply_text(text)
        return
    chunk = ""
    for para in text.split("\n\n"):
        while len(para) > limit:  # одиночный абзац длиннее лимита — режем жёстко
            if chunk:
                await message.reply_text(chunk.strip())
                chunk = ""
            await message.reply_text(para[:limit])
            para = para[limit:]
        if len(chunk) + len(para) + 2 > limit and chunk:
            await message.reply_text(chunk.strip())
            chunk = ""
        chunk = f"{chunk}\n\n{para}" if chunk else para
    if chunk.strip():
        await message.reply_text(chunk.strip())


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
        "• запишите голосовое или пришлите аудиофайл — я расшифрую его, "
        "причешу текст и (если запись длинная) сделаю краткое резюме;\n"
        "• /raw — сырая расшифровка последнего аудио;\n"
        "• /history — сколько сообщений сохранено;\n"
        "• /reset — очистить историю диалога;\n"
        "• /forget — убрать документ из контекста."
    )


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    deleted = db_clear_user(update.effective_user.id)
    await update.message.reply_text(
        f"История диалога очищена (удалено сообщений: {deleted})."
    )


async def history_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    total = db_count_messages(update.effective_user.id)
    if total == 0:
        await update.message.reply_text("История пуста.")
        return
    shown = min(total, MAX_HISTORY_MESSAGES)
    await update.message.reply_text(
        f"В истории сохранено сообщений: {total}.\n"
        f"Модели передаю последние {shown}."
    )


async def raw_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Показывает сырую (необработанную) расшифровку последнего аудио."""
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    raw = raw_transcripts.get(update.effective_user.id)
    if not raw:
        await update.message.reply_text(
            "Пока нет расшифровок. Пришлите голосовое или аудиофайл."
        )
        return
    await send_chunks(
        update.message, f"Сырая расшифровка последнего аудио:\n\n{raw}"
    )


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
    """Фото, аудио-файлы, видео и прочее, что бот пока не умеет читать."""
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    await update.message.reply_text(
        "Пока я умею работать только с текстом, голосовыми и файлами PDF, DOCX и TXT."
    )


async def get_model_answer(
    update: Update, context: ContextTypes.DEFAULT_TYPE, user_id: int, user_text: str
) -> str | None:
    """Спрашивает модель по истории из БД и сохраняет обе реплики.

    В БД пишем только после успешного ответа — тогда «висящих» вопросов
    без ответа не остаётся. При ошибке сам пишет пояснение в чат и возвращает
    None (бот не падает).
    """
    document = documents.get(user_id)

    # Контекст для модели: последние сообщения из БД + текущий вопрос.
    history = db_recent_messages(user_id, MAX_HISTORY_MESSAGES - 1)
    history.append({"role": "user", "content": user_text})

    if LOG_DIALOG:
        doc_note = f" [документ: {document['filename']}]" if document else ""
        logger.info("[%s] пользователь%s: %s", user_id, doc_note, user_text)

    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action=ChatAction.TYPING
    )

    try:
        answer = await ask_claude(history, document)
    except anthropic.APIStatusError as e:
        logger.exception("Ошибка Anthropic API")
        await update.message.reply_text(
            f"Ошибка при обращении к Claude (код {e.status_code}). "
            "Попробуйте ещё раз чуть позже."
        )
        return None
    except anthropic.APIConnectionError:
        logger.exception("Проблема сети при обращении к Anthropic")
        await update.message.reply_text(
            "Не удалось связаться с Claude (проблема сети). Попробуйте ещё раз."
        )
        return None
    except Exception:
        logger.exception("Непредвиденная ошибка при обращении к Claude")
        await update.message.reply_text(
            "Что-то пошло не так при обращении к Claude. Попробуйте ещё раз позже."
        )
        return None

    db_add_message(user_id, "user", user_text)
    db_add_message(user_id, "assistant", answer)

    if LOG_DIALOG:
        logger.info("[%s] бот: %s", user_id, answer)
    return answer


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        logger.info(
            "Отказано в доступе: id=%s username=%s",
            update.effective_user.id if update.effective_user else "?",
            update.effective_user.username if update.effective_user else "?",
        )
        return

    answer = await get_model_answer(
        update, context, update.effective_user.id, update.message.text
    )
    if answer is not None:
        await update.message.reply_text(answer)


async def _transcribe_and_reply(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    tg_object,
    duration: int | None,
    kind: str,
) -> None:
    """Общий путь для голосовых и аудиофайлов: скачать → расшифровать → ответить.

    tg_object — telegram.Voice / telegram.Audio / telegram.Document (у всех есть
    get_file() и file_size). kind — слово для сообщений («Голосовое» / «Аудио»).
    """
    logger.info(
        "%s получено: duration=%s size=%s mime=%s",
        kind, duration, tg_object.file_size, getattr(tg_object, "mime_type", None),
    )

    if duration and duration > MAX_VOICE_SECONDS:
        await update.message.reply_text(
            f"{kind} слишком длинное ({duration} с). "
            f"Максимум — {MAX_VOICE_SECONDS} с. Разбейте на части или напишите текстом."
        )
        return

    if tg_object.file_size and tg_object.file_size > MAX_FILE_SIZE:
        await update.message.reply_text(
            f"Файл слишком большой ({human_size(tg_object.file_size)}). "
            f"Максимум — {MAX_FILE_SIZE_MB} МБ."
        )
        return

    status = None
    try:
        status = await update.message.reply_text(
            f"🎧 Расшифровываю ({kind.lower()})… это не мгновенно, подождите минуту."
        )
    except Exception:
        logger.warning("Не удалось отправить статус-сообщение", exc_info=True)

    async def fail(text: str) -> None:
        """Показать сообщение об ошибке (через статус или новым сообщением)."""
        if status is not None:
            await _safe(status.edit_text(text))
        else:
            await deliver(update.message, text)

    try:
        tg_file = await tg_object.get_file()
        audio_bytes = bytes(await tg_file.download_as_bytearray())
    except Exception:
        logger.exception("Не удалось скачать аудио из Telegram")
        await fail("Не удалось скачать файл. Попробуйте ещё раз.")
        return

    if len(audio_bytes) > MAX_FILE_SIZE:
        await fail(
            f"Файл слишком большой ({human_size(len(audio_bytes))}). "
            f"Максимум — {MAX_FILE_SIZE_MB} МБ."
        )
        return

    logger.info("Аудио скачано: %s байт, начинаю расшифровку", len(audio_bytes))
    try:
        recognized = await asyncio.to_thread(transcribe_voice, audio_bytes)
    except Exception:
        logger.exception("Ошибка расшифровки аудио")
        await fail(
            "Не удалось расшифровать запись. "
            "Попробуйте другой файл или напишите текстом."
        )
        return

    logger.info("Расшифровка готова: %r", (recognized or "")[:200])
    recognized = (recognized or "").strip()
    if not recognized:
        await fail(
            "В этой записи не удалось разобрать речь. "
            "Попробуйте другую запись или напишите текстом."
        )
        return

    user_id = update.effective_user.id
    # Сырую расшифровку сохраняем (для /raw), но в чат не выводим.
    raw_transcripts[user_id] = recognized
    if LOG_DIALOG:
        logger.info("[%s] аудио, сырая расшифровка: %s", user_id, recognized)

    if status is not None:
        await _safe(status.edit_text("✍️ Причёсываю текст…"))
    try:
        result = await process_transcript(recognized[:MAX_TRANSCRIPT_CHARS])
    except Exception:
        logger.exception("Не удалось обработать расшифровку моделью")
        result = (
            "⚠️ Не получилось обработать текст через модель — вот сырая "
            f"расшифровка:\n\n{recognized}"
        )

    if status is not None:
        await _safe(status.delete())

    await deliver(update.message, result)
    if LOG_DIALOG:
        logger.info("[%s] аудио, результат: %s", user_id, result)


async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    voice = update.message.voice
    await _transcribe_and_reply(update, context, voice, voice.duration, "Голосовое")


async def handle_audio(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Аудиофайлы: и присланные как аудио (message.audio), и как документ (audio/*)."""
    if not is_allowed(update):
        await update.message.reply_text(DENY_TEXT)
        return
    audio = update.message.audio or update.message.document
    duration = getattr(update.message.audio, "duration", None)
    await _transcribe_and_reply(update, context, audio, duration, "Аудио")


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

    # Открываем базу истории и показываем, что в ней уже есть.
    db = get_db()
    stats = db.execute(
        "SELECT COUNT(*) AS n, COUNT(DISTINCT user_id) AS u FROM messages"
    ).fetchone()
    logger.info(
        "История: %s сообщений от %s пользователей (%s)",
        stats["n"], stats["u"], DB_PATH,
    )

    # Таймауты побольше — сеть на этой машине бывает нестабильной.
    app = (
        Application.builder()
        .token(TELEGRAM_TOKEN)
        .connect_timeout(20)
        .read_timeout(20)
        .write_timeout(20)
        .pool_timeout(10)
        .get_updates_connect_timeout(20)
        .get_updates_read_timeout(40)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("history", history_cmd))
    app.add_handler(CommandHandler("raw", raw_cmd))
    app.add_handler(CommandHandler("forget", forget))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    # Аудио: и как аудио-сообщение, и как документ audio/* — ДО общего
    # обработчика документов, иначе .m4a перехватится как обычный документ.
    app.add_handler(
        MessageHandler(
            filters.AUDIO | filters.Document.Category("audio/"), handle_audio
        )
    )
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(
        MessageHandler(filters.PHOTO | filters.VIDEO, handle_unsupported)
    )
    app.add_error_handler(on_error)

    # Заранее греем модель распознавания речи в фоне, чтобы первое голосовое
    # не ждало загрузки модели.
    threading.Thread(target=_get_whisper_model, daemon=True).start()

    logger.info("Бот запущен. Разрешённые пользователи: %s", sorted(ALLOWED_USERS))
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
