"""
Телеграм-бот, который отвечает через модель Claude (claude-sonnet-5).

Что делает:
- отвечает только пользователям, чьи Telegram ID указаны в ALLOWED_USERS (.env),
  остальным — вежливый отказ;
- хранит историю диалога отдельно для каждого пользователя (последние 20 сообщений);
- команда /reset очищает историю конкретного пользователя;
- если Anthropic API вернул ошибку — бот пишет об этом в чат и продолжает работать.

Токены читаются из файла .env и в коде не хранятся.
"""

import logging
import os

from anthropic import AsyncAnthropic
import anthropic
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

SYSTEM_PROMPT = (
    "Ты дружелюбный ассистент в Telegram. Отвечай кратко и по делу, "
    "на том же языке, на котором пишет пользователь."
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


def is_allowed(update: Update) -> bool:
    user = update.effective_user
    return user is not None and user.id in ALLOWED_USERS


async def ask_claude(history: list[dict]) -> str:
    """Отправляет историю в Claude и возвращает текст ответа."""
    response = await claude.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=history,
    )
    # Ответ приходит списком блоков; берём текстовые и склеиваем.
    parts = [block.text for block in response.content if block.type == "text"]
    return "\n".join(parts).strip() or "(пустой ответ от модели)"


# --- Обработчики команд и сообщений --------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(
            "Извините, у вас нет доступа к этому боту."
        )
        return
    await update.message.reply_text(
        "Привет! Я отвечаю с помощью Claude. Просто напишите мне сообщение.\n"
        "Команда /reset очистит историю нашего диалога."
    )


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(
            "Извините, у вас нет доступа к этому боту."
        )
        return
    conversations.pop(update.effective_user.id, None)
    await update.message.reply_text("История диалога очищена.")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text(
            "Извините, у вас нет доступа к этому боту."
        )
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

    # Показываем «печатает...» пока ждём ответ модели.
    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action=ChatAction.TYPING
    )

    try:
        answer = await ask_claude(history)
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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_error_handler(on_error)

    logger.info("Бот запущен. Разрешённые пользователи: %s", sorted(ALLOWED_USERS))
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
