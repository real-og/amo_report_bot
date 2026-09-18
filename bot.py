from __future__ import annotations

import asyncio
import json
import os
import sys
from pathlib import Path

from aiogram import Bot, Dispatcher, types
from aiogram.utils import executor
from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
CONFIG_PATH = BASE_DIR / "config.json"
REPORT_SCRIPT = BASE_DIR / "amocrm_report.py"

PIPELINE_NAMES = (
    "РЕКЛАМА НА ТРАНСПОРТЕ",
    "БРЕНДИРОВАНИЕ АВТО",
    "РЕКЛАМА В ЧЕХЛАХ",
)


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Не заполнена переменная {name} в .env")
    return value


def parse_allowed_user_ids(value: str) -> set[int]:
    try:
        result = {int(item.strip()) for item in value.split(",") if item.strip()}
    except ValueError as error:
        raise RuntimeError("ALLOWED_USER_IDS должен содержать Telegram ID через запятую") from error

    if not result:
        raise RuntimeError("В ALLOWED_USER_IDS не указан ни один Telegram ID")
    return result


def sanitize_filename(value: str) -> str:
    forbidden = '<>:"/\\|?*'
    return "".join("_" if char in forbidden else char for char in value).strip(" .")


def report_paths() -> list[Path]:
    config = json.loads(CONFIG_PATH.read_text("utf-8"))
    output_dir = (CONFIG_PATH.parent / config.get("output_dir", "reports")).resolve()
    return [output_dir / f"{sanitize_filename(name)}.xlsx" for name in PIPELINE_NAMES]


load_dotenv(ENV_PATH)

BOT_TOKEN = require_env("BOT_TOKEN")
ALLOWED_USER_IDS = parse_allowed_user_ids(require_env("ALLOWED_USER_IDS"))

bot = Bot(token=BOT_TOKEN)
dispatcher = Dispatcher(bot)
generation_lock = asyncio.Lock()


async def generate_reports() -> None:
    process = await asyncio.create_subprocess_exec(
        sys.executable,
        str(REPORT_SCRIPT),
        "--config",
        str(CONFIG_PATH),
        "--env",
        str(ENV_PATH),
        cwd=str(BASE_DIR),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    stdout, stderr = await process.communicate()

    if process.returncode != 0:
        error_text = stderr.decode("utf-8", errors="replace").strip()
        if not error_text:
            error_text = stdout.decode("utf-8", errors="replace").strip()
        raise RuntimeError(error_text[-3500:] or "Неизвестная ошибка генерации отчётов")


@dispatcher.message_handler(commands=["generate"])
async def generate_command(message: types.Message) -> None:
    if message.from_user.id not in ALLOWED_USER_IDS:
        return

    if generation_lock.locked():
        await message.answer("Отчёты уже формируются. Подождите завершения.")
        return

    async with generation_lock:
        status_message = await message.answer("Формирую отчёты…")

        try:
            await generate_reports()
            paths = report_paths()
            missing = [path.name for path in paths if not path.is_file()]
            if missing:
                raise RuntimeError(
                    "После генерации не найдены файлы: " + ", ".join(missing)
                )

            for path in paths:
                await message.answer_document(
                    types.InputFile(str(path)),
                    caption=path.stem,
                )

            await status_message.edit_text("Готово: отправлено 3 отчёта.")
        except Exception as error:
            await status_message.edit_text(f"Не удалось сформировать отчёты:\n{error}")


if __name__ == "__main__":
    executor.start_polling(dispatcher, skip_updates=True)
