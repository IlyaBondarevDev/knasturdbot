from os import path, makedirs
from aiogram import Bot, Router, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile 
from constants import *
from macro_executor import MacroExecuter


file_dialog_router = Router()


@file_dialog_router.message(Command("start"))
async def start_dialog(message: Message):
    await message.answer(
        START_MESSAGE.format(username=message.from_user.full_name)
    )


@file_dialog_router.message(F.document)
async def get_and_change_file(message: Message, bot: Bot):
    file_path = await download_file(message, bot)
    await change_file(file_path)
    final_file = FSInputFile(file_path)
    
    await bot.send_document(
        chat_id=message.chat.id,
        document=final_file,
        caption='Исправленный файл.')


async def download_file(message: Message, bot: Bot):
    download_path = path.join(path.dirname(
        __file__), '..', 'files',  str(message.chat.id))

    if not path.exists(download_path):
        makedirs(download_path)

    download_path = path.join(download_path, message.document.file_name)
    await bot.download(message.document, download_path)
    
    return download_path


async def change_file(path):
    await MacroExecuter().Execute(path)