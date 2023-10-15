import asyncio
from aiogram import Bot, Dispatcher
from config import TOKEN
from handlers.file_dialog import file_dialog_router


async def main():
    bot = Bot(token=TOKEN)
    dp = Dispatcher()
    dp.include_routers(file_dialog_router)
    await bot.delete_webhook(drop_pending_updates=True)
    print('Bot started')
    await dp.start_polling(bot)


if __name__ == '__main__':
    asyncio.run(main())

