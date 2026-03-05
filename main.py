
import os
from dotenv import load_dotenv
from telegram.ext import Application, filters, MessageHandler, CommandHandler


async def answer(update, context):
    input = update.message.text
    output = input + ' + Адрес'   # Разработать настоящую обработку данных
    await update.message.reply_text(output)


async def start(update, context):
    await update.message.reply_text('Отправьте Excel‑файл со списком отгрузок')


def main():
    load_dotenv()
    application = Application.builder().token(os.getenv('TOKEN')).build()

    application.add_handler(CommandHandler('start', start))
    application.add_handler(MessageHandler(filters.TEXT, answer))

    print('Чтобы остановить бота, нажмите Ctrl+C')
    application.run_polling()


if __name__ == '__main__':
    main()
