
import os
from dotenv import load_dotenv
from telegram.ext import Updater, Filters, MessageHandler, CommandHandler


load_dotenv()
token = os.getenv('TOKEN')
updater = Updater(token)


def answer(update, context):
    chat = update.effective_chat
    input = update.message.text
    output = input + ' + Адрес'   # Разработать настоящую обработку данных
    context.bot.send_message(chat_id=chat.id, text=output)


def start(update, context):
    update.message.reply_text('Отправьте Excel‑файл со списком отгрузок')


updater.dispatcher.add_handler(CommandHandler('start', start))
updater.dispatcher.add_handler(MessageHandler(Filters.text, answer))
updater.start_polling()

print('Чтобы остановить бота, нажмите Ctrl+C')
updater.idle()
