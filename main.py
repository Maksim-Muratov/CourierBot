
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


def wake_up(update, context):
    chat = update.effective_chat
    context.bot.send_message(chat_id=chat.id, text='Бот активирован')


updater.dispatcher.add_handler(CommandHandler('start', wake_up))
updater.dispatcher.add_handler(MessageHandler(Filters.text, answer))
updater.start_polling()
updater.idle()   # Бот будет работать до тех пор, пока не нажмете Ctrl+C
