import asyncio
import pandas as pd
import os
import sys
from dotenv import load_dotenv
from telegram.ext import Application, filters, MessageHandler, CommandHandler


if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())


async def start(update, context):
    await update.message.reply_text('Отправьте Excel‑файл со списком отгрузок')


async def answer(update, context):
    # Проверка, запущен ли бот
    if not context.application.running:
        return
    entered_name = update.message.text.strip()
    filename = context.user_data.get('current_file')
    if not filename:
        await update.message.reply_text(
            'Сначала отправьте Excel‑файл со списком отгрузок'
            )
        return
    shipment_df = pd.read_excel(filename)
    filtered_df = shipment_df[shipment_df['ФИО Курьера'] == entered_name]
    if filtered_df.empty:
        await update.message.reply_text(
            f'В файле не нашлось строчек "{entered_name}"'
            )
        return
    output_filename = f'Отгрузки {entered_name}.xlsx'
    filtered_df.to_excel(output_filename, index=False)
    with open(output_filename, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename=output_filename,
            read_timeout=60,
            write_timeout=60
            )
    await update.message.reply_text(
        f'Готово! Отгрузки для {entered_name} успешно отфильтрованы!'
        )
    # Удалить название файла из контекста
    context.user_data['current_file'] = None
    # Добавить удаление самих файлов


async def handle_xlsx(update, context):
    # Проверка, запущен ли бот
    if not context.application.running:
        return
    document = update.message.document
    original_filename = document.file_name
    file = await document.get_file()
    await file.download_to_drive(original_filename)
    try:
        pd.read_excel(original_filename)
    except Exception as e:
        await update.message.reply_text(f"Ошибка чтения файла: {e}")
        return
    context.user_data['current_file'] = original_filename
    await update.message.reply_text('Теперь отправьте своё ФИО, как в файле')


async def main():
    load_dotenv()
    application = Application.builder() \
        .token(os.getenv('TOKEN')) \
        .http_version('2.0') \
        .pool_timeout(60.0) \
        .get_updates_http_version('2.0') \
        .build()
    application.add_handler(CommandHandler('start', start))
    application.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND,
        answer
    ))
    application.add_handler(MessageHandler(
        filters.Document.FileExtension('xlsx'),
        handle_xlsx
    ))
    try:
        await application.initialize()
        await application.start()
        await application.updater.start_polling(
            poll_interval=2.0,
            timeout=60,
            allowed_updates=None
        )
        print('Чтобы остановить бота, нажмите Ctrl+C')
        await application.idle()
    except KeyboardInterrupt:
        print("\nБот остановлен пользователем")
    finally:
        if application.running:
            await application.stop()
            await application.updater.stop()
        await application.shutdown()

if __name__ == '__main__':
    asyncio.run(main())
