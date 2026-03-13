
import pandas as pd
import os
from dotenv import load_dotenv
from telegram.ext import Application, filters, MessageHandler, CommandHandler


async def start(update, context):
    await update.message.reply_text('Отправьте Excel‑файл со списком отгрузок')


async def answer(update, context):
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
            filename=output_filename
            )
    await update.message.reply_text(
        f'Готово! Отгрузки для {entered_name} успешно отфильтрованы!'
        )
    # Удалить название файла из контекста
    context.user_data['current_file'] = None
    # Добавить удаление самих файлов


async def handle_xlsx(update, context):
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


def main():
    load_dotenv()
    application = Application.builder().token(os.getenv('TOKEN')).build()

    application.add_handler(CommandHandler('start', start))
    application.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND,
        answer
        ))
    application.add_handler(MessageHandler(
        filters.Document.FileExtension('xlsx'),
        handle_xlsx
        ))

    print('Чтобы остановить бота, нажмите Ctrl+C')
    application.run_polling()


if __name__ == '__main__':
    main()
