import os
import random
import requests
from dotenv import load_dotenv

import pandas as pd
import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType

# Максимальное допустимое значение для random_id в VK API
MAX_RANDOM_ID = 2**63 - 1

# Файл "Список адресов ПВЗ":
# Полное название актуального файла в корне проекта
PVZ_REFERENCE_FILE = 'Список адресов ПВЗ от 08.02.2026.xlsx'
# Имя нужного "листа" (Указаны на нижней панели Excel)
PVZ_SHEET_NAME = 'Отчет слоя ПВЗ'
# Названия столбцов
PVZ_CODE_COLUMN = 'Наименование склада'   # Номера ПВЗ
PVZ_ADDRESS_COLUMN = 'Наименование СД'   # Адреса ПВЗ

# Файл "Список курьерских отгрузок"
# Названия столбцов
COURIER_FIO_COLUMN = 'ФИО Курьера'   # ФИО Курьеров
MAIN_PVZ_CODE_COLUMN = 'Получатели'   # Номера ПВЗ
NEW_ADDRESS_COLUMN = 'Адрес ПВЗ'   # (Новая) Адреса ПВЗ

# Словарь file_path текущих файлов по user_id
user_data = {}


def get_random_id():
    """
    Генерирует значение random_id для сообщений.
    """
    return random.randint(1, MAX_RANDOM_ID)


def load_pvz_mapping():
    """
    Загружает список адресов ПВЗ и возвращает словарь {Номер_ПВЗ: Адрес}.
    Если файл не найден или не удалось прочитать, возвращает None.
    """
    if not os.path.exists(PVZ_REFERENCE_FILE):
        print(f'[!!!] Файл {PVZ_REFERENCE_FILE} не найден')
        return None
    try:
        df_ref = pd.read_excel(PVZ_REFERENCE_FILE, sheet_name=PVZ_SHEET_NAME)
        df_ref = df_ref[[PVZ_CODE_COLUMN, PVZ_ADDRESS_COLUMN]].dropna()
        mapping = dict(zip(df_ref[PVZ_CODE_COLUMN], df_ref[PVZ_ADDRESS_COLUMN]))
        return mapping
    except Exception as e:
        print(f'[!!!] Ошибка загрузки списка адресов ПВЗ": {e}')
        return None


PVZ_MAPPING = load_pvz_mapping()


def handle_attachments(event, vk):
    user_id = event.user_id
    attachment_type = event.attachments.get('attach1_type')
    attachment_id = event.attachments.get('attach1')

    if attachment_type != 'doc':
        vk.messages.send(
            user_id=user_id,
            message='Пожалуйста, отправьте Excel‑файл со списком отгрузок',
            random_id=get_random_id()
        )
        return

    # Получаем access_key документа из message
    try:
        message = vk.messages.getById(message_ids=event.message_id)
        attachment = message['items'][0]['attachments'][0]
        access_key = attachment['doc']['access_key']
    except (KeyError, IndexError) as e:
        print(f'[!!!] Ошибка при получении access_key: {e}')
        vk.messages.send(
            user_id=user_id,
            message='Не удалось получить ключ доступа к документу. \n'
            'Проверьте настройки приватности файла.',
            random_id=get_random_id()
        )
        return

    # Получаем документ
    try:
        full_doc_id = f'{attachment_id}_{access_key}'
        doc = vk.docs.getById(docs=full_doc_id)[0]
        file_url = doc['url']
        filename = doc['title']
    except Exception as e:
        print(f'[!!!] Ошибка при получении документа: {e}')
        vk.messages.send(
            user_id=user_id,
            message='Документ не найден или недоступен.',
            random_id=get_random_id()
        )
        return

    # Проверяем расширение файла
    if not filename.lower().endswith('.xlsx'):
        vk.messages.send(
            user_id=user_id,
            message='Пожалуйста, отправьте файл в формате .xlsx',
            random_id=get_random_id()
        )
        return

    # Скачиваем файл
    try:
        response = requests.get(file_url, timeout=30)
        response.raise_for_status()
    except Exception as e:
        vk.messages.send(
            user_id=user_id,
            message=f'Ошибка скачивания файла: {e}',
            random_id=get_random_id()
        )
        return

    file_path = filename
    with open(file_path, 'wb') as f:
        f.write(response.content)

    try:
        pd.read_excel(file_path)
    except Exception as e:
        vk.messages.send(
            user_id=user_id,
            message=f'Файл повреждён или не является Excel: {e}',
            random_id=get_random_id()
        )
        os.remove(file_path)
        return

    user_data[user_id] = file_path

    vk.messages.send(
        user_id=user_id,
        message='Теперь отправьте своё ФИО, как в файле',
        random_id=get_random_id()
    )


def answer(event, vk):
    user_id = event.user_id

    if user_id not in user_data:
        vk.messages.send(
            user_id=user_id,
            message='Сначала отправьте Excel‑файл со списком отгрузок',
            random_id=get_random_id()
        )
        return

    filename = user_data[user_id]
    entered_name = event.text.strip()
    if PVZ_MAPPING is None:
        vk.messages.send(
            user_id=user_id,
            message='Не удалось загрузить список адресов ПВЗ. Адреса не будут добавлены.',
            random_id=get_random_id()
        )

    try:
        shipment_df = pd.read_excel(filename)
        filtered_df = shipment_df[shipment_df[COURIER_FIO_COLUMN] == entered_name].copy()

        if filtered_df.empty:
            vk.messages.send(
                user_id=user_id,
                message=f'В файле не нашлось строчек "{entered_name}"',
                random_id=get_random_id()
            )
        else:
            if PVZ_MAPPING and MAIN_PVZ_CODE_COLUMN in filtered_df.columns:
                filtered_df[NEW_ADDRESS_COLUMN] = filtered_df[MAIN_PVZ_CODE_COLUMN].map(PVZ_MAPPING)
                filtered_df[NEW_ADDRESS_COLUMN] = filtered_df[NEW_ADDRESS_COLUMN].fillna('[Адрес не найден]')
            elif PVZ_MAPPING and MAIN_PVZ_CODE_COLUMN not in filtered_df.columns:
                vk.messages.send(
                    user_id=user_id,
                    message=f'Не могу добавить адреса: не нахожу в этом файле столбец "{MAIN_PVZ_CODE_COLUMN}".',
                    random_id=get_random_id()
                )

            output_filename = f'Отгрузки {entered_name}.xlsx'
            filtered_df.to_excel(output_filename, index=False)

            # Получаем URL для загрузки
            upload_server = vk.docs.getMessagesUploadServer(
                type='doc',
                peer_id=user_id
            )
            upload_url = upload_server['upload_url']

            # Загружаем файл на полученный URL
            with open(output_filename, 'rb') as f:
                files = {'file': (output_filename, f)}
                response = requests.post(upload_url, files=files)
                if response.status_code != 200:
                    raise Exception(
                        f'HTTP ошибка {response.status_code}: {response.text}'
                    )
                upload_response = response.json()

            if 'file' not in upload_response:
                error_msg = upload_response.get(
                    'error',
                    'Отсутствует ключ "file" в ответе сервера'
                )
                raise Exception(f'Ошибка загрузки файла: {error_msg}')

            # Сохраняем документ в сообществе
            saved_doc = vk.docs.save(
                file=upload_response['file'],
                title=output_filename
            )['doc']
            attachment = f'doc{saved_doc["owner_id"]}_{saved_doc["id"]}'

            vk.messages.send(
                user_id=user_id,
                message=f'Готово! Отгрузки для {entered_name} отфильтрованы!',
                attachment=attachment,
                random_id=get_random_id()
            )

    except Exception as e:
        vk.messages.send(
            user_id=user_id,
            message=f'Ошибка обработки файла: {e}',
            random_id=get_random_id()
        )
    finally:
        # Очистка состояния пользователя
        if user_id in user_data:
            del user_data[user_id]
        # Очистка временных файлов
        if 'output_filename' in locals() and os.path.exists(output_filename):
            os.remove(output_filename)
        if os.path.exists(filename):
            os.remove(filename)


def main():
    load_dotenv()
    vk_session = vk_api.VkApi(token=os.getenv('TOKEN'))
    vk = vk_session.get_api()

    print('[i] Чтобы остановить бота, нажмите Ctrl+C')

    for event in VkLongPoll(vk_session).listen():
        if event.type == VkEventType.MESSAGE_NEW and event.to_me:
            if event.attachments:
                handle_attachments(event, vk)
            else:
                answer(event, vk)


if __name__ == '__main__':
    main()
