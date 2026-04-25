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
        mapping = dict(zip(df_ref[PVZ_CODE_COLUMN],
                           df_ref[PVZ_ADDRESS_COLUMN]))
        return mapping
    except Exception as e:
        print(f'[!!!] Ошибка загрузки списка адресов ПВЗ": {e}')
        return None


PVZ_MAPPING = load_pvz_mapping()


def handle_attachments(event, vk):
    user_id = event.user_id

    # Проверяем тип вложения
    attachment_type = event.attachments.get('attach1_type')
    if attachment_type != 'doc':
        vk.messages.send(
            user_id=user_id,
            message='Пожалуйста, отправьте Excel‑файл со списком отгрузок',
            random_id=get_random_id()
        )
        return

    # Получаем данные о файле
    try:
        message = vk.messages.getById(message_ids=event.message_id)
        file = message['items'][0]['attachments'][0]['doc']
        file_url = file['url']
        filename = file['title']
        file_extension = file['ext']
    except Exception as e:
        print(f'[!!!] Ошибка при получении данных о файле: {e}')
        vk.messages.send(
            user_id=user_id,
            message='Не удалось получить данные о файле',
            random_id=get_random_id()
        )
        return

    # Проверяем расширение файла
    if file_extension != 'xlsx':
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
        print(f'[!!!] Ошибка скачивания файла: {e}')
        vk.messages.send(
            user_id=user_id,
            message='Не удалось скачать файл',
            random_id=get_random_id()
        )
        return

    # Сохраняем файл
    with open(filename, 'wb') as f:
        f.write(response.content)

    # Проверяем читаемость файла
    try:
        pd.read_excel(filename)
    except Exception as e:
        print(f'[!!!] Ошибка чтения файла: {e}')
        vk.messages.send(
            user_id=user_id,
            message='Не удалось прочитать файл',
            random_id=get_random_id()
        )
        os.remove(filename)
        return

    # Сохраняем данные о пользователе
    user_data[user_id] = filename

    vk.messages.send(
        user_id=user_id,
        message='Теперь отправьте своё ФИО, как в файле',
        random_id=get_random_id()
    )


def answer(event, vk):
    user_id = event.user_id

    # Если пользователь ещё не скидывал файл отгрузок
    if user_id not in user_data:
        vk.messages.send(
            user_id=user_id,
            message='Сначала отправьте Excel‑файл со списком отгрузок',
            random_id=get_random_id()
        )
        return

    if PVZ_MAPPING is None:
        print('[!!!] Внимание: Не удалось сформировать список адресов')
        vk.messages.send(
            user_id=user_id,
            message=('Не удалось загрузить список адресов ПВЗ, '
                     'адреса не будут добавлены'),
            random_id=get_random_id()
        )

    # Формируем данные для ответа
    try:
        filename = user_data[user_id]
        entered_name = event.text.strip()
        shipment_df = pd.read_excel(filename)
        mask = shipment_df[COURIER_FIO_COLUMN] == entered_name
        filtered_df = shipment_df[mask].copy()
        if filtered_df.empty:
            vk.messages.send(
                user_id=user_id,
                message=f'В файле не нашлось строчек "{entered_name}"',
                random_id=get_random_id()
            )
        else:
            if PVZ_MAPPING:
                if MAIN_PVZ_CODE_COLUMN not in filtered_df.columns:
                    print('[!!!] Внимание: Не найден столбец '
                          f'"{MAIN_PVZ_CODE_COLUMN}" с номерами ПВЗ. '
                          'Возможно, поменялась конфигурация '
                          'Excel‑файла со списком отгрузок')
                    vk.messages.send(
                        user_id=user_id,
                        message=('Не могу добавить адреса: Возможно, '
                                 'поменялась конфигурация Excel‑файла '
                                 'со списком отгрузок'),
                        random_id=get_random_id()
                    )
                # Добавляем адреса
                else:
                    filtered_df[NEW_ADDRESS_COLUMN] = (
                        filtered_df[MAIN_PVZ_CODE_COLUMN]
                        .map(PVZ_MAPPING)
                        .fillna('Адрес не найден')
                    )

            # Формируем отгрузки в текстовый формат
            lines = []
            for _, row in filtered_df.iterrows():
                cells = row['Ячейки']
                target = row['Получатели']
                address = row['Адрес ПВЗ']
                if len(address) > 40:
                    # Обрезаем адрес до 40 последних символов
                    address = f'[ ...{address[-40:]} ]'
                else:
                    address = f'[ {address} ]'
                lines.append(f'{cells}\n{target}\n{address}\n')
            message_text = (f'Готово! Отгрузки для {entered_name} '
                            'отфильтрованы!\n\n'
                            + '\n'.join(lines))

            vk.messages.send(
                user_id=user_id,
                message=message_text,
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
