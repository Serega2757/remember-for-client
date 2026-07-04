import os
import time
import json
import requests
import gspread
from datetime import datetime, date, timedelta
from google.oauth2.service_account import Credentials
from zoneinfo import ZoneInfo

GID = 782899969

MAX_MESSAGES = 4
DELAY_BETWEEN_MESSAGES = 20
MAX_ATTEMPTS = 3
DELAY_BETWEEN_ATTEMPTS = 13

DAYS_AFTER_ORDER = 4
DAYS_AFTER_MESSAGE = 10
DELETE_AFTER_DAYS = 180

API_URL = os.environ["ECHAT_API_URL"]
API_KEY = os.environ["ECHAT_API_KEY"]
CHANNEL_NUMBER = os.environ["ECHAT_CHANNEL_NUMBER"]

SPREADSHEET_ID = os.environ["GOOGLE_SPREADSHEET_ID"]
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]


def parse_date(value):
    if not value:
        return None

    if isinstance(value, date):
        return value

    value = str(value).strip()
    formats = [
        "%d.%m.%y",
        "%d.%m.%Y",
        "%d.%m.%y %H:%M:%S",
        "%d.%m.%Y %H:%M:%S",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass

    return None


def normalize_phone(value):
    if not value:
        return ""

    phone = "".join(ch for ch in str(value) if ch.isdigit())

    if not phone:
        return ""

    if phone.startswith("0"):
        phone = "38" + phone

    if not phone.startswith("380") and len(phone) == 9:
        phone = "380" + phone

    return phone


def first_name(value):
    if not value:
        return ""

    return str(value).strip().split()[0] if str(value).strip() else ""


def build_message(name, product_name):
    if name:
        return (
            f"Вітаємо, {name} 👋\n\n"
            f"Товар, який Ви очікували, знову є в наявності:\n\n"
            f"{product_name}\n\n"
            f"Можемо швидко оформити замовлення."
        )

    return (
        "Вітаємо 👋\n\n"
        "Товар, який Ви очікували, знову є в наявності:\n\n"
        f"{product_name}\n\n"
        "Можемо швидко оформити замовлення."
    )
def should_run_now(spreadsheet):
    kyiv_now = datetime.now(ZoneInfo("Europe/Kyiv"))

    slots = [
        ("10:30", 10, 30),
        ("13:10", 13, 10),
        ("17:45", 17, 45),
    ]

    matched_slot = None

    for slot_name, hour, minute in slots:
        target = kyiv_now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        diff_minutes = (kyiv_now - target).total_seconds() / 60

        if 0 <= diff_minutes <= 90:
            matched_slot = slot_name
            break

    if not matched_slot:
        print(f"Зараз {kyiv_now.strftime('%H:%M')} Київ | Не час запуску")
        return False

    log_sheet_name = "BotRunLog"

    try:
        log_ws = spreadsheet.worksheet(log_sheet_name)
    except gspread.WorksheetNotFound:
        log_ws = spreadsheet.add_worksheet(title=log_sheet_name, rows=1000, cols=3)
        log_ws.append_row(["date", "slot", "datetime"])

    today_text = kyiv_now.strftime("%Y-%m-%d")
    log_rows = log_ws.get_all_values()

    for row in log_rows[1:]:
        if len(row) >= 2 and row[0] == today_text and row[1] == matched_slot:
            print(f"Слот {matched_slot} за {today_text} вже виконувався. Пропуск.")
            return False

    log_ws.append_row([
        today_text,
        matched_slot,
        kyiv_now.strftime("%Y-%m-%d %H:%M:%S")
    ])

    print(f"Дозволено запуск | Слот {matched_slot} | Київ {kyiv_now.strftime('%H:%M')}")
    return True

def send_viber(phone, text, row_number):
    payload = {
        "number": CHANNEL_NUMBER,
        "contact": {
            "number": phone
        },
        "message": {
            "id": f"msg_{int(time.time() * 1000)}_{row_number}",
            "text": text
        }
    }

    headers = {
        "Api-Key": API_KEY,
        "Content-Type": "application/json"
    }

    last_error = ""

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            print(f"Спроба {attempt}/{MAX_ATTEMPTS} | Рядок {row_number} | {phone}")

            response = requests.post(
                API_URL,
                headers=headers,
                json=payload,
                timeout=30
            )

            if 200 <= response.status_code < 300:
                print(f"API OK | Рядок {row_number} | Код {response.status_code}")
                return True

            last_error = f"API код {response.status_code} | {response.text}"
            print(f"API ПОМИЛКА | Рядок {row_number} | {last_error}")

        except Exception as e:
            last_error = str(e)
            print(f"КРИТИЧНА ПОМИЛКА | Рядок {row_number} | {last_error}")

        if attempt < MAX_ATTEMPTS:
            time.sleep(DELAY_BETWEEN_ATTEMPTS)

    print(f"НЕ ВІДПРАВЛЕНО | Рядок {row_number} | {last_error}")
    return False


def get_sheet():
    service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes
    )

    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    for ws in spreadsheet.worksheets():
        if ws.id == GID:
            return spreadsheet, ws

    raise RuntimeError(f"Лист з gid={GID} не знайдено")


def main():
    print("=== Старт БотПотрібніТовариV2 Python ===")

    spreadsheet, ws = get_sheet()

    if not should_run_now(spreadsheet):
        return
    rows = ws.get_all_values()

    if len(rows) < 2:
        print("Немає рядків для обробки")
        return

    today = date.today()
    now_text = datetime.now().strftime("%d.%m.%y %H:%M:%S")
    today_text = today.strftime("%d.%m.%y")

    # Беремо тільки рядки до першої порожньої E
    active_rows = []

    for index, row in enumerate(rows[1:], start=2):
        phone_raw = row[4].strip() if len(row) > 4 else ""

        if not phone_raw:
            print(f"Рядок {index} | Перша порожня E. Нижче не перевіряємо.")
            break

        active_rows.append((index, row))

    print(f"Активних рядків до першої порожньої E: {len(active_rows)}")

    # 1. Видалення старих рядків: J > 0 і F старіша ніж 180 днів
    rows_to_delete = []

    for row_number, row in active_rows:
        last_order_date = parse_date(row[5] if len(row) > 5 else "")
        message_count = int(row[9]) if len(row) > 9 and str(row[9]).strip().isdigit() else 0

        if message_count > 0 and last_order_date:
            if today - last_order_date > timedelta(days=DELETE_AFTER_DAYS):
                rows_to_delete.append(row_number)

    for row_number in reversed(rows_to_delete):
        ws.delete_rows(row_number)
        print(f"Видалено рядок {row_number} | J > 0 | F старіша ніж 180 днів")

    if rows_to_delete:
        rows = ws.get_all_values()
        active_rows = []

        for index, row in enumerate(rows[1:], start=2):
            phone_raw = row[4].strip() if len(row) > 4 else ""

            if not phone_raw:
                break

            active_rows.append((index, row))

    # 2. Для всіх активних рядків: K дата створення, J = 0 якщо порожньо
    j_updates = []
    k_updates = []

    for row_number, row in active_rows:
        phone_raw = row[4].strip() if len(row) > 4 else ""
        j_value = row[9].strip() if len(row) > 9 else ""
        k_value = row[10].strip() if len(row) > 10 else ""

        if phone_raw and not k_value:
            k_updates.append({
                "range": f"K{row_number}",
                "values": [[today_text]]
            })

        if j_value == "":
            j_updates.append({
                "range": f"J{row_number}",
                "values": [[0]]
            })

    if j_updates:
        ws.batch_update(j_updates)
        print(f"Оновлено J: {len(j_updates)}")

    if k_updates:
        ws.batch_update(k_updates)
        print(f"Оновлено K: {len(k_updates)}")

    # Перечитуємо після оновлення J/K
    rows = ws.get_all_values()
    active_rows = []

    for index, row in enumerate(rows[1:], start=2):
        phone_raw = row[4].strip() if len(row) > 4 else ""

        if not phone_raw:
            break

        active_rows.append((index, row))

    # 3. Відправка максимум 4 повідомлень
    sent_count = 0

    for row_number, row in active_rows:
        if sent_count >= MAX_MESSAGES:
            print(f"Досягнуто ліміт за запуск: {MAX_MESSAGES}")
            break

        product_name = row[1].strip() if len(row) > 1 else ""
        current_status = row[2].strip() if len(row) > 2 else ""
        client_name = first_name(row[3] if len(row) > 3 else "")
        phone = normalize_phone(row[4] if len(row) > 4 else "")

        last_order_date = parse_date(row[5] if len(row) > 5 else "")
        last_sent_date = parse_date(row[7] if len(row) > 7 else "")
        previous_status = row[8].strip() if len(row) > 8 else ""
        current_message_count = int(row[9]) if len(row) > 9 and row[9].strip().isdigit() else 0

        if not product_name:
            print(f"Рядок {row_number} | Пропуск: немає назви товару")
            continue

        if current_status == "-":
            if previous_status != "-":
                ws.update(f"I{row_number}", [["-"]])
                print(f"Рядок {row_number} | Товар відсутній | I => '-'")
            continue

        if current_status != "+":
            print(f"Рядок {row_number} | Пропуск: C не '+'")
            continue

        if previous_status not in ("", "-"):
            print(f"Рядок {row_number} | Пропуск: не нова поява товару | I={previous_status}")
            continue

        if last_order_date and today - last_order_date <= timedelta(days=DAYS_AFTER_ORDER):
            print(f"Рядок {row_number} | Пропуск: замовлення було менше ніж 4 дні тому")
            continue

        if last_sent_date and today - last_sent_date <= timedelta(days=DAYS_AFTER_MESSAGE):
            print(f"Рядок {row_number} | Пропуск: повідомлення було менше ніж 10 днів тому")
            continue

        message = build_message(client_name, product_name)

        if send_viber(phone, message, row_number):
            sent_count += 1

            ws.batch_update([
                {
                    "range": f"H{row_number}",
                    "values": [[now_text]]
                },
                {
                    "range": f"I{row_number}",
                    "values": [["+"]]
                },
                {
                    "range": f"J{row_number}",
                    "values": [[current_message_count + 1]]
                }
            ])

            print(f"ВІДПРАВЛЕНО {sent_count}/{MAX_MESSAGES} | Рядок {row_number} | {phone}")

            if sent_count < MAX_MESSAGES:
                time.sleep(DELAY_BETWEEN_MESSAGES)

    print(f"=== Завершено | Надіслано: {sent_count} ===")


if __name__ == "__main__":
    main()
