"""
history_manager.py
Ghi và đọc lịch sử báo cáo từ Google Sheets (vĩnh viễn, không mất khi Render restart).
"""
import os
import json
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_ID = os.environ.get("GSHEET_ID", "1kiAXF-RVXjrHhVAtBArjESsk7-M9H427g7H6C3bhD2A")
CREDS_FILE = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
SHEET_NAME = "Lịch Sử"


def _get_client():
    creds_json = os.environ.get("GOOGLE_CREDS_JSON")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def _get_sheet():
    client = _get_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    try:
        sheet = spreadsheet.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        # Tự tạo sheet nếu chưa có, thêm header ngay
        sheet = spreadsheet.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)
        sheet.append_row(
            ["Thời Gian", "Username", "Chat ID", "BC ĐI", "BC ĐẾN", "CHECK IN", "SORTER", "Tổng Dòng", "Link Drive"],
            value_input_option="USER_ENTERED"
        )
    return sheet


def add_entry(chat_id: int, username: str, sheet_counts: dict, drive_link: str = ""):
    """Ghi 1 dòng mới vào Google Sheet."""
    try:
        sheet = _get_sheet()
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        row = [
            timestamp,
            username or "Ẩn danh",
            str(chat_id),
            sheet_counts.get("BC ĐI", 0),
            sheet_counts.get("BC ĐẾN", 0),
            sheet_counts.get("CHECK IN", 0),
            sheet_counts.get("SORTER", 0),
            sum(sheet_counts.values()),
            drive_link,
        ]
        sheet.append_row(row, value_input_option="USER_ENTERED")
    except Exception as e:
        print(f"⚠️ Không ghi được lịch sử Google Sheets: {e}")


def format_history(limit: int = 10) -> str:
    """Đọc N dòng cuối từ Google Sheet và format để gửi Telegram."""
    try:
        sheet = _get_sheet()
        all_rows = sheet.get_all_values()
        if len(all_rows) <= 1:
            return "📭 Chưa có báo cáo nào được tạo."

        data_rows = all_rows[1:]        # bỏ header
        recent = data_rows[-limit:]     # lấy N dòng cuối
        recent.reverse()                # mới nhất lên đầu

        lines = ["📜 *Lịch Sử Báo Cáo* (10 bản gần nhất)\n"]
        for i, r in enumerate(recent, 1):
            # r = [timestamp, username, chat_id, bc_di, bc_den, checkin, sorter, total, link]
            timestamp = r[0] if len(r) > 0 else "?"
            uname     = f"@{r[1]}" if (len(r) > 1 and r[1] != "Ẩn danh") else "Ẩn danh"
            di        = r[3] if len(r) > 3 else "0"
            den       = r[4] if len(r) > 4 else "0"
            checkin   = r[5] if len(r) > 5 else "0"
            sorter    = r[6] if len(r) > 6 else "0"
            total     = r[7] if len(r) > 7 else "0"
            link      = r[8] if len(r) > 8 else ""
            link_text = f" [📂 Xem]({link})" if link else ""

            lines.append(
                f"*{i}.* {timestamp} — {uname}{link_text}\n"
                f"    📊 Tổng {total} dòng  (ĐI: {di} | ĐẾN: {den} | Check\\-In: {checkin} | Sorter: {sorter})\n"
            )

        return "\n".join(lines)

    except Exception as e:
        return f"❌ Không thể đọc lịch sử: `{e}`"
