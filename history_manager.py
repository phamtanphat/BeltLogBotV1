import json
import os
from datetime import datetime

HISTORY_FILE = "report_history.json"


def load_history() -> list:
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_history(history: list):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def add_entry(chat_id: int, username: str, sheet_counts: dict):
    """
    Thêm 1 bản ghi vào lịch sử.
    sheet_counts = {"BC ĐI": 120, "BC ĐẾN": 45, "CHECK IN": 21, "SORTER": 25}
    """
    history = load_history()
    entry = {
        "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "chat_id": chat_id,
        "username": username or "Ẩn danh",
        "sheets": sheet_counts,
        "total": sum(sheet_counts.values()),
    }
    history.insert(0, entry)          # Mới nhất lên đầu
    history = history[:50]            # Giữ tối đa 50 bản ghi
    save_history(history)
    return entry


def format_history(limit: int = 10) -> str:
    history = load_history()
    if not history:
        return "📭 Chưa có báo cáo nào được tạo."

    lines = ["📜 *Lịch Sử Báo Cáo* (10 bản gần nhất)\n"]
    for i, e in enumerate(history[:limit], 1):
        name = f"@{e['username']}" if e.get("username") else "Ẩn danh"
        sheets = e.get("sheets", {})
        detail = " | ".join(f"{k}: {v}" for k, v in sheets.items())
        lines.append(
            f"*{i}.* {e['timestamp']} — {name}\n"
            f"    📊 Tổng {e['total']} dòng  ({detail})\n"
        )
    return "\n".join(lines)
