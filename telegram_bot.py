import telebot
import os
import generate_report
from datetime import datetime
from keep_alive import keep_alive

# Đọc token từ biến môi trường (an toàn, không lộ token trên GitHub)
TOKEN = os.environ.get("BOT_TOKEN", "8730792867:AAFyKUfJvnoF0uIZM5xGV3O9JtJcq7WdbUc")
bot = telebot.TeleBot(TOKEN)

# Lưu trạng thái từng người dùng (theo Chat ID)
# Cấu trúc: { chat_id: ["path_file_eq", "path_file_mds"] }
user_sessions = {}

# Thư mục lưu file tạm
DOWNLOAD_DIR = "downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# ─────────────────────────────────────────────
# LỆNH /start  →  Hướng dẫn sử dụng
# ─────────────────────────────────────────────
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(
        message,
        "👋 *Chào mừng đến với Bot Tổng Hợp Báo Cáo Bảo Trì BHS!*\n\n"
        "📌 *Cách dùng:*\n"
        "1️⃣ Gửi file *Sổ Lý Lịch Thiết Bị* (source\\_eq\\_history.xlsx)\n"
        "2️⃣ Gửi file *Nhật Ký MDS* (source\\_mds\\_maintenance.xlsx)\n"
        "3️⃣ Bot sẽ tự động xử lý và gửi trả file *Báo Cáo 4 Bảng*.\n\n"
        "⚠️ Gửi *đúng thứ tự* và *đúng 2 file* để nhận kết quả.",
        parse_mode='Markdown'
    )

# ─────────────────────────────────────────────
# LỆNH /reset  →  Xóa session nếu gửi nhầm file
# ─────────────────────────────────────────────
@bot.message_handler(commands=['reset'])
def reset_session(message):
    chat_id = message.chat.id
    # Dọn file cũ nếu có
    if chat_id in user_sessions:
        for f in user_sessions[chat_id]:
            if os.path.exists(f): os.remove(f)
        user_sessions[chat_id] = []
    bot.reply_to(message, "♻️ Đã reset! Bạn có thể gửi lại từ đầu.")

# ─────────────────────────────────────────────
# NHẬN FILE EXCEL
# ─────────────────────────────────────────────
@bot.message_handler(content_types=['document'])
def handle_docs(message):
    chat_id = message.chat.id
    file_name = message.document.file_name

    # Chỉ nhận file Excel
    if not (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
        bot.reply_to(message, "⚠️ Vui lòng chỉ gửi file Excel (.xlsx hoặc .xls)\nGõ /reset nếu cần gửi lại.")
        return

    # Khởi tạo session nếu người dùng mới
    if chat_id not in user_sessions:
        user_sessions[chat_id] = []

    # Không nhận quá 2 file, nhắc reset
    if len(user_sessions[chat_id]) >= 2:
        bot.reply_to(message, "⚠️ Bạn đã gửi đủ 2 file rồi!\nNếu muốn bắt đầu lại, hãy gõ /reset")
        return

    try:
        # Tải file về Server
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        # Đặt tên file theo Chat ID + thứ tự để tránh xung đột
        slot = len(user_sessions[chat_id]) + 1
        safe_name = f"user{chat_id}_file{slot}_{file_name}"
        save_path = os.path.join(DOWNLOAD_DIR, safe_name)

        with open(save_path, 'wb') as f:
            f.write(downloaded_file)

        user_sessions[chat_id].append(save_path)

        # ── Nhận file 1 → Chờ file 2 ──
        if len(user_sessions[chat_id]) == 1:
            bot.reply_to(
                message,
                f"✅ Đã nhận file 1: `{file_name}`\n"
                "📎 Hãy gửi tiếp *file thứ 2* để tôi bắt đầu xử lý.",
                parse_mode='Markdown'
            )

        # ── Nhận file 2 → Kích hoạt xử lý ──
        elif len(user_sessions[chat_id]) == 2:
            processing_msg = bot.reply_to(
                message,
                f"✅ Đã nhận file 2: `{file_name}`\n\n"
                "⚙️ *Đang chạy AI phân loại dữ liệu...*\nVui lòng chờ vài giây ☕",
                parse_mode='Markdown'
            )

            file_1 = user_sessions[chat_id][0]
            file_2 = user_sessions[chat_id][1]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            out_path = os.path.join(DOWNLOAD_DIR, f"Report_{chat_id}_{timestamp}.xlsx")

            try:
                success = generate_report.process_and_merge_excel(
                    file_equipment=file_1,
                    file_mds=file_2,
                    output_excel=out_path,
                    verbose=False
                )

                if success:
                    with open(out_path, 'rb') as doc:
                        bot.send_document(
                            chat_id, doc,
                            caption=(
                                "🎉 *THÀNH CÔNG!* Báo cáo đã được tổng hợp:\n\n"
                                "📋 BẢNG 1 – Băng Chuyền ĐI\n"
                                "📋 BẢNG 2 – Băng Chuyền ĐẾN\n"
                                "📋 BẢNG 3 – Check-In\n"
                                "📋 BẢNG 4 – Sorter"
                            ),
                            parse_mode='Markdown'
                        )
                else:
                    bot.reply_to(message, "❌ Dữ liệu trong file trống hoặc không đúng định dạng.\nGõ /reset để thử lại.")

            except Exception as e:
                bot.reply_to(message, f"❌ Lỗi xử lý: `{e}`\nGõ /reset để thử lại.", parse_mode='Markdown')

            finally:
                # Dọn rác ngay lập tức
                for f in user_sessions.get(chat_id, []):
                    if os.path.exists(f): os.remove(f)
                if os.path.exists(out_path): os.remove(out_path)
                user_sessions[chat_id] = []

    except Exception as e:
        bot.reply_to(message, f"❌ Lỗi tải file từ Telegram: `{e}`", parse_mode='Markdown')
        user_sessions[chat_id] = []

# ─────────────────────────────────────────────
# KHỞI ĐỘNG BOT
# ─────────────────────────────────────────────
if __name__ == '__main__':
    keep_alive()  # Bật Web Server để Render.com / UptimeRobot giữ chương trình sống
    print("🤖 Bot Telegram BHS đang chạy 24/7...")
    bot.infinity_polling(timeout=10, long_polling_timeout=5)
