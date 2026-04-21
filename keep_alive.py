from flask import Flask
from threading import Thread

app = Flask(__name__)

@app.route('/')
def home():
    return "🚀 BHS Maintenance Bot is running 24/7!"

def run():
    # Chạy Web Server ở port 8080 (hoặc port của Cloud tự gán)
    import os
    port = int(os.environ.get("PORT", 8080))
    app.run(host='0.0.0.0', port=port)

def keep_alive():
    server = Thread(target=run)
    server.start()
