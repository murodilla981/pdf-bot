import os
import threading
import telebot
from flask import Flask

TOKEN = os.environ.get("BOT_TOKEN")
bot = telebot.TeleBot(TOKEN)

app = Flask(__name__)


@app.route("/")
def home():
  return "Bot ishlayapti!"


def run_web():
  port = int(os.environ.get("PORT", 10000))
  app.run(host="0.0.0.0", port=port)


threading.Thread(target=run_web).start()


@bot.message_handler(commands=["start"])
def send_welcome(message):
  bot.reply_to(
      message,
      "Salom! Men PDF botman va slaydmaker botga aylanishga tayyorman! 😊",
  )


@bot.message_handler(func=lambda message: True)
def echo_all(message):
  bot.reply_to(
      message, "Xabaringiz qabul qilindi! Tez orada yangi funksiyalar qo'shiladi."
  )


if __name__ == "__main__":
  bot.infinity_polling()
