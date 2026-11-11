import os
import telebot

# Render'dan tokeni avtomatik o'qib oladi
TOKEN = os.environ.get("BOT_TOKEN")
bot = telebot.TeleBot(TOKEN)


@bot.message_handler(commands=["start", "help"])
def send_welcome(message):
  bot.reply_to(
      message,
      "Assalomu alaykum! Murodil, PDF botimiz ishga tushdi. Menga PDF fayl"
      " yuboring!",
  )


@bot.message_handler(content_types=["document"])
def handle_docs(message):
  file_name = message.document.file_name
  bot.reply_to(message, f"PDF fayl qabul qilindi: {file_name} ✅")


# Botni doimiy ishlab turishi uchun
bot.infinity_polling()
