import os
import threading
from io import BytesIO
import telebot
from flask import Flask
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

TOKEN = os.environ.get("BOT_TOKEN")
bot = telebot.TeleBot(TOKEN)

app = Flask(__name__)


@app.route("/")
def home():
  return "Medical & 3D Slide Generator Bot Running!"


def run_web():
  port = int(os.environ.get("PORT", 10000))
  app.run(host="0.0.0.0", port=port)


threading.Thread(target=run_web, daemon=True).start()


def create_presentation(topic):
  prs = Presentation()
  prs.slide_width = Inches(13.33)
  prs.slide_height = Inches(7.5)

  # 1. Muqova slaydi (Dark 3D Style)
  blank_layout = prs.slide_layouts[6]
  slide = prs.slides.add_slide(blank_layout)

  # Fon
  bg = slide.shapes.add_shape(
      MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
  )
  bg.fill.solid()
  bg.fill.fore_color.rgb = RGBColor(15, 23, 42)  # Slate dark
  bg.line.fill.background()

  # 3D Effect Card Shape
  card = slide.shapes.add_shape(
      MSO_SHAPE.ROUNDED_RECTANGLE,
      Inches(1.5),
      Inches(1.5),
      Inches(10.33),
      Inches(4.5),
  )
  card.fill.solid()
  card.fill.fore_color.rgb = RGBColor(30, 41, 59)
  card.line.color.rgb = RGBColor(56, 189, 248)
  card.line.width = Pt(2)

  # Sarlavha
  tx_box = slide.shapes.add_textbox(Inches(2), Inches(2.2), Inches(9.33), Inches(2))
  tf = tx_box.text_frame
  tf.word_wrap = True
  p = tf.paragraphs[0]
  p.text = topic.upper()
  p.font.size = Pt(40)
  p.font.bold = True
  p.font.color.rgb = RGBColor(248, 250, 252)
  p.alignment = PP_ALIGN.CENTER

  p2 = tf.add_paragraph()
  p2.text = "Tibbiy Taqdimot & 3D Vizualizatsiya"
  p2.font.size = Pt(20)
  p2.font.color.rgb = RGBColor(56, 189, 248)
  p2.alignment = PP_ALIGN.CENTER

  # 2. Tarkibiy slayd
  slide2 = prs.slides.add_slide(blank_layout)
  bg2 = slide2.shapes.add_shape(
      MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
  )
  bg2.fill.solid()
  bg2.fill.fore_color.rgb = RGBColor(248, 250, 252)
  bg2.line.fill.background()

  # 3 ta Kartochka (3D Stil)
  card_data = [
      (
          "01. Anamnez va Kirish",
          "Mavzuga oid nazariy asoslar hamda zamonaviy tibbiy ma'lumotlar"
          " tahlili.",
      ),
      (
          "02. Patofiziologiya",
          (
              "Organizmda kechadigan jarayonlar va 3D vizual struktura"
              " namoyishi."
          ),
      ),
      (
          "03. Davolash va Tavsiyalar",
          (
              "Klinik protokol bo'yicha beriladigan zamonaviy dori va"
              " muolajalar."
          ),
      ),
  ]

  left_pos = Inches(1.0)
  for title, desc in card_data:
    box = slide2.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left_pos,
        Inches(2.2),
        Inches(3.5),
        Inches(4.0),
    )
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    box.line.color.rgb = RGBColor(226, 232, 240)
    box.line.width = Pt(1.5)

    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(14, 116, 144)

    p_desc = tf.add_paragraph()
    p_desc.text = "\n" + desc
    p_desc.font.size = Pt(14)
    p_desc.font.color.rgb = RGBColor(71, 85, 105)

    left_pos += Inches(3.9)

  stream = BytesIO()
  prs.save(stream)
  stream.seek(0)
  return stream


@bot.message_handler(commands=["start"])
def send_welcome(message):
  bot.reply_to(
      message,
      "Salom! Men Tibbiy ma'lumot beruvchi va 3D modelli slaydlar tayyorlaydigan"
      " botman.\n\nPrezentatsiya qilish uchun istalgan mavzuni yozing!",
  )


@bot.message_handler(func=lambda message: True)
def handle_text(message):
  topic = message.text
  bot.reply_to(message, f"⏳ '{topic}' mavzusida 3D vizual slayd tayyorlanmoqda...")

  try:
    ppt_file = create_presentation(topic)
    bot.send_document(
        message.chat.id,
        ppt_file,
        visible_file_name=f"{topic}_presentation.pptx",
        caption=f"✅ '{topic}' mavzusidagi prezentatsiyangiz tayyor!",
    )
  except Exception as e:
    bot.reply_to(message, f"Xatolik yuz berdi: {str(e)}")


if __name__ == "__main__":
  bot.infinity_polling()
