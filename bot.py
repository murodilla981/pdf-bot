import os
import threading
import requests
from io import BytesIO
import telebot
from flask import Flask
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import google.generativeai as genai

# Tokenlar
TELEGRAM_TOKEN = os.environ.get("BOT_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)
    ai_model = genai.GenerativeModel("gemini-1.5-flash")
else:
    ai_model = None

bot = telebot.TeleBot(TELEGRAM_TOKEN)
app = Flask(__name__)

@app.route('/')
def home():
    return "AI & 3D Slide Bot with Rich Content is Running!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

threading.Thread(target=run_web, daemon=True).start()

def get_main_menu():
    markup = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btn_slide = telebot.types.KeyboardButton("📊 Slayd yaratish")
    btn_help = telebot.types.KeyboardButton("💡 Yordam")
    markup.add(btn_slide, btn_help)
    return markup

# Internetdan mavzuga mos rasm topish
def get_topic_image(query):
    try:
        url = f"https://source.unsplash.com/800x600/?{requests.utils.quote(query)}"
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            return BytesIO(response.content)
    except Exception:
        pass
    return None

def create_presentation(topic, details):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # 1. Muqova slaydi (3D Dark Style)
    slide = prs.slides.add_slide(blank_layout)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(15, 23, 42)
    bg.line.fill.background()

    # 3D Kub shakli
    cube_3d = slide.shapes.add_shape(MSO_SHAPE.CUBE, Inches(1.0), Inches(2.2), Inches(3.0), Inches(3.0))
    cube_3d.fill.solid()
    cube_3d.fill.fore_color.rgb = RGBColor(56, 189, 248)
    cube_3d.line.color.rgb = RGBColor(255, 255, 255)

    # Sarlavha kartasi
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.3), Inches(1.5), Inches(8.0), Inches(4.5))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(30, 41, 59)
    card.line.color.rgb = RGBColor(56, 189, 248)
    card.line.width = Pt(2)

    tx_box = slide.shapes.add_textbox(Inches(4.6), Inches(2.2), Inches(7.4), Inches(3))
    tf = tx_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = topic.upper()
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(248, 250, 252)
    p.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = "✨ 3D Vizualizatsiya, Rasmlar va To'liq Ma'lumotlar"
    p2.font.size = Pt(15)
    p2.font.color.rgb = RGBColor(56, 189, 248)
    p2.alignment = PP_ALIGN.CENTER

    # 2. Ma'lumotlar va Rasmli Slayd
    slide2 = prs.slides.add_slide(blank_layout)
    bg2 = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = RGBColor(248, 250, 252)
    bg2.line.fill.background()

    # Slayd ichiga mavzuga oid rasm qo'shish
    img_stream = get_topic_image(topic)
    if img_stream:
        try:
            slide2.shapes.add_picture(img_stream, Inches(0.8), Inches(1.5), width=Inches(4.5), height=Inches(5.0))
        except Exception:
            pass

    # Slayd ichiga to'liq ma'lumotlarni joylash uchun ikkita keng kartochka
    parts = details.split("\n\n")
    part1 = parts[0] if len(parts) > 0 else details[:400]
    part2 = parts[1] if len(parts) > 1 else details[400:800]

    # 1-Ma'lumotlar kartasi
    box1 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), Inches(1.5), Inches(7.0), Inches(2.4))
    box1.fill.solid()
    box1.fill.fore_color.rgb = RGBColor(255, 255, 255)
    box1.line.color.rgb = RGBColor(203, 213, 225)
    box1.line.width = Pt(1.5)

    tf1 = box1.text_frame
    tf1.word_wrap = True
    p_t1 = tf1.paragraphs[0]
    p_t1.text = "📌 Asosiy Tushunchalar"
    p_t1.font.size = Pt(15)
    p_t1.font.bold = True
    p_t1.font.color.rgb = RGBColor(14, 116, 144)

    p_d1 = tf1.add_paragraph()
    p_d1.text = "\n" + part1[:300] + "..."
    p_d1.font.size = Pt(11)
    p_d1.font.color.rgb = RGBColor(71, 85, 105)

    # 2-Ma'lumotlar kartasi
    box2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), Inches(4.1), Inches(7.0), Inches(2.4))
    box2.fill.solid()
    box2.fill.fore_color.rgb = RGBColor(255, 255, 255)
    box2.line.color.rgb = RGBColor(203, 213, 225)
    box2.line.width = Pt(1.5)

    tf2 = box2.text_frame
    tf2.word_wrap = True
    p_t2 = tf2.paragraphs[0]
    p_t2.text = "🔍 Batafsil Tahlil va Xulosalar"
    p_t2.font.size = Pt(15)
    p_t2.font.bold = True
    p_t2.font.color.rgb = RGBColor(14, 116, 144)

    p_d2 = tf2.add_paragraph()
    p_d2.text = "\n" + part2[:300] + "..."
    p_d2.font.size = Pt(11)
    p_d2.font.color.rgb = RGBColor(71, 85, 105)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(
        message,
        "🖥 **Asosiy menyudasiz!**\n\n"
        "• Menga istalgan savolingizni yozing — **AI barcha sun'iy intellekt ma'lumotlari asosida to'g'ri va aniq javob beradi**.\n"
        "• Yuborgan mavzuingiz bo'yicha **ichida rasmlari, 3D shakllari va batafsil ma'lumotlari bo'lgan slayd** tayyorlab beraman.",
        reply_markup=get_main_menu(),
        parse_mode="Markdown"
    )

@bot.message_handler(func=lambda message: True)
def handle_messages(message):
    user_text = message.text

    if user_text == "📊 Slayd yaratish":
        bot.reply_to(message, "Marhamat, qaysi mavzuda slayd yaratish kerak? Mavzu nomini yuboring:", reply_markup=get_main_menu())
        return
    elif user_text == "💡 Yordam":
        bot.reply_to(message, "💡 Istalgan savol yoki mavzuni yozing. Bot AI orqali to'liq javob qaytaradi va ichida ma'lumotlari bor rasmli 3D slayd taqdim etadi.", reply_markup=get_main_menu())
        return

    if ai_model:
        try:
            bot.reply_to(message, f"⏳ '{user_text}' bo'yicha AI ma'lumotlari qidirilmoqda va slayd tayyorlanmoqda...", reply_markup=get_main_menu())

            # AI orqali aniq va to'g'ri ma'lumot olish
            prompt = f"'{user_text}' haqida o'zbek tilida ilmiy, aniq, to'g'ri va batafsil ma'lumot ber."
            response = ai_model.generate_content(prompt)
            ai_text = response.text

            # 1. Matnli to'g'ri javobni yuborish
            bot.send_message(message.chat.id, f"📖 **AI Javobi va Ma'lumot:**\n\n{ai_text[:4000]}", parse_mode="Markdown", reply_markup=get_main_menu())

            # 2. Ichida ma'lumotlari, rasmi va 3D shakli bo'lgan slaydni yuborish
            ppt_file = create_presentation(user_text, ai_text)
            bot.send_document(
                message.chat.id,
                ppt_file,
                visible_file_name=f"{user_text[:15]}_slayd.pptx",
                caption=f"✅ '{user_text}' bo'yicha rasmli, ma'lumotli 3D slayd tayyor!"
            )
        except Exception as e:
            bot.reply_to(message, f"Xatolik yuz berdi: {str(e)}", reply_markup=get_main_menu())
    else:
        bot.reply_to(message, "AI sozlanmagan.", reply_markup=get_main_menu())

if __name__ == '__main__':
    bot.infinity_polling()
