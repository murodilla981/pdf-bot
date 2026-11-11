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
import google.generativeai as genai

# Tokenlar
TELEGRAM_TOKEN = os.environ.get("BOT_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

# Gemini AI ni sozlash
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)
    ai_model = genai.GenerativeModel("gemini-1.5-flash")
else:
    ai_model = None

bot = telebot.TeleBot(TELEGRAM_TOKEN)
app = Flask(__name__)

@app.route('/')
def home():
    return "Menu & 3D Slide Bot is Running!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

threading.Thread(target=run_web, daemon=True).start()

# Asosiy menyu tugmalarini yaratish
def get_main_menu():
    markup = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btn_slide = telebot.types.KeyboardButton("📊 Slayd yaratish")
    btn_help = telebot.types.KeyboardButton("💡 Qo'llanma / Yordam")
    btn_about = telebot.types.KeyboardButton("ℹ️ Bot haqida")
    markup.add(btn_slide, btn_help, btn_about)
    return markup

def create_presentation(topic):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    sections = [
        ("01. Kirish va Asosiy tushunchalar", f"{topic} bo'yicha umumiy nazariy ma'lumotlar va fundamental asoslar."),
        ("02. Tuzilishi va Xususiyatlari", f"{topic} ning muhim tarkibiy qismlari va tahliliy ko'rsatkichlari."),
        ("03. Amaliy ahamiyati va Xulosalar", f"Zamonaviy dunyoda {topic} ning o'rni va kelgusidagi istiqbollari.")
    ]

    if ai_model:
        try:
            prompt = f"'{topic}' mavzusi uchun taqdimotga 3 ta qisqa bo'lim (Sarlavha va 1-2 gaplik matn) tuzib ber. Format: Sarlavha | Matn"
            res = ai_model.generate_content(prompt)
            lines = res.text.strip().split("\n")
            extracted = []
            for line in lines:
                if "|" in line:
                    parts = line.split("|")
                    extracted.append((parts[0].strip(), parts[1].strip()))
            if len(extracted) >= 3:
                sections = extracted[:3]
        except Exception:
            pass

    # 1. Muqova slaydi
    slide = prs.slides.add_slide(blank_layout)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(15, 23, 42)
    bg.line.fill.background()

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.5), Inches(10.33), Inches(4.5))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(30, 41, 59)
    card.line.color.rgb = RGBColor(56, 189, 248)
    card.line.width = Pt(2)

    tx_box = slide.shapes.add_textbox(Inches(2), Inches(2.2), Inches(9.33), Inches(2))
    tf = tx_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = topic.upper()
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(248, 250, 252)
    p.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = "3D Vizualizatsiya & Avtomatik Taqdimot"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(56, 189, 248)
    p2.alignment = PP_ALIGN.CENTER

    # 2. Tarkibiy slayd
    slide2 = prs.slides.add_slide(blank_layout)
    bg2 = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = RGBColor(248, 250, 252)
    bg2.line.fill.background()

    left_pos = Inches(1.0)
    for title, desc in sections:
        box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_pos, Inches(2.0), Inches(3.5), Inches(4.5))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        box.line.color.rgb = RGBColor(203, 213, 225)
        box.line.width = Pt(1.5)

        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = RGBColor(14, 116, 144)

        p_desc = tf.add_paragraph()
        p_desc.text = "\n" + desc
        p_desc.font.size = Pt(13)
        p_desc.font.color.rgb = RGBColor(71, 85, 105)

        left_pos += Inches(3.9)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(
        message,
        "🖥 **Asosiy menyudasiz!**\n\n"
        "📝 Botdan foydalanish uchun pastdagi tugmalardan foydalaning yoki istalgan mavzuni yozib yuboring.",
        reply_markup=get_main_menu(),
        parse_mode="Markdown"
    )

@bot.message_handler(func=lambda message: True)
def handle_messages(message):
    user_text = message.text

    if user_text == "📊 Slayd yaratish":
        bot.reply_to(message, "Marhamat, qaysi mavzuda slayd yaratish kerak? Mavzu nomini yuboring (masalan: *Yurak anatomiyasi*, *Kvant fizikasi*):", parse_mode="Markdown")
        return
    elif user_text == "💡 Qo'llanma / Yordam":
        bot.reply_to(message, "💡 **Qo'llanma:**\n\nIstalgan mavzuni yozing yoki tugma orqali murojaat qiling. Bot sizga ham batafsil ma'lumot beradi, ham 3D elementli `.pptx` slayd tayyorlab beradi.")
        return
    elif user_text == "ℹ️ Bot haqida":
        bot.reply_to(message, "🤖 Bu bot Sun'iy Intellekt va 3D formatdagi taqdimotlar yaratishga mo'ljallangan universal yordamchidir.")
        return

    # Agar foydalanuvchi mavzu yuborsa
    topic = user_text
    bot.reply_to(message, f"⏳ '{topic}' bo'yicha ma'lumot tayyorlanmoqda va 3D slayd yasalmoqda...", reply_markup=get_main_menu())

    try:
        ppt_file = create_presentation(topic)
        bot.send_document(
            message.chat.id,
            ppt_file,
            visible_file_name=f"{topic}_slayd.pptx",
            caption=f"✅ '{topic}' bo'yicha 3D formatdagi taqdimot tayyor!"
        )

        if ai_model:
            prompt = f"'{topic}' mavzusi haqida o'zbek tilida qiziqarli, aniq va foydali ma'lumot ber."
            response = ai_model.generate_content(prompt)
            bot.send_message(message.chat.id, f"📖 **Mavzu bo'yicha ma'lumot:**\n\n{response.text}", parse_mode="Markdown", reply_markup=get_main_menu())

    except Exception as e:
        bot.reply_to(message, f"Xatolik yuz berdi: {str(e)}", reply_markup=get_main_menu())

if __name__ == '__main__':
    bot.infinity_polling()
