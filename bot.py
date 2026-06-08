import os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, WebAppInfo, BotCommand
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
MINIAPP_URL = os.getenv("MINIAPP_URL", "")

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    name = user.first_name or "Foydalanuvchi"
    kb = []
    if MINIAPP_URL:
        kb.append([InlineKeyboardButton(
            "🚀 Emerland AI — Ochish",
            web_app=WebAppInfo(url=MINIAPP_URL)
        )])
    kb.append([InlineKeyboardButton("💬 Admin", url="https://t.me/temur_uzb7779")])
    text = (
        f"✨ Salom, *{name}*\\!\n\n"
        "🤖 *Emerland AI* — Professional AI yordamchingiz\\!\n\n"
        "📌 *Nimalар qila olaman:*\n\n"
        "💬 AI Suhbat — Istalgan savolga javob\n"
        "🎨 AI Rasm — So'zdan rasm yaratish\n"
        "🌐 Tarjima — 100\\+ til\n"
        "💻 Kod — Professional dastur yozish\n"
        "📄 PDF — Hujjat tahlil qilish\n"
        "🎤 Ovoz — Audio matnga aylantirish\n"
        "📊 PowerPoint — Prezentatsiya yaratish\n"
        "📝 Word — Hujjat yaratish\n"
        "👤 CV — Professional rezyume\n"
        "📧 Email — Biznes xat yozish\n"
        "🌦 Ob\\-havo — Real vaqt ma'lumot\n"
        "💰 Crypto — Joriy narxlar\n"
        "📰 Yangiliklar — Dunyo yangiliklari\n\n"
        "👇 *Quyidagi tugmani bosib kirish:*"
    )
    await update.message.reply_text(
        text,
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="MarkdownV2"
    )

async def post_init(app):
    await app.bot.set_my_commands([
        BotCommand("start", "🚀 Botni boshlash")
    ])

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start", start))
    print("🚀 Emerland AI Bot ishga tushdi!")
    app.run_polling()

if __name__ == "__main__":
    main()
