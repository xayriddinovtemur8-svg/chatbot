import os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, WebAppInfo
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
MINIAPP_URL    = os.getenv("MINIAPP_URL", "")
BOT_NAME       = "Emerland AI"

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    kb = [[InlineKeyboardButton(
        "🚀 Emerland AI ni Ochish",
        web_app=WebAppInfo(url=MINIAPP_URL)
    )]]
    await update.message.reply_photo(
        photo="https://i.imgur.com/YOUR_IMAGE.jpg",  # Bot rasmi URL
        caption=(
            f"✨ *{BOT_NAME}*\n\n"
            "Men qila oladigan ishlarim:\n\n"
            "🤖 AI Suhbat — Istalgan savolga javob\n"
            "🎨 AI Rasm — So'zdan rasm yaratish\n"
            "🌐 Tarjima — 100+ til\n"
            "💻 Kod — Dastur yozish\n"
            "📄 PDF — Hujjat tahlil\n"
            "🎤 Ovoz — Audio matnga\n"
            "📊 PowerPoint — Prezentatsiya\n"
            "📝 Word — Hujjat yaratish\n"
            "👤 CV — Professional rezyume\n"
            "📧 Email — Biznes xat\n"
            "🌦 Ob-havo — Real vaqt\n"
            "💰 Crypto — Narxlar\n\n"
            "👇 Ochish uchun tugmani bosing:"
        ),
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )

async def post_init(app):
    from telegram import BotCommand
    await app.bot.set_my_commands([
        BotCommand("start", "🚀 Botni boshlash")
    ])

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start", start))
    print("🚀 Bot ishga tushdi!")
    app.run_polling()

if __name__ == "__main__":
    main()
