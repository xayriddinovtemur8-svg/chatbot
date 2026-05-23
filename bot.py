import os
import io
import requests
import fitz
import json
import psycopg2
import random
import string
from datetime import datetime, timedelta
from groq import Groq
from tavily import TavilyClient
from dotenv import load_dotenv
from gtts import gTTS
from langdetect import detect
from telegram import Update, BotCommand, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes
)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import Pt as DocPt, RGBColor as DocRGB

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY")
DATABASE_URL = os.getenv("DATABASE_URL")
HF_TOKEN = os.getenv("HF_TOKEN")
CARD_NUMBER = os.getenv("CARD_NUMBER")

ADMIN_ID = 8230883785
ADMIN_USERNAME = "temur_uzb7779"

client = Groq(api_key=GROQ_API_KEY)
tavily = TavilyClient(api_key=TAVILY_API_KEY)

user_histories = {}

SEARCH_KEYWORDS = [
    "today", "now", "current", "latest", "news", "price", "rate", "weather",
    "bugun", "hozir", "narx", "kurs", "yangilik", "ob-havo", "oxirgi",
    "сегодня", "сейчас", "курс", "цена", "новости"
]

GTTS_LANG_MAP = {
    "uz": "ru", "en": "en", "ru": "ru", "tr": "tr", "de": "de",
    "fr": "fr", "es": "es", "ar": "ar", "ko": "ko", "ja": "ja",
    "zh-cn": "zh-CN", "zh-tw": "zh-TW", "it": "it", "pt": "pt",
    "nl": "nl", "pl": "pl", "uk": "uk", "kk": "ru", "hi": "hi"
}

LANGUAGES = {
    "en": ("🇬🇧", "English"),
    "ru": ("🇷🇺", "Russian"),
    "uz": ("🇺🇿", "Uzbek"),
    "tr": ("🇹🇷", "Turkish"),
    "de": ("🇩🇪", "German"),
    "fr": ("🇫🇷", "French"),
    "es": ("🇪🇸", "Spanish"),
    "ar": ("🇸🇦", "Arabic"),
    "ko": ("🇰🇷", "Korean"),
    "ja": ("🇯🇵", "Japanese"),
    "zh": ("🇨🇳", "Chinese"),
    "it": ("🇮🇹", "Italian"),
    "pt": ("🇵🇹", "Portuguese"),
    "hi": ("🇮🇳", "Hindi"),
}

TEXTS = {
    "en": {
        "welcome": "Hello! I am Chatbot 🤖\nYour plan: {plan_emoji} {plan}\n\n💬 Chat with me\n🌐 Current news, prices, weather\n📄 Send PDF to analyze\n🖼️ Send image\n🎤 Send voice message\n\n📊 /pptx — PowerPoint\n📝 /word — Word document\n👤 /cv — Write CV\n📧 /email — Write email\n📱 /post — Marketing post\n🔊 /ai_sound — AI Voice (Standard+)\n👥 /referral — Invite friends & earn bonuses\n\n💰 /updateplan — Update plan\n🌐 /language — Change language\n/help — Help\n/reset — Clear history",
        "help": "📌 Commands:\n\n/pptx — PowerPoint 💎\n/word — Word document 💎\n/cv — CV/Resume ⭐\n/email — Email ⭐\n/post — Marketing post\n/biznes — Business plan\n/ai_sound — AI Voice 🔊 ⭐\n/referral — Invite friends 👥\n/stats — My statistics\n/language — Change language 🌐\n/updateplan — Plans & pricing\n/reset — Clear history\n/help — Help\n\n⭐ = Standard or Premium\n💎 = Premium only",
        "plan_info": "💰 Plans & Pricing\n{'─' * 28}\n\n🆓 FREE — Free\n• Chat: 20/day\n• Search: 20/day\n• Image: 20/day\n• Post: 20/day\n• Business: 20/day\n\n⭐ STANDARD — 5 USDT/month\n• Everything in Free: 30/day\n• PDF: 30/day\n• CV: 30/day\n• Email: 30/day\n• AI Voice: 30/day\n\n💎 PREMIUM — 10 USDT/month\n• Everything: Unlimited\n• Voice messages ✓\n• PowerPoint ✓\n• Word documents ✓",
        "limit_reached": "❌ Daily limit reached! Use /updateplan to upgrade.",
        "blocked": "🚫 You are blocked. Contact admin.",
        "history_cleared": "Chat history cleared ✅",
        "generating_voice": "⏳ Generating voice...",
        "voice_example": "Example: /ai_sound Hello, how are you?",
        "choose_language": "🌐 Choose your language:",
        "language_changed": "✅ Language changed to English!",
        "language_cooldown": "⏳ You can change language once per 24 hours. Try again later.",
        "pdf_required": "Please send a PDF file! 📄",
        "reading_pdf": "⏳ Reading PDF...",
        "pdf_locked": "⭐ PDF requires Standard or Premium!\nUse /updateplan to upgrade.",
        "voice_locked": "💎 Voice requires Premium!\nUse /updateplan to upgrade.",
        "pptx_locked": "💎 PowerPoint requires Premium!\nUse /updateplan to upgrade.",
        "word_locked": "💎 Word requires Premium!\nUse /updateplan to upgrade.",
        "cv_locked": "⭐ CV requires Standard or Premium!\nUse /updateplan to upgrade.",
        "email_locked": "⭐ Email requires Standard or Premium!\nUse /updateplan to upgrade.",
        "ai_sound_locked": "⭐ AI Sound requires Standard or Premium!\nUse /updateplan to upgrade.",
        "writing_cv": "⏳ Writing your CV...",
        "writing_email": "⏳ Writing email...",
        "writing_post": "⏳ Writing marketing post...",
        "writing_biznes": "⏳ Writing business plan...",
        "pptx_example": "Example: /pptx artificial intelligence",
        "word_example": "Example: /word business plan",
        "cv_example": "Example: /cv Python developer, 3 years experience",
        "email_example": "Example: /email follow up after interview",
        "post_example": "Example: /post new coffee shop opening",
        "biznes_example": "Example: /biznes online clothing store",
        "creating_pptx": "⏳ Creating presentation on '{topic}'...",
        "creating_word": "⏳ Creating document on '{topic}'...",
        "ready": "✅ Ready!",
        "buy_standard": "⭐ Buy Standard — 5 USDT/month",
        "buy_premium": "💎 Buy Premium — 10 USDT/month",
        "upgrade_standard": "⭐ Upgrade to Standard",
        "upgrade_premium": "💎 Upgrade to Premium",
        "upgrade_plan": "💰 Upgrade Plan",
        "contact_admin": "💬 Contact Admin",
        "you_said": "🎤 You said: ",
        "referral_title": "👥 Referral Program\n{'─' * 28}\n\n🔗 Your invite link:\n{link}\n\n📊 Invited: {count} friends\n\n🎁 Rewards:\n• 10 friends → ⭐ Standard 15 days — {standard_status}\n• 30 friends → 💎 Premium 15 days — {premium_status}\n\nShare this link! 🚀",
        "claim_standard": "🎁 Claim Standard 15 days",
        "claim_premium": "🎁 Claim Premium 15 days",
        "claim_standard_success": "🎉 Congratulations!\n⭐ Standard plan activated for 15 days!",
        "claim_premium_success": "🎉 Congratulations!\n💎 Premium plan activated for 15 days!",
        "claim_error": "❌ Already claimed or not enough referrals!",
        "stats_title": "📊 Your Statistics",
        "payment_info": "💳 Pay to card:\n`{card}`\n\n💵 Amount: {amount} USDT equivalent\n\n📋 After payment:\n1. Take a screenshot\n2. Send it to admin\n3. Your plan will be activated within 1 hour ✅",
    },
    "uz": {
        "welcome": "Salom! Men Chatbot 🤖\nSizning tarifingiz: {plan_emoji} {plan}\n\n💬 Men bilan suhbatlashing\n🌐 Yangiliklar, narxlar, ob-havo\n📄 PDF tahlil uchun yuboring\n🖼️ Rasm yuboring\n🎤 Ovozli xabar yuboring\n\n📊 /pptx — PowerPoint\n📝 /word — Word hujjat\n👤 /cv — CV yozish\n📧 /email — Email yozish\n📱 /post — Marketing post\n🔊 /ai_sound — AI Ovoz (Standard+)\n👥 /referral — Do'stlarni taklif qiling\n\n💰 /updateplan — Tarifni yangilash\n🌐 /language — Tilni o'zgartirish\n/help — Yordam\n/reset — Tarixni tozalash",
        "help": "📌 Buyruqlar:\n\n/pptx — PowerPoint 💎\n/word — Word hujjat 💎\n/cv — CV/Rezyume ⭐\n/email — Email ⭐\n/post — Marketing post\n/biznes — Biznes reja\n/ai_sound — AI Ovoz 🔊 ⭐\n/referral — Do'stlarni taklif qiling 👥\n/stats — Mening statistikam\n/language — Tilni o'zgartirish 🌐\n/updateplan — Tariflar\n/reset — Tarixni tozalash\n/help — Yordam\n\n⭐ = Standart yoki Premium\n💎 = Faqat Premium",
        "plan_info": "💰 Tariflar\n\n🆓 BEPUL\n• Chat: 20/kun\n• Qidiruv: 20/kun\n\n⭐ STANDART — 5 USDT/oy\n• Hamma narsa: 30/kun\n• PDF, CV, Email, AI Ovoz\n\n💎 PREMIUM — 10 USDT/oy\n• Hamma narsa: Cheksiz\n• Ovozli xabar ✓\n• PowerPoint ✓\n• Word ✓",
        "limit_reached": "❌ Kunlik limit tugadi! Yangilash uchun /updateplan.",
        "blocked": "🚫 Siz bloklangansiz. Admin bilan bog'laning.",
        "history_cleared": "Suhbat tarixi tozalandi ✅",
        "generating_voice": "⏳ Ovoz yaratilmoqda...",
        "voice_example": "Misol: /ai_sound Salom, qandaysiz?",
        "choose_language": "🌐 Tilni tanlang:",
        "language_changed": "✅ Til o'zbekchaga o'zgartirildi!",
        "language_cooldown": "⏳ Tilni 24 soatda 1 marta o'zgartirish mumkin. Keyinroq urinib ko'ring.",
        "pdf_required": "Iltimos, PDF fayl yuboring! 📄",
        "reading_pdf": "⏳ PDF o'qilmoqda...",
        "pdf_locked": "⭐ PDF uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "voice_locked": "💎 Ovoz uchun Premium kerak!\n/updateplan orqali yangilang.",
        "pptx_locked": "💎 PowerPoint uchun Premium kerak!\n/updateplan orqali yangilang.",
        "word_locked": "💎 Word uchun Premium kerak!\n/updateplan orqali yangilang.",
        "cv_locked": "⭐ CV uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "email_locked": "⭐ Email uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "ai_sound_locked": "⭐ AI Ovoz uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "writing_cv": "⏳ CV yozilmoqda...",
        "writing_email": "⏳ Email yozilmoqda...",
        "writing_post": "⏳ Marketing post yozilmoqda...",
        "writing_biznes": "⏳ Biznes reja yozilmoqda...",
        "pptx_example": "Misol: /pptx sun'iy intellekt",
        "word_example": "Misol: /word biznes reja",
        "cv_example": "Misol: /cv Python dasturchi, 3 yil tajriba",
        "email_example": "Misol: /email intervyudan keyin",
        "post_example": "Misol: /post yangi kofe do'kon ochilishi",
        "biznes_example": "Misol: /biznes online kiyim do'koni",
        "creating_pptx": "⏳ '{topic}' mavzusida prezentatsiya yaratilmoqda...",
        "creating_word": "⏳ '{topic}' mavzusida hujjat yaratilmoqda...",
        "ready": "✅ Tayyor!",
        "buy_standard": "⭐ Standart sotib olish — 5 USDT/oy",
        "buy_premium": "💎 Premium sotib olish — 10 USDT/oy",
        "upgrade_standard": "⭐ Standartga o'tish",
        "upgrade_premium": "💎 Premiumga o'tish",
        "upgrade_plan": "💰 Tarifni yangilash",
        "contact_admin": "💬 Admin bilan bog'lanish",
        "you_said": "🎤 Siz dedingiz: ",
        "referral_title": "👥 Referal Dasturi\n\n🔗 Sizning havola:\n{link}\n\n📊 Taklif qilingan: {count} do'st\n\n🎁 Mukofotlar:\n• 10 do'st → ⭐ Standart 15 kun — {standard_status}\n• 30 do'st → 💎 Premium 15 kun — {premium_status}\n\nHavolani ulashing! 🚀",
        "claim_standard": "🎁 Standart 15 kunni olish",
        "claim_premium": "🎁 Premium 15 kunni olish",
        "claim_standard_success": "🎉 Tabriklaymiz!\n⭐ Standart tarif 15 kunga yoqildi!",
        "claim_premium_success": "🎉 Tabriklaymiz!\n💎 Premium tarif 15 kunga yoqildi!",
        "claim_error": "❌ Allaqachon olинган yoki yetarli referal yo'q!",
        "stats_title": "📊 Sizning statistikangiz",
        "payment_info": "💳 Kartaga to'lang:\n`{card}`\n\n💵 Miqdor: {amount} USDT ekvivalent\n\n📋 To'lovdan keyin:\n1. Skrinshot oling\n2. Adminga yuboring\n3. Tarifingiz 1 soat ichida yoqiladi ✅",
    },
    "ru": {
        "welcome": "Привет! Я Chatbot 🤖\nВаш тариф: {plan_emoji} {plan}\n\n💬 Общайтесь со мной\n🌐 Новости, цены, погода\n📄 Отправьте PDF для анализа\n🖼️ Отправьте изображение\n🎤 Отправьте голосовое сообщение\n\n📊 /pptx — PowerPoint\n📝 /word — Word документ\n👤 /cv — Написать резюме\n📧 /email — Написать email\n📱 /post — Маркетинговый пост\n🔊 /ai_sound — AI Голос (Standard+)\n👥 /referral — Пригласить друзей\n\n💰 /updateplan — Обновить тариф\n🌐 /language — Сменить язык\n/help — Помощь\n/reset — Очистить историю",
        "help": "📌 Команды:\n\n/pptx — PowerPoint 💎\n/word — Word документ 💎\n/cv — Резюме ⭐\n/email — Email ⭐\n/post — Маркетинговый пост\n/biznes — Бизнес-план\n/ai_sound — AI Голос 🔊 ⭐\n/referral — Пригласить друзей 👥\n/stats — Моя статистика\n/language — Сменить язык 🌐\n/updateplan — Тарифы\n/reset — Очистить историю\n/help — Помощь\n\n⭐ = Standard или Premium\n💎 = Только Premium",
        "plan_info": "💰 Тарифы\n\n🆓 БЕСПЛАТНО\n• Чат: 20/день\n• Поиск: 20/день\n\n⭐ СТАНДАРТ — 5 USDT/месяц\n• Всё: 30/день\n• PDF, CV, Email, AI Голос\n\n💎 ПРЕМИУМ — 10 USDT/месяц\n• Всё: Безлимитно\n• Голосовые сообщения ✓\n• PowerPoint ✓\n• Word ✓",
        "limit_reached": "❌ Дневной лимит исчерпан! Используйте /updateplan.",
        "blocked": "🚫 Вы заблокированы. Свяжитесь с администратором.",
        "history_cleared": "История чата очищена ✅",
        "generating_voice": "⏳ Генерация голоса...",
        "voice_example": "Пример: /ai_sound Привет, как дела?",
        "choose_language": "🌐 Выберите язык:",
        "language_changed": "✅ Язык изменён на русский!",
        "language_cooldown": "⏳ Язык можно менять раз в 24 часа. Попробуйте позже.",
        "pdf_required": "Пожалуйста, отправьте PDF файл! 📄",
        "reading_pdf": "⏳ Читаю PDF...",
        "pdf_locked": "⭐ PDF требует Standard или Premium!\nИспользуйте /updateplan.",
        "voice_locked": "💎 Голос требует Premium!\nИспользуйте /updateplan.",
        "pptx_locked": "💎 PowerPoint требует Premium!\nИспользуйте /updateplan.",
        "word_locked": "💎 Word требует Premium!\nИспользуйте /updateplan.",
        "cv_locked": "⭐ Резюме требует Standard или Premium!\nИспользуйте /updateplan.",
        "email_locked": "⭐ Email требует Standard или Premium!\nИспользуйте /updateplan.",
        "ai_sound_locked": "⭐ AI Голос требует Standard или Premium!\nИспользуйте /updateplan.",
        "writing_cv": "⏳ Пишу резюме...",
        "writing_email": "⏳ Пишу email...",
        "writing_post": "⏳ Пишу маркетинговый пост...",
        "writing_biznes": "⏳ Пишу бизнес-план...",
        "pptx_example": "Пример: /pptx искусственный интеллект",
        "word_example": "Пример: /word бизнес-план",
        "cv_example": "Пример: /cv Python разработчик, 3 года опыта",
        "email_example": "Пример: /email после собеседования",
        "post_example": "Пример: /post открытие новой кофейни",
        "biznes_example": "Пример: /biznes онлайн магазин одежды",
        "creating_pptx": "⏳ Создаю презентацию на тему '{topic}'...",
        "creating_word": "⏳ Создаю документ на тему '{topic}'...",
        "ready": "✅ Готово!",
        "buy_standard": "⭐ Купить Standard — 5 USDT/месяц",
        "buy_premium": "💎 Купить Premium — 10 USDT/месяц",
        "upgrade_standard": "⭐ Перейти на Standard",
        "upgrade_premium": "💎 Перейти на Premium",
        "upgrade_plan": "💰 Обновить тариф",
        "contact_admin": "💬 Связаться с админом",
        "you_said": "🎤 Вы сказали: ",
        "referral_title": "👥 Реферальная программа\n\n🔗 Ваша ссылка:\n{link}\n\n📊 Приглашено: {count} друзей\n\n🎁 Награды:\n• 10 друзей → ⭐ Standard 15 дней — {standard_status}\n• 30 друзей → 💎 Premium 15 дней — {premium_status}\n\nПоделитесь ссылкой! 🚀",
        "claim_standard": "🎁 Получить Standard 15 дней",
        "claim_premium": "🎁 Получить Premium 15 дней",
        "claim_standard_success": "🎉 Поздравляем!\n⭐ Standard тариф активирован на 15 дней!",
        "claim_premium_success": "🎉 Поздравляем!\n💎 Premium тариф активирован на 15 дней!",
        "claim_error": "❌ Уже получено или недостаточно рефералов!",
        "stats_title": "📊 Ваша статистика",
        "payment_info": "💳 Оплатите на карту:\n`{card}`\n\n💵 Сумма: {amount} USDT\n\n📋 После оплаты:\n1. Сделайте скриншот\n2. Отправьте админу\n3. Тариф активируется в течение 1 часа ✅",
    },
}

# Qolgan tillar uchun inglizcha ishlatamiz
for lang in LANGUAGES:
    if lang not in TEXTS:
        TEXTS[lang] = TEXTS["en"]

def get_text(lang, key, **kwargs):
    t = TEXTS.get(lang, TEXTS["en"]).get(key, TEXTS["en"].get(key, ""))
    if kwargs:
        try:
            t = t.format(**kwargs)
        except:
            pass
    return t

def detect_lang(text):
    try:
        lang = detect(text)
        return GTTS_LANG_MAP.get(lang, "en")
    except:
        return "en"

def generate_referral_code():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))

def get_conn():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            user_id BIGINT UNIQUE,
            username TEXT,
            full_name TEXT,
            plan TEXT DEFAULT 'free',
            expires_at TIMESTAMP,
            is_blocked BOOLEAN DEFAULT FALSE,
            joined_at TIMESTAMP DEFAULT NOW(),
            usage_chat INTEGER DEFAULT 0,
            usage_search INTEGER DEFAULT 0,
            usage_image INTEGER DEFAULT 0,
            usage_post INTEGER DEFAULT 0,
            usage_biznes INTEGER DEFAULT 0,
            usage_pdf INTEGER DEFAULT 0,
            usage_cv INTEGER DEFAULT 0,
            usage_email INTEGER DEFAULT 0,
            usage_tts INTEGER DEFAULT 0,
            referral_code TEXT,
            referred_by BIGINT,
            referral_count INTEGER DEFAULT 0,
            claimed_standard BOOLEAN DEFAULT FALSE,
            claimed_premium BOOLEAN DEFAULT FALSE,
            language TEXT DEFAULT 'en',
            language_changed_at TIMESTAMP,
            last_reset DATE DEFAULT CURRENT_DATE
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS user_memory (
            user_id BIGINT PRIMARY KEY,
            name TEXT,
            facts TEXT
        )
    """)
    conn.commit()
    conn.close()

def ensure_user(user_id, username=None, full_name=None, referred_by=None):
    try:
        conn = get_conn()
        c = conn.cursor()
        code = generate_referral_code()
        c.execute("""
            INSERT INTO users (user_id, username, full_name, referral_code, referred_by)
            VALUES (%s, %s, %s, %s, %s)
            ON CONFLICT (user_id) DO UPDATE SET
                username = EXCLUDED.username,
                full_name = EXCLUDED.full_name
        """, (user_id, username, full_name, code, referred_by))
        conn.commit()
        if referred_by:
            c.execute("UPDATE users SET referral_count = referral_count + 1 WHERE user_id = %s", (referred_by,))
            conn.commit()
        conn.close()
    except:
        pass

def get_user(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            SELECT plan, expires_at, is_blocked, full_name, username,
                   referral_code, referral_count, claimed_standard, claimed_premium,
                   language, language_changed_at
            FROM users WHERE user_id = %s
        """, (user_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return {"plan": "free", "expires_at": None, "is_blocked": False,
                    "full_name": None, "username": None, "referral_code": None,
                    "referral_count": 0, "claimed_standard": False, "claimed_premium": False,
                    "language": "en", "language_changed_at": None}
        plan, expires_at, is_blocked, full_name, username, referral_code, referral_count, claimed_standard, claimed_premium, language, language_changed_at = row
        if expires_at and datetime.now() > expires_at and plan != 'free':
            set_plan(user_id, "free", None)
            plan = "free"
            expires_at = None
        return {
            "plan": plan, "expires_at": expires_at, "is_blocked": is_blocked,
            "full_name": full_name, "username": username, "referral_code": referral_code,
            "referral_count": referral_count or 0, "claimed_standard": claimed_standard or False,
            "claimed_premium": claimed_premium or False, "language": language or "en",
            "language_changed_at": language_changed_at
        }
    except:
        return {"plan": "free", "expires_at": None, "is_blocked": False,
                "full_name": None, "username": None, "referral_code": None,
                "referral_count": 0, "claimed_standard": False, "claimed_premium": False,
                "language": "en", "language_changed_at": None}

def set_plan(user_id, plan, days=30):
    try:
        conn = get_conn()
        c = conn.cursor()
        expires_at = datetime.now() + timedelta(days=days) if days else None
        c.execute("UPDATE users SET plan=%s, expires_at=%s WHERE user_id=%s", (plan, expires_at, user_id))
        conn.commit()
        conn.close()
    except:
        pass

def set_blocked(user_id, blocked):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("UPDATE users SET is_blocked=%s WHERE user_id=%s", (blocked, user_id))
        conn.commit()
        conn.close()
    except:
        pass

def set_language(user_id, lang):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("UPDATE users SET language=%s, language_changed_at=%s WHERE user_id=%s",
                  (lang, datetime.now(), user_id))
        conn.commit()
        conn.close()
    except:
        pass

def check_limit(user_id, feature, limit):
    if limit == -1:
        return True
    if limit == 0:
        return False
    try:
        conn = get_conn()
        c = conn.cursor()
        today = datetime.now().date()
        c.execute("SELECT last_reset FROM users WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        if row and row[0] < today:
            c.execute("""
                UPDATE users SET
                usage_chat=0, usage_search=0, usage_image=0,
                usage_post=0, usage_biznes=0, usage_pdf=0,
                usage_cv=0, usage_email=0, usage_tts=0,
                last_reset=%s WHERE user_id=%s
            """, (today, user_id))
            conn.commit()
        c.execute(f"SELECT usage_{feature} FROM users WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        usage = row[0] if row else 0
        if usage >= limit:
            conn.close()
            return False
        c.execute(f"UPDATE users SET usage_{feature} = usage_{feature} + 1 WHERE user_id = %s", (user_id,))
        conn.commit()
        conn.close()
        return True
    except:
        return True

def get_limits(plan):
    if plan == "premium":
        return {k: -1 for k in ["chat","search","image","post","biznes","pdf","cv","email","voice","pptx","word","tts"]}
    elif plan == "standard":
        return {
            "chat": 30, "search": 30, "image": 30, "post": 30, "biznes": 30,
            "pdf": 30, "cv": 30, "email": 30, "voice": 0, "pptx": 0, "word": 0, "tts": 30
        }
    else:
        return {
            "chat": 20, "search": 20, "image": 20, "post": 20, "biznes": 20,
            "pdf": 0, "cv": 0, "email": 0, "voice": 0, "pptx": 0, "word": 0, "tts": 0
        }

def get_memory(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT name, facts FROM user_memory WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        conn.close()
        return {"name": row[0], "facts": row[1]} if row else None
    except:
        return None

def save_memory(user_id, name, facts):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO user_memory (user_id, name, facts) VALUES (%s, %s, %s)
            ON CONFLICT (user_id) DO UPDATE SET name=%s, facts=%s
        """, (user_id, name, facts, name, facts))
        conn.commit()
        conn.close()
    except:
        pass

async def update_memory(user_id, user_text, reply):
    memory = get_memory(user_id)
    current_facts = memory["facts"] if memory else ""
    current_name = memory["name"] if memory else ""
    prompt = f"""Extract user info from conversation.
Current name: {current_name}, facts: {current_facts}
User: {user_text}
Assistant: {reply}
Reply ONLY JSON: {{"name": "name or empty", "facts": "short facts"}}"""
    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}]
        )
        content = response.choices[0].message.content.strip()
        if "```" in content:
            content = content.split("```")[1]
            if content.startswith("json"):
                content = content[4:]
        data = json.loads(content)
        save_memory(user_id, data.get("name", current_name), data.get("facts", current_facts))
    except:
        pass

def needs_search(text):
    return any(k in text.lower() for k in SEARCH_KEYWORDS)

def create_pptx(title, slides_data):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(0x1E, 0x1E, 0x2E)
    title_box = slide.shapes.title
    title_box.text = title
    title_box.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    title_box.text_frame.paragraphs[0].font.size = Pt(40)
    title_box.text_frame.paragraphs[0].font.bold = True
    for s in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0x1E, 0x1E, 0x2E)
        t = slide.shapes.title
        t.text = s["title"]
        t.text_frame.paragraphs[0].font.color.rgb = RGBColor(0x89, 0xB4, 0xFA)
        t.text_frame.paragraphs[0].font.size = Pt(28)
        t.text_frame.paragraphs[0].font.bold = True
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for i, point in enumerate(s["points"]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"• {point}"
            p.font.color.rgb = RGBColor(0xCD, 0xD6, 0xF4)
            p.font.size = Pt(18)
    prs.save("presentation.pptx")
    return "presentation.pptx"

def create_docx(title, sections):
    doc = Document()
    h = doc.add_heading(title, 0)
    h.runs[0].font.color.rgb = DocRGB(0x1E, 0x3A, 0x8A)
    for section in sections:
        doc.add_heading(section["title"], level=1)
        for point in section["points"]:
            doc.add_paragraph(point, style="List Bullet")
        doc.add_paragraph()
    doc.save("document.docx")
    return "document.docx"

async def ai_generate(prompt):
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

async def generate_content(topic, doc_type):
    if doc_type == "pptx":
        prompt = f"""Create presentation for: '{topic}'. ONLY JSON:
{{"title":"title","slides":[{{"title":"s","points":["p1","p2","p3"]}}]}}
Min 5 slides. Same language as topic."""
    else:
        prompt = f"""Create Word doc for: '{topic}'. ONLY JSON:
{{"title":"title","sections":[{{"title":"s","points":["p1","p2","p3"]}}]}}
Min 4 sections. Same language as topic."""
    return await ai_generate(prompt)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    referred_by = None
    if context.args:
        try:
            ref_code = context.args[0]
            conn = get_conn()
            c = conn.cursor()
            c.execute("SELECT user_id FROM users WHERE referral_code = %s", (ref_code,))
            row = c.fetchone()
            conn.close()
            if row and row[0] != user.id:
                referred_by = row[0]
        except:
            pass
    ensure_user(user.id, user.username, user.full_name, referred_by)
    u = get_user(user.id)
    lang = u["language"]
    plan = u["plan"].upper()
    plan_emoji = {"FREE": "🆓", "STANDARD": "⭐", "PREMIUM": "💎"}.get(plan, "🆓")
    await update.message.reply_text(get_text(lang, "welcome", plan_emoji=plan_emoji, plan=plan))

async def language_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    last_changed = u["language_changed_at"]
    if last_changed and datetime.now() - last_changed < timedelta(hours=24):
        await update.message.reply_text(get_text(lang, "language_cooldown"))
        return
    keyboard = []
    row = []
    for code, (flag, name) in LANGUAGES.items():
        row.append(InlineKeyboardButton(f"{flag} {name}", callback_data=f"setlang_{code}"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    await update.message.reply_text(
        get_text(lang, "choose_language"),
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            SELECT usage_chat, usage_search, usage_image, usage_post,
                   usage_biznes, usage_pdf, usage_cv, usage_email, usage_tts
            FROM users WHERE user_id = %s
        """, (user.id,))
        row = c.fetchone()
        conn.close()
        if not row:
            await update.message.reply_text("No stats yet.")
            return
        chat, search, image, post, biznes, pdf, cv, email, tts = row
        plan = u["plan"]
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}.get(plan, "🆓")
        expires = u["expires_at"].strftime("%d.%m.%Y") if u.get("expires_at") else "—"
        await update.message.reply_text(
            f"{get_text(lang, 'stats_title')}\n"
            f"{'─' * 28}\n\n"
            f"👤 Plan: {plan_emoji} {plan.upper()}\n"
            f"📅 Expires: {expires}\n"
            f"👥 Referrals: {u['referral_count']}\n\n"
            f"📊 Today's usage:\n"
            f"💬 Chat: {chat}\n"
            f"🌐 Search: {search}\n"
            f"🖼️ Image: {image}\n"
            f"📱 Post: {post}\n"
            f"💼 Business: {biznes}\n"
            f"📄 PDF: {pdf}\n"
            f"👤 CV: {cv}\n"
            f"📧 Email: {email}\n"
            f"🔊 AI Voice: {tts}"
        )
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def referral_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    code = u["referral_code"]
    count = u["referral_count"]
    claimed_standard = u["claimed_standard"]
    claimed_premium = u["claimed_premium"]
    bot = await context.bot.get_me()
    link = f"https://t.me/{bot.username}?start={code}"
    keyboard = []
    if count >= 10 and not claimed_standard:
        keyboard.append([InlineKeyboardButton(get_text(lang, "claim_standard"), callback_data="claim_standard")])
    if count >= 30 and not claimed_premium:
        keyboard.append([InlineKeyboardButton(get_text(lang, "claim_premium"), callback_data="claim_premium")])
    standard_status = "✅" if claimed_standard else ("🔓 Ready!" if count >= 10 else f"{count}/10")
    premium_status = "✅" if claimed_premium else ("🔓 Ready!" if count >= 30 else f"{count}/30")
    await update.message.reply_text(
        get_text(lang, "referral_title", link=link, count=count,
                 standard_status=standard_status, premium_status=premium_status),
        reply_markup=InlineKeyboardMarkup(keyboard) if keyboard else None
    )

async def updateplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    plan = u["plan"]
    expires = u["expires_at"].strftime("%d.%m.%Y") if u.get("expires_at") else "—"
    plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}.get(plan, "🆓")
    keyboard = [
        [InlineKeyboardButton(get_text(lang, "buy_standard"), callback_data="buy_standard")],
        [InlineKeyboardButton(get_text(lang, "buy_premium"), callback_data="buy_premium")],
    ]
    await update.message.reply_text(
        f"💰 Plans & Pricing\n{'─' * 28}\n\n"
        f"🆓 FREE\n• Chat: 20/day • Search: 20/day • Image: 20/day\n• Post: 20/day • Business: 20/day\n\n"
        f"⭐ STANDARD — 5 USDT/month\n• Everything: 30/day\n• PDF ✓ • CV ✓ • Email ✓ • AI Voice ✓\n\n"
        f"💎 PREMIUM — 10 USDT/month\n• Everything: Unlimited\n• Voice messages ✓ • PowerPoint ✓ • Word ✓\n\n"
        f"{'─' * 28}\n"
        f"👤 {plan_emoji} {plan.upper()}\n📅 {expires}",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    u = get_user(user_id)
    lang = u["language"]

    if data.startswith("setlang_"):
        new_lang = data.split("_")[1]
        last_changed = u["language_changed_at"]
        if last_changed and datetime.now() - last_changed < timedelta(hours=24):
            await query.answer(get_text(lang, "language_cooldown"), show_alert=True)
            return
        set_language(user_id, new_lang)
        lang_name = LANGUAGES.get(new_lang, ("", "Unknown"))[1]
        await query.message.edit_text(f"✅ Language changed to {lang_name}!")

    elif data == "claim_standard":
        if u["referral_count"] >= 10 and not u["claimed_standard"]:
            set_plan(user_id, "standard", 15)
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE users SET claimed_standard=TRUE WHERE user_id=%s", (user_id,))
            conn.commit()
            conn.close()
            await query.message.edit_text(get_text(lang, "claim_standard_success"))
        else:
            await query.answer(get_text(lang, "claim_error"), show_alert=True)

    elif data == "claim_premium":
        if u["referral_count"] >= 30 and not u["claimed_premium"]:
            set_plan(user_id, "premium", 15)
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE users SET claimed_premium=TRUE WHERE user_id=%s", (user_id,))
            conn.commit()
            conn.close()
            await query.message.edit_text(get_text(lang, "claim_premium_success"))
        else:
            await query.answer(get_text(lang, "claim_error"), show_alert=True)

    elif data == "buy_standard":
        keyboard = [[InlineKeyboardButton(get_text(lang, "contact_admin"), url=f"https://t.me/{ADMIN_USERNAME}")]]
        await query.message.reply_text(
            f"⭐ STANDARD — 5 USDT/month\n{'─' * 30}\n\n" +
            get_text(lang, "payment_info", card=CARD_NUMBER, amount=5),
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data == "buy_premium":
        keyboard = [[InlineKeyboardButton(get_text(lang, "contact_admin"), url=f"https://t.me/{ADMIN_USERNAME}")]]
        await query.message.reply_text(
            f"💎 PREMIUM — 10 USDT/month\n{'─' * 30}\n\n" +
            get_text(lang, "payment_info", card=CARD_NUMBER, amount=10),
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data.startswith("ap_setplan_"):
        parts = data.split("_")
        target_id = int(parts[2])
        plan = parts[3]
        days = 30 if plan != "free" else None
        set_plan(target_id, plan, days)
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        await query.message.edit_text(f"✅ Done!\nUser {target_id} → {plan_emoji[plan]} {plan.upper()}")
        try:
            await context.bot.send_message(
                chat_id=target_id,
                text=f"{plan_emoji[plan]} Your plan has been updated to {plan.upper()}!"
            )
        except:
            pass

    elif data.startswith("ap_block_"):
        target_id = int(data.split("_")[2])
        set_blocked(target_id, True)
        await query.message.edit_text(f"🚫 User {target_id} blocked!")

    elif data.startswith("ap_unblock_"):
        target_id = int(data.split("_")[2])
        set_blocked(target_id, False)
        await query.message.edit_text(f"✅ User {target_id} unblocked!")

async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM users")
        total = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM users WHERE plan='standard'")
        standard = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM users WHERE plan='premium'")
        premium = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM users WHERE is_blocked=TRUE")
        blocked = c.fetchone()[0]
        conn.close()
        revenue = standard * 5 + premium * 10
    except:
        total = standard = premium = blocked = revenue = 0
    await update.message.reply_text(
        f"🔧 Admin Panel\n{'═' * 28}\n\n"
        f"👥 Total: {total}\n🆓 Free: {total-standard-premium}\n"
        f"⭐ Standard: {standard}\n💎 Premium: {premium}\n"
        f"🚫 Blocked: {blocked}\n💰 Revenue: ~${revenue}\n\n"
        f"/users — All users\n/find [id] — Manage user\n/broadcast [text] — Message all"
    )

async def users_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            SELECT user_id, username, plan, is_blocked FROM users
            ORDER BY CASE plan WHEN 'premium' THEN 1 WHEN 'standard' THEN 2 ELSE 3 END, joined_at DESC
        """)
        rows = c.fetchall()
        conn.close()
        if not rows:
            await update.message.reply_text("No users yet.")
            return
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        text = f"👥 All Users: {len(rows)}\n{'═' * 30}\n\n"
        for row in rows:
            uid, username, plan, is_blocked = row
            uname = f"@{username}" if username else "no_username"
            blocked = " 🚫" if is_blocked else ""
            text += f"{plan_emoji.get(plan,'🆓')} {uname}    {uid}{blocked}\n"
        if len(text) > 4000:
            for part in [text[i:i+4000] for i in range(0, len(text), 4000)]:
                await update.message.reply_text(part)
        else:
            await update.message.reply_text(text)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def find_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    if not context.args:
        await update.message.reply_text("Usage: /find [user_id]")
        return
    try:
        target_id = int(context.args[0])
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT plan, expires_at, is_blocked, full_name, username FROM users WHERE user_id = %s", (target_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            await update.message.reply_text(f"❌ User {target_id} not found!")
            return
        plan, expires_at, is_blocked, full_name, username = row
        expires = expires_at.strftime("%d.%m.%Y") if expires_at else "—"
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        keyboard = [
            [
                InlineKeyboardButton("🆓 Free", callback_data=f"ap_setplan_{target_id}_free"),
                InlineKeyboardButton("⭐ Standard", callback_data=f"ap_setplan_{target_id}_standard"),
                InlineKeyboardButton("💎 Premium", callback_data=f"ap_setplan_{target_id}_premium"),
            ],
            [InlineKeyboardButton(
                "✅ Unblock" if is_blocked else "🚫 Block",
                callback_data=f"ap_unblock_{target_id}" if is_blocked else f"ap_block_{target_id}"
            )]
        ]
        await update.message.reply_text(
            f"👤 User Info\n{'─' * 25}\n"
            f"🆔 ID: {target_id}\n👤 Name: {full_name or '—'}\n"
            f"📱 @{username or '—'}\n"
            f"📋 Plan: {plan_emoji.get(plan,'🆓')} {plan.upper()}\n"
            f"📅 Expires: {expires}\n"
            f"🔰 {'🚫 Blocked' if is_blocked else '✅ Active'}\n"
            f"{'─' * 25}",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    message = " ".join(context.args)
    if not message:
        await update.message.reply_text("Usage: /broadcast [message]")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT user_id FROM users WHERE is_blocked = FALSE")
        rows = c.fetchall()
        conn.close()
        sent = 0
        for row in rows:
            try:
                await context.bot.send_message(chat_id=row[0], text=f"📢 {message}")
                sent += 1
            except:
                pass
        await update.message.reply_text(f"✅ Sent to {sent} users!")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    await update.message.reply_text(get_text(u["language"], "help"))

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    user_histories[user.id] = []
    await update.message.reply_text(get_text(u["language"], "history_cleared"))

async def pptx_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    if get_limits(u["plan"])["pptx"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_premium"), callback_data="buy_premium")]]
        await update.message.reply_text(get_text(lang, "pptx_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text(get_text(lang, "pptx_example"))
        return
    await update.message.reply_text(get_text(lang, "creating_pptx", topic=topic))
    try:
        content_str = await generate_content(topic, "pptx")
        content_str = content_str.strip()
        if "```" in content_str:
            content_str = content_str.split("```")[1]
            if content_str.startswith("json"):
                content_str = content_str[4:]
        content = json.loads(content_str)
        path = create_pptx(content["title"], content["slides"])
        with open(path, "rb") as f:
            await update.message.reply_document(document=f, filename=f"{topic}.pptx", caption=get_text(lang, "ready"))
        os.remove(path)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def word_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    if get_limits(u["plan"])["word"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_premium"), callback_data="buy_premium")]]
        await update.message.reply_text(get_text(lang, "word_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text(get_text(lang, "word_example"))
        return
    await update.message.reply_text(get_text(lang, "creating_word", topic=topic))
    try:
        content_str = await generate_content(topic, "docx")
        content_str = content_str.strip()
        if "```" in content_str:
            content_str = content_str.split("```")[1]
            if content_str.startswith("json"):
                content_str = content_str[4:]
        content = json.loads(content_str)
        path = create_docx(content["title"], content["sections"])
        with open(path, "rb") as f:
            await update.message.reply_document(document=f, filename=f"{topic}.docx", caption=get_text(lang, "ready"))
        os.remove(path)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def cv_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["cv"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "cv_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "cv", limits["cv"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    info = " ".join(context.args)
    if not info:
        await update.message.reply_text(get_text(lang, "cv_example"))
        return
    await update.message.reply_text(get_text(lang, "writing_cv"))
    try:
        reply = await ai_generate(f"Write a professional CV/Resume for: {info}\nFormat: Summary, Experience, Skills, Education. ATS-friendly.")
        await update.message.reply_text(f"✅ CV:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def email_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["email"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "email_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "email", limits["email"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text(get_text(lang, "email_example"))
        return
    await update.message.reply_text(get_text(lang, "writing_email"))
    try:
        reply = await ai_generate(f"Write a professional email about: {topic}\nInclude: Subject, greeting, body, closing.")
        await update.message.reply_text(f"✅ Email:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def post_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if not check_limit(user.id, "post", limits["post"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text(get_text(lang, "post_example"))
        return
    await update.message.reply_text(get_text(lang, "writing_post"))
    try:
        reply = await ai_generate(f"Write an engaging social media post about: {topic}\nInclude: Hook, content, call to action, hashtags.")
        await update.message.reply_text(f"✅ Post:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def biznes_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if not check_limit(user.id, "biznes", limits["biznes"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    idea = " ".join(context.args)
    if not idea:
        await update.message.reply_text(get_text(lang, "biznes_example"))
        return
    await update.message.reply_text(get_text(lang, "writing_biznes"))
    try:
        reply = await ai_generate(f"Write a detailed business plan for: {idea}\nInclude: Executive Summary, Market Analysis, Products, Marketing Strategy, Financial Plan.")
        await update.message.reply_text(f"✅ Business plan:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_tts(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    ensure_user(user_id, user.username, user.full_name)
    u = get_user(user_id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits.get("tts") == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "ai_sound_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user_id, "tts", limits["tts"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    text = " ".join(context.args)
    if not text:
        await update.message.reply_text(get_text(lang, "voice_example"))
        return
    await update.message.reply_text(get_text(lang, "generating_voice"))
    try:
        tts_lang = detect_lang(text)
        tts = gTTS(text=text, lang=tts_lang)
        buf = io.BytesIO()
        tts.write_to_fp(buf)
        buf.seek(0)
        await update.message.reply_voice(voice=buf)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    ensure_user(user_id, user.username, user.full_name)
    u = get_user(user_id)
    lang = u["language"]
    if u["is_blocked"]:
        await update.message.reply_text(get_text(lang, "blocked"))
        return
    limits = get_limits(u["plan"])
    if not check_limit(user_id, "chat", limits["chat"]):
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_plan"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "limit_reached"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    user_text = update.message.text
    memory = get_memory(user_id)
    memory_context = ""
    if memory and (memory.get("name") or memory.get("facts")):
        memory_context = f"User info — name: {memory['name']}, facts: {memory['facts']}. "
    system_prompt = (
        "You are a professional AI assistant. "
        "VERY IMPORTANT: Always reply in the SAME language the user uses in their CURRENT message. "
        "If user writes in Uzbek, reply in Uzbek. "
        "If user writes in English, reply in English. "
        "If user writes in Russian, reply in Russian. "
        "If user asks you to speak in another language, do it immediately without resistance. "
        "Never switch languages on your own without user request. "
        "Keep answers short, clear and natural. "
        + memory_context
    )
    if user_id not in user_histories:
        user_histories[user_id] = [{"role": "system", "content": system_prompt}]
    else:
        user_histories[user_id][0]["content"] = system_prompt
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        if needs_search(user_text) and limits["search"] != 0:
            search_results = tavily.search(query=user_text, max_results=3)
            search_content = "\n\n".join([f"Source: {r['url']}\n{r['content']}" for r in search_results.get("results", [])])
            message_content = f"User question: {user_text}\n\nWeb results:\n{search_content}\n\nAnswer in the same language as the question."
        else:
            message_content = user_text
        user_histories[user_id].append({"role": "user", "content": message_content})
        response = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=user_histories[user_id])
        reply = response.choices[0].message.content
        user_histories[user_id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
        await update_memory(user_id, user_text, reply)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    ensure_user(user_id, user.username, user.full_name)
    u = get_user(user_id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    if get_limits(u["plan"])["voice"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_premium"), callback_data="buy_premium")]]
        await update.message.reply_text(get_text(lang, "voice_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        file = await context.bot.get_file(update.message.voice.file_id)
        audio_data = requests.get(file.file_path).content
        with open("voice.ogg", "wb") as f:
            f.write(audio_data)
        with open("voice.ogg", "rb") as f:
            transcription = client.audio.transcriptions.create(file=("voice.ogg", f.read()), model="whisper-large-v3")
        user_text = transcription.text
        await update.message.reply_text(get_text(lang, "you_said") + user_text)
        if user_id not in user_histories:
            user_histories[user_id] = [{"role": "system", "content": "You are a professional AI assistant. Always reply in the same language the user writes in."}]
        user_histories[user_id].append({"role": "user", "content": user_text})
        response = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=user_histories[user_id])
        reply = response.choices[0].message.content
        user_histories[user_id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    ensure_user(user_id, user.username, user.full_name)
    u = get_user(user_id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if not check_limit(user_id, "image", limits["image"]):
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_plan"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "limit_reached"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)
        caption = update.message.caption or "What is in this image?"
        response = client.chat.completions.create(
            model="meta-llama/llama-4-scout-17b-16e-instruct",
            messages=[{"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": file.file_path}},
                {"type": "text", "text": caption}
            ]}]
        )
        await update.message.reply_text(response.choices[0].message.content)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    ensure_user(user_id, user.username, user.full_name)
    u = get_user(user_id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["pdf"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "pdf_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user_id, "pdf", limits["pdf"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        doc = update.message.document
        if not doc.file_name.endswith(".pdf"):
            await update.message.reply_text(get_text(lang, "pdf_required"))
            return
        await update.message.reply_text(get_text(lang, "reading_pdf"))
        file = await context.bot.get_file(doc.file_id)
        pdf_data = requests.get(file.file_path).content
        with open("temp.pdf", "wb") as f:
            f.write(pdf_data)
        pdf = fitz.open("temp.pdf")
        text = "".join(page.get_text() for page in pdf)
        pdf.close()
        os.remove("temp.pdf")
        if len(text) > 12000:
            text = text[:12000] + "..."
        caption = update.message.caption or "Summarize this document and explain the key points."
        if user_id not in user_histories:
            user_histories[user_id] = [{"role": "system", "content": "You are a professional AI assistant. Always reply in the same language the user writes in."}]
        user_histories[user_id].append({"role": "user", "content": f"PDF:\n\n{text}\n\nRequest: {caption}"})
        response = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=user_histories[user_id])
        reply = response.choices[0].message.content
        user_histories[user_id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def post_init(app):
    init_db()
    await app.bot.set_my_commands([
        BotCommand("start", "Start the bot"),
        BotCommand("updateplan", "Plans & pricing"),
        BotCommand("pptx", "PowerPoint (Premium)"),
        BotCommand("word", "Word document (Premium)"),
        BotCommand("cv", "Write CV (Standard+)"),
        BotCommand("email", "Write email (Standard+)"),
        BotCommand("post", "Marketing post"),
        BotCommand("biznes", "Business plan"),
        BotCommand("ai_sound", "AI Voice (Standard+)"),
        BotCommand("referral", "Invite friends & earn bonuses"),
        BotCommand("stats", "My statistics"),
        BotCommand("language", "Change language"),
        BotCommand("reset", "Clear history"),
        BotCommand("help", "Help"),
    ])

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("updateplan", updateplan_command))
    app.add_handler(CommandHandler("pptx", pptx_command))
    app.add_handler(CommandHandler("word", word_command))
    app.add_handler(CommandHandler("cv", cv_command))
    app.add_handler(CommandHandler("email", email_command))
    app.add_handler(CommandHandler("post", post_command))
    app.add_handler(CommandHandler("biznes", biznes_command))
    app.add_handler(CommandHandler("ai_sound", handle_tts))
    app.add_handler(CommandHandler("referral", referral_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("language", language_command))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("find", find_command))
    app.add_handler(CommandHandler("users", users_command))
    app.add_handler(CommandHandler("broadcast", broadcast_command))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.PHOTO, handle_image))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_document))
    print("Bot is running... ✅")
    app.run_polling()

if __name__ == "__main__":
    main()
