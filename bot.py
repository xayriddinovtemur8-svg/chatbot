import os
import io
import requests
import fitz
import json
import psycopg2
import random
import string
from datetime import datetime, timedelta, date
from groq import Groq
from tavily import TavilyClient
from dotenv import load_dotenv
from gtts import gTTS
from langdetect import detect
import google.generativeai as genai
from telegram import Update, BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, LabeledPrice
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    CallbackQueryHandler,
    PreCheckoutQueryHandler,
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
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

genai.configure(api_key=GEMINI_API_KEY)

ADMIN_ID = 8230883785
ADMIN_USERNAME = "temur_uzb7779"

STARS_STANDARD = 500
STARS_PREMIUM = 1000
STARS_GROUP = 7700

client = Groq(api_key=GROQ_API_KEY)
tavily = TavilyClient(api_key=TAVILY_API_KEY)

user_histories = {}
user_states = {}

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
        "welcome": "Hello! I am Chatbot 🤖\nYour plan: {plan_emoji} {plan}\n\n💬 Chat with me\n🌐 Current news, prices, weather\n📄 Send PDF to analyze\n🖼️ Send image\n🎤 Send voice message\n\n📊 /pptx — PowerPoint\n📝 /word — Word document\n👤 /cv — Write CV\n📧 /email — Write email\n📱 /post — Marketing post\n🔊 /ai_sound — AI Voice\n🎨 /imagine — AI Image\n🌐 /translate — Translator\n💻 /code — Code writer\n📋 /document — Document creator\n👥 /referral — Invite friends\n🎁 /gift — Daily bonus\n🏆 /top — Top users\n📊 /stats — My statistics\n\n💰 /updateplan — Update plan\n🌐 /language — Change language\n/help — Help\n/reset — Clear history",
        "help": "📌 Commands:\n\n/pptx — PowerPoint 💎\n/word — Word document 💎\n/cv — CV/Resume ⭐\n/email — Email ⭐\n/post — Marketing post\n/biznes — Business plan\n/ai_sound — AI Voice 🔊 ⭐\n/imagine — AI Image 🎨 ⭐\n/translate — Translator 🌐\n/code — Code writer 💻\n/document — Document creator 📋 ⭐\n/referral — Invite friends 👥\n/gift — Daily bonus 🎁\n/top — Top users 🏆\n/stats — My statistics 📊\n/language — Change language 🌐\n/updateplan — Plans & pricing\n/reset — Clear history\n/help — Help\n\n⭐ = Standard or Premium\n💎 = Premium only",
        "limit_reached": "❌ Daily limit reached! Use /updateplan to upgrade.",
        "blocked": "🚫 You are blocked. Contact admin.",
        "history_cleared": "Chat history cleared ✅",
        "generating_voice": "⏳ Generating voice...",
        "generating_image": "⏳ Generating image, please wait...",
        "voice_example": "Example: /ai_sound Hello, how are you?",
        "imagine_example": "Example: /imagine a beautiful sunset over mountains",
        "translate_prompt": "Send me the text you want to translate and target language.\nExample: /translate Hello → Uzbek",
        "code_prompt": "Describe what code you need.\nExample: /code Python function to sort a list",
        "document_prompt": "What document do you need? (CV, contract, application, etc.)\nExample: /document job application letter for software developer",
        "choose_language": "🌐 Choose your language:",
        "language_changed": "✅ Language changed to English!",
        "language_cooldown": "⏳ You can change language once per 24 hours.",
        "pdf_required": "Please send a PDF file! 📄",
        "reading_pdf": "⏳ Reading PDF...",
        "pdf_locked": "⭐ PDF requires Standard or Premium!\nUse /updateplan to upgrade.",
        "voice_locked": "💎 Voice requires Premium!\nUse /updateplan to upgrade.",
        "pptx_locked": "💎 PowerPoint requires Premium!\nUse /updateplan to upgrade.",
        "word_locked": "💎 Word requires Premium!\nUse /updateplan to upgrade.",
        "cv_locked": "⭐ CV requires Standard or Premium!\nUse /updateplan to upgrade.",
        "email_locked": "⭐ Email requires Standard or Premium!\nUse /updateplan to upgrade.",
        "ai_sound_locked": "⭐ AI Sound requires Standard or Premium!\nUse /updateplan to upgrade.",
        "imagine_locked": "⭐ AI Image requires Standard or Premium!\nUse /updateplan to upgrade.",
        "document_locked": "⭐ Document creator requires Standard or Premium!\nUse /updateplan to upgrade.",
        "writing_cv": "⏳ Writing your CV...",
        "writing_email": "⏳ Writing email...",
        "writing_post": "⏳ Writing marketing post...",
        "writing_biznes": "⏳ Writing business plan...",
        "writing_document": "⏳ Creating document...",
        "translating": "⏳ Translating...",
        "writing_code": "⏳ Writing code...",
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
        "buy_standard_stars": f"⭐ Buy Standard — {STARS_STANDARD}⭐",
        "buy_premium_stars": f"💎 Buy Premium — {STARS_PREMIUM}⭐",
        "buy_group": f"👥 Buy Group Mode — {STARS_GROUP}⭐ ($99.9)",
        "upgrade_standard": "⭐ Upgrade to Standard",
        "upgrade_premium": "💎 Upgrade to Premium",
        "upgrade_plan": "💰 Upgrade Plan",
        "contact_admin": "💬 Contact Admin",
        "pay_with_stars": "⭐ Pay with Stars",
        "pay_with_card": "💳 Pay with Card",
        "you_said": "🎤 You said: ",
        "referral_title": "👥 Referral Program\n\n🔗 Your invite link:\n{link}\n\n📊 Invited: {count} friends\n\n🎁 Rewards:\n• 10 friends → ⭐ Standard 15 days — {standard_status}\n• 30 friends → 💎 Premium 15 days — {premium_status}\n\nShare this link! 🚀",
        "claim_standard": "🎁 Claim Standard 15 days",
        "claim_premium": "🎁 Claim Premium 15 days",
        "claim_standard_success": "🎉 Congratulations!\n⭐ Standard plan activated for 15 days!",
        "claim_premium_success": "🎉 Congratulations!\n💎 Premium plan activated for 15 days!",
        "claim_error": "❌ Already claimed or not enough referrals!",
        "stats_title": "📊 Your Statistics",
        "gift_already": "🎁 You already claimed today's gift! Come back tomorrow.",
        "gift_success": "🎁 Daily bonus claimed!\nStreak: {streak} days 🔥\n\n+5 extra messages today!",
        "gift_streak": "🎉 Amazing! 10 days streak!\n⭐ Standard plan activated for 10 days!",
        "top_title": "🏆 Top 10 Most Active Users",
        "payment_info": "💳 Pay to card:\n`{card}`\n\n💵 Amount: {amount} USDT\n\n📋 After payment:\n1. Screenshot\n2. Send to admin\n3. Activated within 1 hour ✅",
        "promo_enter": "Enter your promo code:",
        "promo_invalid": "❌ Invalid or expired promo code!",
        "promo_used": "❌ This promo code already used!",
        "promo_success": "✅ Promo code applied!\n{reward}",
        "group_info": "👥 Group Mode\n\nActivate bot for your group!\n\nPrice: {stars}⭐ ($99.9)\n\nAfter activation:\n• All group members can use the bot\n• Unlimited messages in group\n• AI responses in group chat",
    },
    "uz": {
        "welcome": "Salom! Men Chatbot 🤖\nSizning tarifingiz: {plan_emoji} {plan}\n\n💬 Men bilan suhbatlashing\n🌐 Yangiliklar, narxlar, ob-havo\n📄 PDF tahlil uchun yuboring\n🖼️ Rasm yuboring\n🎤 Ovozli xabar yuboring\n\n📊 /pptx — PowerPoint\n📝 /word — Word hujjat\n👤 /cv — CV yozish\n📧 /email — Email yozish\n📱 /post — Marketing post\n🔊 /ai_sound — AI Ovoz\n🎨 /imagine — AI Rasm\n🌐 /translate — Tarjimon\n💻 /code — Kod yozuvchi\n📋 /document — Hujjat yaratuvchi\n👥 /referral — Do'stlarni taklif\n🎁 /gift — Kunlik bonus\n🏆 /top — Top foydalanuvchilar\n📊 /stats — Statistika\n\n💰 /updateplan — Tarifni yangilash\n🌐 /language — Tilni o'zgartirish\n/help — Yordam\n/reset — Tarixni tozalash",
        "help": "📌 Buyruqlar:\n\n/pptx — PowerPoint 💎\n/word — Word hujjat 💎\n/cv — CV ⭐\n/email — Email ⭐\n/post — Marketing post\n/biznes — Biznes reja\n/ai_sound — AI Ovoz 🔊 ⭐\n/imagine — AI Rasm 🎨 ⭐\n/translate — Tarjimon 🌐\n/code — Kod yozuvchi 💻\n/document — Hujjat yaratuvchi 📋 ⭐\n/referral — Do'stlarni taklif 👥\n/gift — Kunlik bonus 🎁\n/top — Top foydalanuvchilar 🏆\n/stats — Statistika 📊\n/language — Tilni o'zgartirish 🌐\n/updateplan — Tariflar\n/reset — Tarixni tozalash\n/help — Yordam\n\n⭐ = Standart yoki Premium\n💎 = Faqat Premium",
        "limit_reached": "❌ Kunlik limit tugadi! /updateplan orqali yangilang.",
        "blocked": "🚫 Siz bloklangansiz. Admin bilan bog'laning.",
        "history_cleared": "Suhbat tarixi tozalandi ✅",
        "generating_voice": "⏳ Ovoz yaratilmoqda...",
        "generating_image": "⏳ Rasm yaratilmoqda, kuting...",
        "voice_example": "Misol: /ai_sound Salom, qandaysiz?",
        "imagine_example": "Misol: /imagine tog'lar ustidagi chiroyli quyosh botishi",
        "translate_prompt": "Tarjima qilmoqchi bo'lgan matn va tilni yuboring.\nMisol: /translate Hello → O'zbek",
        "code_prompt": "Qanday kod kerakligini tasvirlab bering.\nMisol: /code Python ro'yxatni saralash funksiyasi",
        "document_prompt": "Qanday hujjat kerak? (CV, shartnoma, ariza va h.k.)\nMisol: /document dasturchi uchun ish arizasi",
        "choose_language": "🌐 Tilni tanlang:",
        "language_changed": "✅ Til o'zbekchaga o'zgartirildi!",
        "language_cooldown": "⏳ Tilni 24 soatda 1 marta o'zgartirish mumkin.",
        "pdf_required": "Iltimos, PDF fayl yuboring! 📄",
        "reading_pdf": "⏳ PDF o'qilmoqda...",
        "pdf_locked": "⭐ PDF uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "voice_locked": "💎 Ovoz uchun Premium kerak!\n/updateplan orqali yangilang.",
        "pptx_locked": "💎 PowerPoint uchun Premium kerak!\n/updateplan orqali yangilang.",
        "word_locked": "💎 Word uchun Premium kerak!\n/updateplan orqali yangilang.",
        "cv_locked": "⭐ CV uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "email_locked": "⭐ Email uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "ai_sound_locked": "⭐ AI Ovoz uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "imagine_locked": "⭐ AI Rasm uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "document_locked": "⭐ Hujjat yaratuvchi uchun Standart yoki Premium kerak!\n/updateplan orqali yangilang.",
        "writing_cv": "⏳ CV yozilmoqda...",
        "writing_email": "⏳ Email yozilmoqda...",
        "writing_post": "⏳ Marketing post yozilmoqda...",
        "writing_biznes": "⏳ Biznes reja yozilmoqda...",
        "writing_document": "⏳ Hujjat yaratilmoqda...",
        "translating": "⏳ Tarjima qilinmoqda...",
        "writing_code": "⏳ Kod yozilmoqda...",
        "pptx_example": "Misol: /pptx sun'iy intellekt",
        "word_example": "Misol: /word biznes reja",
        "cv_example": "Misol: /cv Python dasturchi, 3 yil tajriba",
        "email_example": "Misol: /email intervyudan keyin",
        "post_example": "Misol: /post yangi kofe do'kon ochilishi",
        "biznes_example": "Misol: /biznes online kiyim do'koni",
        "creating_pptx": "⏳ '{topic}' mavzusida prezentatsiya yaratilmoqda...",
        "creating_word": "⏳ '{topic}' mavzusida hujjat yaratilmoqda...",
        "ready": "✅ Tayyor!",
        "buy_standard": "⭐ Standart — 5 USDT/oy",
        "buy_premium": "💎 Premium — 10 USDT/oy",
        "buy_standard_stars": f"⭐ Standart — {STARS_STANDARD}⭐",
        "buy_premium_stars": f"💎 Premium — {STARS_PREMIUM}⭐",
        "buy_group": f"👥 Guruh rejimi — {STARS_GROUP}⭐ ($99.9)",
        "upgrade_standard": "⭐ Standartga o'tish",
        "upgrade_premium": "💎 Premiumga o'tish",
        "upgrade_plan": "💰 Tarifni yangilash",
        "contact_admin": "💬 Admin bilan bog'lanish",
        "pay_with_stars": "⭐ Yulduzlar bilan to'lash",
        "pay_with_card": "💳 Karta bilan to'lash",
        "you_said": "🎤 Siz dedingiz: ",
        "referral_title": "👥 Referal Dasturi\n\n🔗 Sizning havola:\n{link}\n\n📊 Taklif qilingan: {count} do'st\n\n🎁 Mukofotlar:\n• 10 do'st → ⭐ Standart 15 kun — {standard_status}\n• 30 do'st → 💎 Premium 15 kun — {premium_status}\n\nHavolani ulashing! 🚀",
        "claim_standard": "🎁 Standart 15 kunni olish",
        "claim_premium": "🎁 Premium 15 kunni olish",
        "claim_standard_success": "🎉 Tabriklaymiz!\n⭐ Standart tarif 15 kunga yoqildi!",
        "claim_premium_success": "🎉 Tabriklaymiz!\n💎 Premium tarif 15 kunga yoqildi!",
        "claim_error": "❌ Allaqachon olинган yoki yetarli referal yo'q!",
        "stats_title": "📊 Sizning statistikangiz",
        "gift_already": "🎁 Bugun bonus allaqachon olindi! Ertaga qaytib keling.",
        "gift_success": "🎁 Kunlik bonus olindi!\nStreak: {streak} kun 🔥\n\nBugun +5 qo'shimcha xabar!",
        "gift_streak": "🎉 Ajoyib! 10 kun ketma-ket!\n⭐ Standart tarif 10 kunga yoqildi!",
        "top_title": "🏆 Top 10 Eng Faol Foydalanuvchilar",
        "payment_info": "💳 Kartaga to'lang:\n`{card}`\n\n💵 Miqdor: {amount} USDT\n\n📋 To'lovdan keyin:\n1. Skrinshot\n2. Adminga yuboring\n3. 1 soat ichida yoqiladi ✅",
        "promo_enter": "Promo kodingizni kiriting:",
        "promo_invalid": "❌ Noto'g'ri yoki muddati o'tgan promo kod!",
        "promo_used": "❌ Bu promo kod allaqachon ishlatilgan!",
        "promo_success": "✅ Promo kod qo'llandi!\n{reward}",
        "group_info": "👥 Guruh Rejimi\n\nBotni guruhingiz uchun yoqing!\n\nNarx: {stars}⭐ ($99.9)\n\nYoqilgandan keyin:\n• Barcha guruh a'zolari botdan foydalana oladi\n• Guruhda cheksiz xabarlar\n• Guruh chatida AI javoblar",
    },
    "ru": {
        "welcome": "Привет! Я Chatbot 🤖\nВаш тариф: {plan_emoji} {plan}\n\n💬 Общайтесь со мной\n🌐 Новости, цены, погода\n📄 Отправьте PDF\n🖼️ Изображение\n🎤 Голосовое сообщение\n\n📊 /pptx — PowerPoint\n📝 /word — Word\n👤 /cv — Резюме\n📧 /email — Email\n📱 /post — Пост\n🔊 /ai_sound — AI Голос\n🎨 /imagine — AI Изображение\n🌐 /translate — Переводчик\n💻 /code — Написать код\n📋 /document — Создать документ\n👥 /referral — Пригласить друзей\n🎁 /gift — Ежедневный бонус\n🏆 /top — Топ пользователей\n📊 /stats — Статистика\n\n💰 /updateplan — Тариф\n🌐 /language — Язык\n/help — Помощь\n/reset — Очистить историю",
        "help": "📌 Команды:\n\n/pptx — PowerPoint 💎\n/word — Word 💎\n/cv — Резюме ⭐\n/email — Email ⭐\n/post — Пост\n/biznes — Бизнес-план\n/ai_sound — AI Голос 🔊 ⭐\n/imagine — AI Изображение 🎨 ⭐\n/translate — Переводчик 🌐\n/code — Код 💻\n/document — Документ 📋 ⭐\n/referral — Друзья 👥\n/gift — Бонус 🎁\n/top — Топ 🏆\n/stats — Статистика 📊\n/language — Язык 🌐\n/updateplan — Тарифы\n/reset — История\n/help — Помощь\n\n⭐ = Standard или Premium\n💎 = Только Premium",
        "limit_reached": "❌ Дневной лимит! Используйте /updateplan.",
        "blocked": "🚫 Вы заблокированы.",
        "history_cleared": "История очищена ✅",
        "generating_voice": "⏳ Генерация голоса...",
        "generating_image": "⏳ Генерация изображения...",
        "voice_example": "Пример: /ai_sound Привет, как дела?",
        "imagine_example": "Пример: /imagine красивый закат над горами",
        "translate_prompt": "Отправьте текст и язык перевода.\nПример: /translate Hello → Русский",
        "code_prompt": "Опишите нужный код.\nПример: /code Python функция сортировки",
        "document_prompt": "Какой документ нужен?\nПример: /document заявление на работу",
        "choose_language": "🌐 Выберите язык:",
        "language_changed": "✅ Язык изменён на русский!",
        "language_cooldown": "⏳ Язык можно менять раз в 24 часа.",
        "pdf_required": "Отправьте PDF файл! 📄",
        "reading_pdf": "⏳ Читаю PDF...",
        "pdf_locked": "⭐ PDF требует Standard или Premium!",
        "voice_locked": "💎 Голос требует Premium!",
        "pptx_locked": "💎 PowerPoint требует Premium!",
        "word_locked": "💎 Word требует Premium!",
        "cv_locked": "⭐ Резюме требует Standard или Premium!",
        "email_locked": "⭐ Email требует Standard или Premium!",
        "ai_sound_locked": "⭐ AI Голос требует Standard или Premium!",
        "imagine_locked": "⭐ AI Изображение требует Standard или Premium!",
        "document_locked": "⭐ Документ требует Standard или Premium!",
        "writing_cv": "⏳ Пишу резюме...",
        "writing_email": "⏳ Пишу email...",
        "writing_post": "⏳ Пишу пост...",
        "writing_biznes": "⏳ Пишу бизнес-план...",
        "writing_document": "⏳ Создаю документ...",
        "translating": "⏳ Перевожу...",
        "writing_code": "⏳ Пишу код...",
        "pptx_example": "Пример: /pptx искусственный интеллект",
        "word_example": "Пример: /word бизнес-план",
        "cv_example": "Пример: /cv Python разработчик",
        "email_example": "Пример: /email после собеседования",
        "post_example": "Пример: /post открытие кофейни",
        "biznes_example": "Пример: /biznes онлайн магазин",
        "creating_pptx": "⏳ Создаю презентацию '{topic}'...",
        "creating_word": "⏳ Создаю документ '{topic}'...",
        "ready": "✅ Готово!",
        "buy_standard": "⭐ Standard — 5 USDT/месяц",
        "buy_premium": "💎 Premium — 10 USDT/месяц",
        "buy_standard_stars": f"⭐ Standard — {STARS_STANDARD}⭐",
        "buy_premium_stars": f"💎 Premium — {STARS_PREMIUM}⭐",
        "buy_group": f"👥 Групповой режим — {STARS_GROUP}⭐ ($99.9)",
        "upgrade_standard": "⭐ Перейти на Standard",
        "upgrade_premium": "💎 Перейти на Premium",
        "upgrade_plan": "💰 Обновить тариф",
        "contact_admin": "💬 Связаться с админом",
        "pay_with_stars": "⭐ Оплатить звёздами",
        "pay_with_card": "💳 Оплатить картой",
        "you_said": "🎤 Вы сказали: ",
        "referral_title": "👥 Реферальная программа\n\n🔗 Ваша ссылка:\n{link}\n\n📊 Приглашено: {count}\n\n🎁 Награды:\n• 10 друзей → ⭐ Standard 15 дней — {standard_status}\n• 30 друзей → 💎 Premium 15 дней — {premium_status}\n\nПоделитесь! 🚀",
        "claim_standard": "🎁 Получить Standard 15 дней",
        "claim_premium": "🎁 Получить Premium 15 дней",
        "claim_standard_success": "🎉 Поздравляем!\n⭐ Standard активирован на 15 дней!",
        "claim_premium_success": "🎉 Поздравляем!\n💎 Premium активирован на 15 дней!",
        "claim_error": "❌ Уже получено или недостаточно рефералов!",
        "stats_title": "📊 Ваша статистика",
        "gift_already": "🎁 Бонус уже получен! Возвращайтесь завтра.",
        "gift_success": "🎁 Бонус получен!\nСтрик: {streak} дней 🔥\n\n+5 сообщений сегодня!",
        "gift_streak": "🎉 10 дней подряд!\n⭐ Standard активирован на 10 дней!",
        "top_title": "🏆 Топ 10 Активных Пользователей",
        "payment_info": "💳 Оплатите на карту:\n`{card}`\n\n💵 Сумма: {amount} USDT\n\n📋 После оплаты:\n1. Скриншот\n2. Отправьте админу\n3. Активация в течение 1 часа ✅",
        "promo_enter": "Введите промокод:",
        "promo_invalid": "❌ Неверный или истёкший промокод!",
        "promo_used": "❌ Промокод уже использован!",
        "promo_success": "✅ Промокод применён!\n{reward}",
        "group_info": "👥 Групповой режим\n\nАктивируйте бота для группы!\n\nЦена: {stars}⭐ ($99.9)\n\nПосле активации:\n• Все участники могут пользоваться\n• Безлимитные сообщения\n• AI ответы в чате",
    },
}

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

def detect_lang_gtts(text):
    try:
        lang = detect(text)
        return GTTS_LANG_MAP.get(lang, "en")
    except:
        return "en"

def generate_referral_code():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))

def generate_promo_code():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=10))

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
            streak_count INTEGER DEFAULT 0,
            last_gift_claim DATE,
            streak_claimed BOOLEAN DEFAULT FALSE,
            group_plan TEXT DEFAULT 'none',
            total_messages INTEGER DEFAULT 0,
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
    c.execute("""
        CREATE TABLE IF NOT EXISTS promo_codes (
            id SERIAL PRIMARY KEY,
            code TEXT UNIQUE,
            reward_type TEXT,
            reward_value INTEGER,
            max_uses INTEGER DEFAULT 1,
            used_count INTEGER DEFAULT 0,
            expires_at TIMESTAMP,
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS promo_uses (
            user_id BIGINT,
            code TEXT,
            PRIMARY KEY (user_id, code)
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
                   language, language_changed_at, streak_count, last_gift_claim,
                   streak_claimed, group_plan, total_messages
            FROM users WHERE user_id = %s
        """, (user_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return {"plan": "free", "expires_at": None, "is_blocked": False,
                    "full_name": None, "username": None, "referral_code": None,
                    "referral_count": 0, "claimed_standard": False, "claimed_premium": False,
                    "language": "en", "language_changed_at": None, "streak_count": 0,
                    "last_gift_claim": None, "streak_claimed": False, "group_plan": "none",
                    "total_messages": 0}
        plan, expires_at, is_blocked, full_name, username, referral_code, referral_count, \
        claimed_standard, claimed_premium, language, language_changed_at, streak_count, \
        last_gift_claim, streak_claimed, group_plan, total_messages = row
        if expires_at and datetime.now() > expires_at and plan != 'free':
            set_plan(user_id, "free", None)
            plan = "free"
            expires_at = None
        return {
            "plan": plan, "expires_at": expires_at, "is_blocked": is_blocked,
            "full_name": full_name, "username": username, "referral_code": referral_code,
            "referral_count": referral_count or 0, "claimed_standard": claimed_standard or False,
            "claimed_premium": claimed_premium or False, "language": language or "en",
            "language_changed_at": language_changed_at, "streak_count": streak_count or 0,
            "last_gift_claim": last_gift_claim, "streak_claimed": streak_claimed or False,
            "group_plan": group_plan or "none", "total_messages": total_messages or 0
        }
    except:
        return {"plan": "free", "expires_at": None, "is_blocked": False,
                "full_name": None, "username": None, "referral_code": None,
                "referral_count": 0, "claimed_standard": False, "claimed_premium": False,
                "language": "en", "language_changed_at": None, "streak_count": 0,
                "last_gift_claim": None, "streak_claimed": False, "group_plan": "none",
                "total_messages": 0}

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

def increment_messages(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("UPDATE users SET total_messages = total_messages + 1 WHERE user_id=%s", (user_id,))
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
        return {k: -1 for k in ["chat","search","image","post","biznes","pdf","cv","email","voice","pptx","word","tts","imagine","translate","code","document"]}
    elif plan == "standard":
        return {
            "chat": 30, "search": 30, "image": 30, "post": 30, "biznes": 30,
            "pdf": 30, "cv": 30, "email": 30, "voice": 0, "pptx": 0, "word": 0,
            "tts": 30, "imagine": 30, "translate": 30, "code": 30, "document": 30
        }
    else:
        return {
            "chat": 20, "search": 20, "image": 20, "post": 20, "biznes": 20,
            "pdf": 0, "cv": 0, "email": 0, "voice": 0, "pptx": 0, "word": 0,
            "tts": 0, "imagine": 0, "translate": 20, "code": 20, "document": 0
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

async def generate_image_gemini(prompt):
    model = genai.ImageGenerationModel("imagen-3.0-generate-002")
    result = model.generate_images(prompt=prompt, number_of_images=1)
    if result.images:
        img_bytes = result.images[0]._image_bytes
        return img_bytes
    return None

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

async def gift_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    today = date.today()
    last_claim = u["last_gift_claim"]
    if last_claim and last_claim >= today:
        await update.message.reply_text(get_text(lang, "gift_already"))
        return
    streak = u["streak_count"]
    if last_claim and (today - last_claim).days == 1:
        streak += 1
    else:
        streak = 1
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            UPDATE users SET streak_count=%s, last_gift_claim=%s,
            usage_chat = usage_chat + 5
            WHERE user_id=%s
        """, (streak, today, user.id))
        conn.commit()
        conn.close()
    except:
        pass
    if streak >= 10 and not u["streak_claimed"]:
        set_plan(user.id, "standard", 10)
        try:
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE users SET streak_claimed=TRUE WHERE user_id=%s", (user.id,))
            conn.commit()
            conn.close()
        except:
            pass
        await update.message.reply_text(get_text(lang, "gift_streak"))
    else:
        await update.message.reply_text(get_text(lang, "gift_success", streak=streak))

async def top_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            SELECT username, full_name, total_messages, plan
            FROM users WHERE is_blocked=FALSE
            ORDER BY total_messages DESC LIMIT 10
        """)
        rows = c.fetchall()
        conn.close()
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        text = f"{get_text(lang, 'top_title')}\n{'─' * 28}\n\n"
        medals = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"]
        for i, row in enumerate(rows):
            username, full_name, total_msgs, plan = row
            name = f"@{username}" if username else (full_name or "Anonymous")
            emoji = plan_emoji.get(plan, "🆓")
            text += f"{medals[i]} {name} {emoji} — {total_msgs} messages\n"
        await update.message.reply_text(text)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def promo_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if context.args:
        code = context.args[0].upper()
        try:
            conn = get_conn()
            c = conn.cursor()
            c.execute("""
                SELECT reward_type, reward_value, max_uses, used_count, expires_at
                FROM promo_codes WHERE code=%s
            """, (code,))
            promo = c.fetchone()
            if not promo:
                await update.message.reply_text(get_text(lang, "promo_invalid"))
                conn.close()
                return
            reward_type, reward_value, max_uses, used_count, expires_at = promo
            if expires_at and datetime.now() > expires_at:
                await update.message.reply_text(get_text(lang, "promo_invalid"))
                conn.close()
                return
            if used_count >= max_uses:
                await update.message.reply_text(get_text(lang, "promo_invalid"))
                conn.close()
                return
            c.execute("SELECT 1 FROM promo_uses WHERE user_id=%s AND code=%s", (user.id, code))
            if c.fetchone():
                await update.message.reply_text(get_text(lang, "promo_used"))
                conn.close()
                return
            c.execute("INSERT INTO promo_uses (user_id, code) VALUES (%s, %s)", (user.id, code))
            c.execute("UPDATE promo_codes SET used_count=used_count+1 WHERE code=%s", (code,))
            conn.commit()
            conn.close()
            if reward_type == "standard":
                set_plan(user.id, "standard", reward_value)
                reward_text = f"⭐ Standard {reward_value} days!"
            elif reward_type == "premium":
                set_plan(user.id, "premium", reward_value)
                reward_text = f"💎 Premium {reward_value} days!"
            else:
                reward_text = f"+{reward_value} days!"
            await update.message.reply_text(get_text(lang, "promo_success", reward=reward_text))
        except Exception as e:
            await update.message.reply_text(f"Error: {str(e)}")
    else:
        await update.message.reply_text(get_text(lang, "promo_enter"))

async def createpromo_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    if len(context.args) < 3:
        await update.message.reply_text(
            "Usage: /createpromo [type] [days] [max_uses]\n"
            "Types: standard, premium\n"
            "Example: /createpromo standard 7 100"
        )
        return
    try:
        reward_type = context.args[0]
        reward_value = int(context.args[1])
        max_uses = int(context.args[2])
        code = generate_promo_code()
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO promo_codes (code, reward_type, reward_value, max_uses, expires_at)
            VALUES (%s, %s, %s, %s, %s)
        """, (code, reward_type, reward_value, max_uses, datetime.now() + timedelta(days=30)))
        conn.commit()
        conn.close()
        await update.message.reply_text(
            f"✅ Promo code created!\n\n"
            f"Code: `{code}`\n"
            f"Type: {reward_type}\n"
            f"Days: {reward_value}\n"
            f"Max uses: {max_uses}\n"
            f"Expires: 30 days",
            parse_mode="Markdown"
        )
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def group_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    keyboard = [[InlineKeyboardButton(get_text(lang, "buy_group"), callback_data="buy_group")]]
    await update.message.reply_text(
        get_text(lang, "group_info", stars=STARS_GROUP),
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

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
    await update.message.reply_text(get_text(lang, "choose_language"), reply_markup=InlineKeyboardMarkup(keyboard))

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
            f"{get_text(lang, 'stats_title')}\n{'─' * 28}\n\n"
            f"👤 Plan: {plan_emoji} {plan.upper()}\n"
            f"📅 Expires: {expires}\n"
            f"👥 Referrals: {u['referral_count']}\n"
            f"🔥 Streak: {u['streak_count']} days\n"
            f"💬 Total messages: {u['total_messages']}\n\n"
            f"📊 Today:\n"
            f"💬 Chat: {chat}\n🌐 Search: {search}\n🖼️ Image: {image}\n"
            f"📱 Post: {post}\n💼 Business: {biznes}\n📄 PDF: {pdf}\n"
            f"👤 CV: {cv}\n📧 Email: {email}\n🔊 Voice: {tts}"
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

async def translate_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    limits = get_limits(u["plan"])
    if not check_limit(user.id, "translate", limits["translate"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    text = " ".join(context.args)
    if not text:
        await update.message.reply_text(get_text(lang, "translate_prompt"))
        return
    await update.message.reply_text(get_text(lang, "translating"))
    try:
        reply = await ai_generate(f"Translate the following text as instructed. Keep the meaning exactly.\n\n{text}")
        await update.message.reply_text(f"🌐 {reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def code_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    limits = get_limits(u["plan"])
    if not check_limit(user.id, "code", limits["code"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    text = " ".join(context.args)
    if not text:
        await update.message.reply_text(get_text(lang, "code_prompt"))
        return
    await update.message.reply_text(get_text(lang, "writing_code"))
    try:
        reply = await ai_generate(f"Write clean, well-commented code for: {text}\nInclude explanation of how it works.")
        await update.message.reply_text(f"💻 {reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def document_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["document"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "document_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "document", limits["document"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    text = " ".join(context.args)
    if not text:
        await update.message.reply_text(get_text(lang, "document_prompt"))
        return
    await update.message.reply_text(get_text(lang, "writing_document"))
    try:
        reply = await ai_generate(
            f"Create a professional document: {text}\n"
            f"Format it properly with all necessary sections, formal language, and structure. "
            f"Make it ready to use. Same language as the request."
        )
        await update.message.reply_text(f"📋 {reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def imagine_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["imagine"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "imagine_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "imagine", limits["imagine"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    prompt = " ".join(context.args)
    if not prompt:
        await update.message.reply_text(get_text(lang, "imagine_example"))
        return
    await update.message.reply_text(get_text(lang, "generating_image"))
    try:
        img_bytes = await generate_image_gemini(prompt)
        if img_bytes:
            buf = io.BytesIO(img_bytes)
            buf.name = "image.png"
            await update.message.reply_photo(photo=buf, caption=f"🎨 {prompt}")
        else:
            await update.message.reply_text("❌ Could not generate image. Try again.")
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {str(e)}\nTry a different prompt.")

async def updateplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    plan = u["plan"]
    expires = u["expires_at"].strftime("%d.%m.%Y") if u.get("expires_at") else "—"
    plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}.get(plan, "🆓")
    keyboard = [
        [
            InlineKeyboardButton(get_text(lang, "buy_standard_stars"), callback_data="stars_standard"),
            InlineKeyboardButton(get_text(lang, "buy_premium_stars"), callback_data="stars_premium"),
        ],
        [
            InlineKeyboardButton(get_text(lang, "buy_standard"), callback_data="buy_standard"),
            InlineKeyboardButton(get_text(lang, "buy_premium"), callback_data="buy_premium"),
        ],
        [InlineKeyboardButton(get_text(lang, "buy_group"), callback_data="buy_group")],
    ]
    await update.message.reply_text(
        f"💰 Plans & Pricing\n{'─' * 28}\n\n"
        f"🆓 FREE\n• Chat: 20/day • Search: 20/day\n• Translate: 20/day • Code: 20/day\n\n"
        f"⭐ STANDARD — 5 USDT / {STARS_STANDARD}⭐\n• Everything: 30/day\n• PDF ✓ CV ✓ Email ✓ Voice ✓\n• AI Image ✓ Document ✓\n\n"
        f"💎 PREMIUM — 10 USDT / {STARS_PREMIUM}⭐\n• Everything: Unlimited\n• Voice messages ✓ PowerPoint ✓ Word ✓\n\n"
        f"👥 GROUP MODE — {STARS_GROUP}⭐ ($99.9)\n• Bot for your group\n\n"
        f"{'─' * 28}\n"
        f"👤 {plan_emoji} {plan.upper()}\n📅 {expires}",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def precheckout_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.pre_checkout_query
    await query.answer(ok=True)

async def successful_payment_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    payload = update.message.successful_payment.invoice_payload
    if payload == "standard_stars":
        set_plan(user.id, "standard", 30)
        await update.message.reply_text("✅ ⭐ Standard plan activated for 30 days!")
    elif payload == "premium_stars":
        set_plan(user.id, "premium", 30)
        await update.message.reply_text("✅ 💎 Premium plan activated for 30 days!")
    elif payload == "group_stars":
        try:
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE users SET group_plan='active' WHERE user_id=%s", (user.id,))
            conn.commit()
            conn.close()
        except:
            pass
        await update.message.reply_text("✅ 👥 Group mode activated!")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    u = get_user(user_id)
    lang = u["language"]

    if data == "stars_standard":
        await context.bot.send_invoice(
            chat_id=user_id,
            title="⭐ Standard Plan",
            description="Standard plan for 30 days",
            payload="standard_stars",
            currency="XTR",
            prices=[LabeledPrice("Standard 30 days", STARS_STANDARD)],
        )

    elif data == "stars_premium":
        await context.bot.send_invoice(
            chat_id=user_id,
            title="💎 Premium Plan",
            description="Premium plan for 30 days",
            payload="premium_stars",
            currency="XTR",
            prices=[LabeledPrice("Premium 30 days", STARS_PREMIUM)],
        )

    elif data == "buy_group" or data == "stars_group":
        await context.bot.send_invoice(
            chat_id=user_id,
            title="👥 Group Mode",
            description="Activate bot for your group",
            payload="group_stars",
            currency="XTR",
            prices=[LabeledPrice("Group Mode", STARS_GROUP)],
        )

    elif data == "setlang_" + data.split("_")[-1] if data.startswith("setlang_") else False:
        pass

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
            await context.bot.send_message(chat_id=target_id, text=f"{plan_emoji[plan]} Plan updated to {plan.upper()}!")
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
        f"/users — All users\n/find [id or @username] — Manage\n"
        f"/broadcast [text] — Message all\n"
        f"/createpromo [type] [days] [uses] — Create promo"
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
        await update.message.reply_text("Usage: /find [user_id or @username]")
        return
    try:
        arg = context.args[0]
        conn = get_conn()
        c = conn.cursor()
        if arg.startswith("@"):
            username = arg[1:]
            c.execute("SELECT plan, expires_at, is_blocked, full_name, username, user_id FROM users WHERE username = %s", (username,))
        else:
            c.execute("SELECT plan, expires_at, is_blocked, full_name, username, user_id FROM users WHERE user_id = %s", (int(arg),))
        row = c.fetchone()
        conn.close()
        if not row:
            await update.message.reply_text(f"❌ User {arg} not found!")
            return
        plan, expires_at, is_blocked, full_name, username, target_id = row
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
            f"🆔 ID: {target_id}\n👤 {full_name or '—'}\n"
            f"📱 @{username or '—'}\n"
            f"📋 {plan_emoji.get(plan,'🆓')} {plan.upper()}\n"
            f"📅 {expires}\n"
            f"{'🚫 Blocked' if is_blocked else '✅ Active'}",
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
    if u["is_blocked"]: return
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
            if content_str.startswith("json"): content_str = content_str[4:]
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
    if u["is_blocked"]: return
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
            if content_str.startswith("json"): content_str = content_str[4:]
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
    if u["is_blocked"]: return
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
    if u["is_blocked"]: return
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
    if u["is_blocked"]: return
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
    if u["is_blocked"]: return
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
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]: return
    limits = get_limits(u["plan"])
    if limits.get("tts") == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "ai_sound_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "tts", limits["tts"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
        return
    text = " ".join(context.args)
    if not text:
        await update.message.reply_text(get_text(lang, "voice_example"))
        return
    await update.message.reply_text(get_text(lang, "generating_voice"))
    try:
        tts_lang = detect_lang_gtts(text)
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
    increment_messages(user_id)
    user_text = update.message.text
    memory = get_memory(user_id)
    memory_context = ""
    if memory and (memory.get("name") or memory.get("facts")):
        memory_context = f"User info — name: {memory['name']}, facts: {memory['facts']}. "
    system_prompt = (
        "You are a professional AI assistant. "
        "VERY IMPORTANT: Always reply in the SAME language the user uses in their CURRENT message. "
        "If user writes in Uzbek → reply in Uzbek. English → English. Russian → Russian. "
        "If user asks to speak another language, do it immediately. "
        "Never switch languages on your own. Keep answers short, clear and natural. "
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
            message_content = f"User question: {user_text}\n\nWeb results:\n{search_content}\n\nAnswer in same language as question."
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
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]: return
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
        if user.id not in user_histories:
            user_histories[user.id] = [{"role": "system", "content": "You are a professional AI assistant. Always reply in the same language the user writes in."}]
        user_histories[user.id].append({"role": "user", "content": user_text})
        response = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=user_histories[user.id])
        reply = response.choices[0].message.content
        user_histories[user.id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]: return
    limits = get_limits(u["plan"])
    if not check_limit(user.id, "image", limits["image"]):
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
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]: return
    limits = get_limits(u["plan"])
    if limits["pdf"] == 0:
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_standard"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "pdf_locked"), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "pdf", limits["pdf"]):
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
        if user.id not in user_histories:
            user_histories[user.id] = [{"role": "system", "content": "You are a professional AI assistant. Always reply in the same language the user writes in."}]
        user_histories[user.id].append({"role": "user", "content": f"PDF:\n\n{text}\n\nRequest: {caption}"})
        response = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=user_histories[user.id])
        reply = response.choices[0].message.content
        user_histories[user.id].append({"role": "assistant", "content": reply})
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
        BotCommand("imagine", "AI Image (Standard+)"),
        BotCommand("translate", "Translator"),
        BotCommand("code", "Code writer"),
        BotCommand("document", "Document creator (Standard+)"),
        BotCommand("referral", "Invite friends & earn bonuses"),
        BotCommand("gift", "Daily bonus"),
        BotCommand("top", "Top users"),
        BotCommand("stats", "My statistics"),
        BotCommand("promo", "Enter promo code"),
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
    app.add_handler(CommandHandler("imagine", imagine_command))
    app.add_handler(CommandHandler("translate", translate_command))
    app.add_handler(CommandHandler("code", code_command))
    app.add_handler(CommandHandler("document", document_command))
    app.add_handler(CommandHandler("referral", referral_command))
    app.add_handler(CommandHandler("gift", gift_command))
    app.add_handler(CommandHandler("top", top_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("promo", promo_command))
    app.add_handler(CommandHandler("language", language_command))
    app.add_handler(CommandHandler("group", group_command))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("find", find_command))
    app.add_handler(CommandHandler("users", users_command))
    app.add_handler(CommandHandler("broadcast", broadcast_command))
    app.add_handler(CommandHandler("createpromo", createpromo_command))
    app.add_handler(PreCheckoutQueryHandler(precheckout_handler))
    app.add_handler(MessageHandler(filters.SUCCESSFUL_PAYMENT, successful_payment_handler))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.PHOTO, handle_image))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_document))
    print("Bot is running... ✅")
    app.run_polling()

if __name__ == "__main__":
    main()
