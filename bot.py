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
CARD_NUMBER = os.getenv("CARD_NUMBER")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

genai.configure(api_key=GEMINI_API_KEY)

ADMIN_ID = 8230883785
ADMIN_USERNAME = "temur_uzb7779"

STARS_STANDARD = 500
STARS_PREMIUM = 1000
STARS_GROUP = 7700
STARS_IMAGINE = 50

COINS_PER_MESSAGE = 5
COINS_STANDARD = 10000
COINS_PREMIUM = 100000
AFFILIATE_PERCENT = 25
MAX_WARNINGS = 10

BAD_WORDS = [
    "ублюдок", "сука", "пизда", "хуй", "мудак", "залупа", "пиздец", "блядь", "ёбаный",
    "stupid", "idiot", "fuck", "shit", "asshole", "bitch", "bastard", "dick", "cunt",
    "ahmoq", "tentak", "befayz", "harom", "yomon", "stupid", "idiot",
    "yaramas", "it", "eshak", "cho'chqa", "kaltak", "go'rso'tar", "haromzoda",
    "onangni", "otangni", "yaratganning", "egangni", "adminni", "seni"
]

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
    "hi": "hi"
}

LANGUAGES = {
    "en": ("🇬🇧", "English"), "ru": ("🇷🇺", "Russian"), "uz": ("🇺🇿", "Uzbek"),
    "tr": ("🇹🇷", "Turkish"), "de": ("🇩🇪", "German"), "fr": ("🇫🇷", "French"),
    "es": ("🇪🇸", "Spanish"), "ar": ("🇸🇦", "Arabic"), "ko": ("🇰🇷", "Korean"),
    "ja": ("🇯🇵", "Japanese"), "zh": ("🇨🇳", "Chinese"), "it": ("🇮🇹", "Italian"),
    "pt": ("🇵🇹", "Portuguese"), "hi": ("🇮🇳", "Hindi"),
}

TEXTS = {
    "en": {
        "welcome": "Hello! I am Chatbot 🤖\nYour plan: {plan_emoji} {plan}\n🪙 Coins: {coins}\n\n💬 Chat with me\n🌐 News, prices, weather\n📄 PDF analyze\n🖼️ Send image\n🎤 Voice message\n\n📊 /pptx — PowerPoint 💎\n📝 /word — Word 💎\n👤 /cv — CV ⭐\n📧 /email — Email ⭐\n📱 /post — Marketing post\n🔊 /ai_sound — AI Voice ⭐\n🎨 /imagine — AI Image 💎⭐\n🌐 /translate — Translator\n💻 /code — Code writer\n📋 /document — Document ⭐\n\n🪙 /coins — My coins\n👥 /referral — Invite friends\n🤝 /affiliate — Earn with affiliate\n🎁 /gift — Daily bonus\n🏆 /top — Top users\n📊 /stats — Statistics\n🎟️ /promo — Promo code\n\n💰 /updateplan — Plans\n🌐 /language — Language\n/help — Help\n/reset — Clear history",
        "help": "📌 Commands:\n\n/pptx — PowerPoint 💎\n/word — Word 💎\n/cv — CV ⭐\n/email — Email ⭐\n/post — Marketing post\n/biznes — Business plan\n/ai_sound — AI Voice ⭐\n/imagine — AI Image 💎\n/translate — Translator\n/code — Code writer\n/document — Document ⭐\n/coins — My coins 🪙\n/referral — Invite friends\n/affiliate — Affiliate program\n/gift — Daily bonus\n/top — Top users\n/stats — Statistics\n/promo — Promo code\n/updateplan — Plans\n/language — Language\n/reset — Clear history\n/help — Help",
        "limit_reached": "❌ Daily limit reached! Use /updateplan.",
        "blocked": "🚫 You are blocked. Contact admin.",
        "history_cleared": "✅ Chat history cleared.",
        "generating_voice": "⏳ Generating voice...",
        "generating_image": "⏳ Generating image...",
        "voice_example": "Example: /ai_sound Hello!",
        "imagine_example": "Example: /imagine beautiful sunset",
        "imagine_locked": "💎 AI Image requires Premium!\n\nYou can also pay {stars}⭐ per image.",
        "translate_prompt": "Example: /translate Hello → Uzbek",
        "code_prompt": "Example: /code Python sort function",
        "document_prompt": "Example: /document job application letter",
        "choose_language": "🌐 Choose language:",
        "language_changed": "✅ Language changed!",
        "language_cooldown": "⏳ Change language once per 24 hours.",
        "pdf_required": "Send a PDF file! 📄",
        "reading_pdf": "⏳ Reading PDF...",
        "pdf_locked": "⭐ PDF requires Standard or Premium!",
        "voice_locked": "💎 Voice requires Premium!",
        "pptx_locked": "💎 PowerPoint requires Premium!",
        "word_locked": "💎 Word requires Premium!",
        "cv_locked": "⭐ CV requires Standard or Premium!",
        "email_locked": "⭐ Email requires Standard or Premium!",
        "ai_sound_locked": "⭐ AI Sound requires Standard or Premium!",
        "document_locked": "⭐ Document requires Standard or Premium!",
        "writing_cv": "⏳ Writing CV...",
        "writing_email": "⏳ Writing email...",
        "writing_post": "⏳ Writing post...",
        "writing_biznes": "⏳ Writing business plan...",
        "writing_document": "⏳ Creating document...",
        "translating": "⏳ Translating...",
        "writing_code": "⏳ Writing code...",
        "pptx_example": "Example: /pptx AI topic",
        "word_example": "Example: /word business plan",
        "cv_example": "Example: /cv Python dev 3 years",
        "email_example": "Example: /email follow up",
        "post_example": "Example: /post coffee shop opening",
        "biznes_example": "Example: /biznes online store",
        "creating_pptx": "⏳ Creating '{topic}'...",
        "creating_word": "⏳ Creating '{topic}'...",
        "ready": "✅ Ready!",
        "buy_standard": "⭐ Standard — 5 USDT/month",
        "buy_premium": "💎 Premium — 10 USDT/month",
        "buy_standard_stars": f"⭐ Standard — {STARS_STANDARD}⭐",
        "buy_premium_stars": f"💎 Premium — {STARS_PREMIUM}⭐",
        "buy_standard_coins": f"⭐ Standard — {COINS_STANDARD:,} coins",
        "buy_premium_coins": f"💎 Premium — {COINS_PREMIUM:,} coins",
        "buy_group": f"👥 Group Mode — {STARS_GROUP}⭐",
        "upgrade_standard": "⭐ Upgrade to Standard",
        "upgrade_premium": "💎 Upgrade to Premium",
        "upgrade_plan": "💰 Upgrade Plan",
        "contact_admin": "💬 Contact Admin",
        "you_said": "🎤 You said: ",
        "coins_info": "🪙 Your Coins: {coins}\n\n📈 Earn coins:\n• 5 coins per message\n• 25% bonus from referrals\n\n🛒 Spend coins:\n• {standard:,} coins = ⭐ Standard 1 month\n• {premium:,} coins = 💎 Premium 1 month",
        "coins_standard_btn": f"Buy Standard ({COINS_STANDARD:,} coins)",
        "coins_premium_btn": f"Buy Premium ({COINS_PREMIUM:,} coins)",
        "not_enough_coins": "❌ Not enough coins! You need {need:,} more coins.",
        "coins_success_standard": "✅ Standard plan activated for 30 days!",
        "coins_success_premium": "✅ Premium plan activated for 30 days!",
        "referral_title": "👥 Referral Program\n\n🔗 Your link:\n{link}\n\n📊 Invited: {count}\n\n🎁 Rewards:\n• 10 friends → ⭐ Standard 15 days — {standard_status}\n• 30 friends → 💎 Premium 15 days — {premium_status}\n\n🪙 You earn 25% of your referrals' coins!",
        "claim_standard": "🎁 Claim Standard 15 days",
        "claim_premium": "🎁 Claim Premium 15 days",
        "claim_standard_success": "🎉 ⭐ Standard activated for 15 days!",
        "claim_premium_success": "🎉 💎 Premium activated for 15 days!",
        "claim_error": "❌ Already claimed or not enough referrals!",
        "affiliate_info": "🤝 Affiliate Program\n\n🔗 Your affiliate link:\n{link}\n\n📊 Referrals: {count}\n🪙 Earned: {earned} coins\n\n💡 How it works:\n• Share your link\n• When someone subscribes, you earn 25% of their coins\n• Withdraw coins for plan upgrades",
        "stats_title": "📊 Your Statistics",
        "gift_already": "🎁 Already claimed today! Come back tomorrow.",
        "gift_success": "🎁 Daily bonus!\nStreak: {streak} days 🔥\n+5 messages & +25 coins!",
        "gift_streak": "🎉 10 days streak!\n⭐ Standard activated for 10 days!",
        "top_title": "🏆 Top 10 Active Users",
        "payment_info": "💳 Card:\n`{card}`\n\n💵 Amount: {amount} USDT\n\n1. Screenshot\n2. Send to admin\n3. Activated in 1 hour ✅",
        "promo_enter": "Enter promo code: /promo CODE",
        "promo_invalid": "❌ Invalid promo code!",
        "promo_used": "❌ Already used!",
        "promo_success": "✅ Promo applied! {reward}",
        "group_info": "👥 Group Mode — {stars}⭐\n\nActivate for your group!\n• All members use the bot\n• Unlimited messages",
        "warning": "⚠️ Warning {count}/{max}: Please be respectful!",
        "banned_for_abuse": "🚫 You have been banned for abuse.",
        "achievement_100k": "🏆 Achievement: 100,000 messages!\n⭐ Standard activated for 5 days!",
        "imagine_pay": f"🎨 Pay {STARS_IMAGINE}⭐ for one image",
    },
    "uz": {
        "welcome": "Salom! Men Chatbot 🤖\nTarifingiz: {plan_emoji} {plan}\n🪙 Coinlar: {coins}\n\n💬 Men bilan suhbatlashing\n🌐 Yangiliklar, narxlar, ob-havo\n📄 PDF tahlil\n🖼️ Rasm\n🎤 Ovozli xabar\n\n📊 /pptx — PowerPoint 💎\n📝 /word — Word 💎\n👤 /cv — CV ⭐\n📧 /email — Email ⭐\n📱 /post — Marketing post\n🔊 /ai_sound — AI Ovoz ⭐\n🎨 /imagine — AI Rasm 💎\n🌐 /translate — Tarjimon\n💻 /code — Kod yozuvchi\n📋 /document — Hujjat ⭐\n\n🪙 /coins — Coinlarim\n👥 /referral — Do'stlarni taklif\n🤝 /affiliate — Affiliate dasturi\n🎁 /gift — Kunlik bonus\n🏆 /top — Top foydalanuvchilar\n📊 /stats — Statistika\n🎟️ /promo — Promo kod\n\n💰 /updateplan — Tariflar\n🌐 /language — Til\n/help — Yordam\n/reset — Tarixni tozalash",
        "help": "📌 Buyruqlar:\n\n/pptx — PowerPoint 💎\n/word — Word 💎\n/cv — CV ⭐\n/email — Email ⭐\n/post — Post\n/biznes — Biznes reja\n/ai_sound — AI Ovoz ⭐\n/imagine — AI Rasm 💎\n/translate — Tarjimon\n/code — Kod\n/document — Hujjat ⭐\n/coins — Coinlar 🪙\n/referral — Referal\n/affiliate — Affiliate\n/gift — Bonus\n/top — Top\n/stats — Statistika\n/promo — Promo\n/updateplan — Tariflar\n/language — Til\n/reset — Tozalash\n/help — Yordam",
        "limit_reached": "❌ Kunlik limit tugadi! /updateplan.",
        "blocked": "🚫 Bloklangansiz. Admin bilan bog'laning.",
        "history_cleared": "✅ Tarix tozalandi.",
        "generating_voice": "⏳ Ovoz yaratilmoqda...",
        "generating_image": "⏳ Rasm yaratilmoqda...",
        "voice_example": "Misol: /ai_sound Salom!",
        "imagine_example": "Misol: /imagine chiroyli quyosh botishi",
        "imagine_locked": "💎 AI Rasm Premium uchun!\n\nYoki {stars}⭐ to'lab bir rasm olish mumkin.",
        "translate_prompt": "Misol: /translate Hello → O'zbek",
        "code_prompt": "Misol: /code Python saralash funksiyasi",
        "document_prompt": "Misol: /document ish arizasi",
        "choose_language": "🌐 Tilni tanlang:",
        "language_changed": "✅ Til o'zgartirildi!",
        "language_cooldown": "⏳ Tilni 24 soatda 1 marta o'zgartirish mumkin.",
        "pdf_required": "PDF fayl yuboring! 📄",
        "reading_pdf": "⏳ PDF o'qilmoqda...",
        "pdf_locked": "⭐ PDF uchun Standart yoki Premium kerak!",
        "voice_locked": "💎 Ovoz uchun Premium kerak!",
        "pptx_locked": "💎 PowerPoint uchun Premium kerak!",
        "word_locked": "💎 Word uchun Premium kerak!",
        "cv_locked": "⭐ CV uchun Standart yoki Premium kerak!",
        "email_locked": "⭐ Email uchun Standart yoki Premium kerak!",
        "ai_sound_locked": "⭐ AI Ovoz uchun Standart yoki Premium kerak!",
        "document_locked": "⭐ Hujjat uchun Standart yoki Premium kerak!",
        "writing_cv": "⏳ CV yozilmoqda...",
        "writing_email": "⏳ Email yozilmoqda...",
        "writing_post": "⏳ Post yozilmoqda...",
        "writing_biznes": "⏳ Biznes reja yozilmoqda...",
        "writing_document": "⏳ Hujjat yaratilmoqda...",
        "translating": "⏳ Tarjima qilinmoqda...",
        "writing_code": "⏳ Kod yozilmoqda...",
        "pptx_example": "Misol: /pptx sun'iy intellekt",
        "word_example": "Misol: /word biznes reja",
        "cv_example": "Misol: /cv Python dasturchi",
        "email_example": "Misol: /email intervyu",
        "post_example": "Misol: /post kofe do'kon",
        "biznes_example": "Misol: /biznes online do'kon",
        "creating_pptx": "⏳ '{topic}' yaratilmoqda...",
        "creating_word": "⏳ '{topic}' yaratilmoqda...",
        "ready": "✅ Tayyor!",
        "buy_standard": "⭐ Standart — 5 USDT/oy",
        "buy_premium": "💎 Premium — 10 USDT/oy",
        "buy_standard_stars": f"⭐ Standart — {STARS_STANDARD}⭐",
        "buy_premium_stars": f"💎 Premium — {STARS_PREMIUM}⭐",
        "buy_standard_coins": f"⭐ Standart — {COINS_STANDARD:,} coin",
        "buy_premium_coins": f"💎 Premium — {COINS_PREMIUM:,} coin",
        "buy_group": f"👥 Guruh — {STARS_GROUP}⭐",
        "upgrade_standard": "⭐ Standartga o'tish",
        "upgrade_premium": "💎 Premiumga o'tish",
        "upgrade_plan": "💰 Tarifni yangilash",
        "contact_admin": "💬 Admin",
        "you_said": "🎤 Siz dedingiz: ",
        "coins_info": "🪙 Sizning coinlaringiz: {coins}\n\n📈 Coin qanday yig'iladi:\n• Har xabar uchun 5 coin\n• Referallardan 25% bonus\n\n🛒 Coin sarflash:\n• {standard:,} coin = ⭐ Standart 1 oy\n• {premium:,} coin = 💎 Premium 1 oy",
        "coins_standard_btn": f"Standart sotib olish ({COINS_STANDARD:,} coin)",
        "coins_premium_btn": f"Premium sotib olish ({COINS_PREMIUM:,} coin)",
        "not_enough_coins": "❌ Coin yetarli emas! Yana {need:,} coin kerak.",
        "coins_success_standard": "✅ Standart tarif 30 kunga yoqildi!",
        "coins_success_premium": "✅ Premium tarif 30 kunga yoqildi!",
        "referral_title": "👥 Referal Dasturi\n\n🔗 Sizning havola:\n{link}\n\n📊 Taklif qilingan: {count}\n\n🎁 Mukofotlar:\n• 10 do'st → ⭐ Standart 15 kun — {standard_status}\n• 30 do'st → 💎 Premium 15 kun — {premium_status}\n\n🪙 Referallaringiz coinining 25%ini olasiz!",
        "claim_standard": "🎁 Standart 15 kun",
        "claim_premium": "🎁 Premium 15 kun",
        "claim_standard_success": "🎉 ⭐ Standart 15 kunga yoqildi!",
        "claim_premium_success": "🎉 💎 Premium 15 kunga yoqildi!",
        "claim_error": "❌ Allaqachon olingan yoki yetarli referal yo'q!",
        "affiliate_info": "🤝 Affiliate Dasturi\n\n🔗 Sizning havola:\n{link}\n\n📊 Referallar: {count}\n🪙 Ishlangan: {earned} coin\n\n💡 Qanday ishlaydi:\n• Havolangizni ulashing\n• Kimdir obuna bo'lsa, ularning coinining 25%ini olasiz",
        "stats_title": "📊 Sizning statistikangiz",
        "gift_already": "🎁 Bugun bonus olindi! Ertaga qaytib keling.",
        "gift_success": "🎁 Kunlik bonus!\nStreak: {streak} kun 🔥\n+5 xabar & +25 coin!",
        "gift_streak": "🎉 10 kun ketma-ket!\n⭐ Standart 10 kunga yoqildi!",
        "top_title": "🏆 Top 10 Faol Foydalanuvchilar",
        "payment_info": "💳 Karta:\n`{card}`\n\n💵 Miqdor: {amount} USDT\n\n1. Skrinshot\n2. Adminga yuboring\n3. 1 soatda yoqiladi ✅",
        "promo_enter": "Promo kodni kiriting: /promo KOD",
        "promo_invalid": "❌ Noto'g'ri promo kod!",
        "promo_used": "❌ Allaqachon ishlatilgan!",
        "promo_success": "✅ Promo kod qo'llandi! {reward}",
        "group_info": "👥 Guruh Rejimi — {stars}⭐\n\nGuruhingiz uchun yoqing!\n• Barcha a'zolar foydalana oladi\n• Cheksiz xabarlar",
        "warning": "⚠️ Ogohlantirish {count}/{max}: Iltimos, hurmat bilan muomala qiling!",
        "banned_for_abuse": "🚫 Haqorat uchun bloklandi.",
        "achievement_100k": "🏆 Yutuq: 100,000 xabar!\n⭐ Standart 5 kunga yoqildi!",
        "imagine_pay": f"🎨 Bir rasm uchun {STARS_IMAGINE}⭐ to'lang",
    },
    "ru": {
        "welcome": "Привет! Я Chatbot 🤖\nТариф: {plan_emoji} {plan}\n🪙 Монеты: {coins}\n\n💬 Общайтесь\n🌐 Новости, цены, погода\n📄 PDF анализ\n🖼️ Изображение\n🎤 Голос\n\n📊 /pptx — PowerPoint 💎\n📝 /word — Word 💎\n👤 /cv — Резюме ⭐\n📧 /email — Email ⭐\n📱 /post — Пост\n🔊 /ai_sound — AI Голос ⭐\n🎨 /imagine — AI Изображение 💎\n🌐 /translate — Переводчик\n💻 /code — Код\n📋 /document — Документ ⭐\n\n🪙 /coins — Мои монеты\n👥 /referral — Пригласить\n🤝 /affiliate — Аффилиат\n🎁 /gift — Бонус\n🏆 /top — Топ\n📊 /stats — Статистика\n🎟️ /promo — Промокод\n\n💰 /updateplan — Тарифы\n🌐 /language — Язык\n/help — Помощь\n/reset — Очистить",
        "help": "📌 Команды:\n\n/pptx 💎 /word 💎 /cv ⭐ /email ⭐\n/post /biznes /ai_sound ⭐ /imagine 💎\n/translate /code /document ⭐\n/coins /referral /affiliate /gift /top\n/stats /promo /updateplan /language\n/reset /help",
        "limit_reached": "❌ Лимит исчерпан! /updateplan",
        "blocked": "🚫 Вы заблокированы.",
        "history_cleared": "✅ История очищена.",
        "generating_voice": "⏳ Генерация голоса...",
        "generating_image": "⏳ Генерация изображения...",
        "voice_example": "Пример: /ai_sound Привет!",
        "imagine_example": "Пример: /imagine красивый закат",
        "imagine_locked": "💎 AI Изображение — Premium!\n\nИли {stars}⭐ за одно изображение.",
        "translate_prompt": "Пример: /translate Hello → Русский",
        "code_prompt": "Пример: /code Python сортировка",
        "document_prompt": "Пример: /document заявление на работу",
        "choose_language": "🌐 Выберите язык:",
        "language_changed": "✅ Язык изменён!",
        "language_cooldown": "⏳ Раз в 24 часа.",
        "pdf_required": "Отправьте PDF! 📄",
        "reading_pdf": "⏳ Читаю PDF...",
        "pdf_locked": "⭐ PDF — Standard или Premium!",
        "voice_locked": "💎 Голос — Premium!",
        "pptx_locked": "💎 PowerPoint — Premium!",
        "word_locked": "💎 Word — Premium!",
        "cv_locked": "⭐ Резюме — Standard или Premium!",
        "email_locked": "⭐ Email — Standard или Premium!",
        "ai_sound_locked": "⭐ AI Голос — Standard или Premium!",
        "document_locked": "⭐ Документ — Standard или Premium!",
        "writing_cv": "⏳ Пишу резюме...",
        "writing_email": "⏳ Пишу email...",
        "writing_post": "⏳ Пишу пост...",
        "writing_biznes": "⏳ Пишу бизнес-план...",
        "writing_document": "⏳ Создаю документ...",
        "translating": "⏳ Перевожу...",
        "writing_code": "⏳ Пишу код...",
        "pptx_example": "Пример: /pptx ИИ",
        "word_example": "Пример: /word бизнес-план",
        "cv_example": "Пример: /cv Python разработчик",
        "email_example": "Пример: /email после собеседования",
        "post_example": "Пример: /post кофейня",
        "biznes_example": "Пример: /biznes магазин",
        "creating_pptx": "⏳ Создаю '{topic}'...",
        "creating_word": "⏳ Создаю '{topic}'...",
        "ready": "✅ Готово!",
        "buy_standard": "⭐ Standard — 5 USDT/мес",
        "buy_premium": "💎 Premium — 10 USDT/мес",
        "buy_standard_stars": f"⭐ Standard — {STARS_STANDARD}⭐",
        "buy_premium_stars": f"💎 Premium — {STARS_PREMIUM}⭐",
        "buy_standard_coins": f"⭐ Standard — {COINS_STANDARD:,} монет",
        "buy_premium_coins": f"💎 Premium — {COINS_PREMIUM:,} монет",
        "buy_group": f"👥 Группа — {STARS_GROUP}⭐",
        "upgrade_standard": "⭐ Standard",
        "upgrade_premium": "💎 Premium",
        "upgrade_plan": "💰 Обновить тариф",
        "contact_admin": "💬 Админ",
        "you_said": "🎤 Вы сказали: ",
        "coins_info": "🪙 Ваши монеты: {coins}\n\n📈 Заработать:\n• 5 монет за сообщение\n• 25% с рефералов\n\n🛒 Потратить:\n• {standard:,} = ⭐ Standard 1 месяц\n• {premium:,} = 💎 Premium 1 месяц",
        "coins_standard_btn": f"Standard ({COINS_STANDARD:,} монет)",
        "coins_premium_btn": f"Premium ({COINS_PREMIUM:,} монет)",
        "not_enough_coins": "❌ Не хватает монет! Нужно ещё {need:,}.",
        "coins_success_standard": "✅ Standard активирован на 30 дней!",
        "coins_success_premium": "✅ Premium активирован на 30 дней!",
        "referral_title": "👥 Реферальная программа\n\n🔗 Ваша ссылка:\n{link}\n\n📊 Приглашено: {count}\n\n🎁 Награды:\n• 10 → ⭐ Standard 15 дней — {standard_status}\n• 30 → 💎 Premium 15 дней — {premium_status}\n\n🪙 25% монет с рефералов!",
        "claim_standard": "🎁 Standard 15 дней",
        "claim_premium": "🎁 Premium 15 дней",
        "claim_standard_success": "🎉 ⭐ Standard на 15 дней!",
        "claim_premium_success": "🎉 💎 Premium на 15 дней!",
        "claim_error": "❌ Уже получено или недостаточно!",
        "affiliate_info": "🤝 Аффилиат программа\n\n🔗 Ваша ссылка:\n{link}\n\n📊 Рефералов: {count}\n🪙 Заработано: {earned} монет\n\n💡 Как работает:\n• Поделитесь ссылкой\n• Получайте 25% монет с подписок",
        "stats_title": "📊 Ваша статистика",
        "gift_already": "🎁 Бонус уже получен! Завтра.",
        "gift_success": "🎁 Бонус!\nСтрик: {streak} дней 🔥\n+5 сообщений & +25 монет!",
        "gift_streak": "🎉 10 дней подряд!\n⭐ Standard на 10 дней!",
        "top_title": "🏆 Топ 10",
        "payment_info": "💳 Карта:\n`{card}`\n\n💵 {amount} USDT\n\n1. Скриншот → 2. Админу → 3. Активация ✅",
        "promo_enter": "Введите промокод: /promo КОД",
        "promo_invalid": "❌ Неверный промокод!",
        "promo_used": "❌ Уже использован!",
        "promo_success": "✅ Применён! {reward}",
        "group_info": "👥 Групповой режим — {stars}⭐\n\nАктивируйте для группы!",
        "warning": "⚠️ Предупреждение {count}/{max}: Будьте вежливы!",
        "banned_for_abuse": "🚫 Заблокированы за оскорбления.",
        "achievement_100k": "🏆 100,000 сообщений!\n⭐ Standard на 5 дней!",
        "imagine_pay": f"🎨 {STARS_IMAGINE}⭐ за одно изображение",
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

def is_bad_message(text):
    text_lower = text.lower()
    return any(word in text_lower for word in BAD_WORDS)

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
            coins INTEGER DEFAULT 0,
            warning_count INTEGER DEFAULT 0,
            affiliate_code TEXT,
            affiliate_earnings INTEGER DEFAULT 0,
            achievement_100k BOOLEAN DEFAULT FALSE,
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
        aff_code = generate_promo_code()
        c.execute("""
            INSERT INTO users (user_id, username, full_name, referral_code, referred_by, affiliate_code)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (user_id) DO UPDATE SET
                username = EXCLUDED.username,
                full_name = EXCLUDED.full_name
        """, (user_id, username, full_name, code, referred_by, aff_code))
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
                   streak_claimed, group_plan, total_messages, coins, warning_count,
                   affiliate_code, affiliate_earnings, achievement_100k
            FROM users WHERE user_id = %s
        """, (user_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return {"plan": "free", "expires_at": None, "is_blocked": False, "full_name": None,
                    "username": None, "referral_code": None, "referral_count": 0,
                    "claimed_standard": False, "claimed_premium": False, "language": "en",
                    "language_changed_at": None, "streak_count": 0, "last_gift_claim": None,
                    "streak_claimed": False, "group_plan": "none", "total_messages": 0,
                    "coins": 0, "warning_count": 0, "affiliate_code": None,
                    "affiliate_earnings": 0, "achievement_100k": False}
        plan, expires_at, is_blocked, full_name, username, referral_code, referral_count, \
        claimed_standard, claimed_premium, language, language_changed_at, streak_count, \
        last_gift_claim, streak_claimed, group_plan, total_messages, coins, warning_count, \
        affiliate_code, affiliate_earnings, achievement_100k = row
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
            "group_plan": group_plan or "none", "total_messages": total_messages or 0,
            "coins": coins or 0, "warning_count": warning_count or 0,
            "affiliate_code": affiliate_code, "affiliate_earnings": affiliate_earnings or 0,
            "achievement_100k": achievement_100k or False
        }
    except:
        return {"plan": "free", "expires_at": None, "is_blocked": False, "full_name": None,
                "username": None, "referral_code": None, "referral_count": 0,
                "claimed_standard": False, "claimed_premium": False, "language": "en",
                "language_changed_at": None, "streak_count": 0, "last_gift_claim": None,
                "streak_claimed": False, "group_plan": "none", "total_messages": 0,
                "coins": 0, "warning_count": 0, "affiliate_code": None,
                "affiliate_earnings": 0, "achievement_100k": False}

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

def add_coins(user_id, amount):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("UPDATE users SET coins = coins + %s WHERE user_id=%s", (amount, user_id))
        conn.commit()
        conn.close()
    except:
        pass

def add_warning(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("UPDATE users SET warning_count = warning_count + 1 WHERE user_id=%s", (user_id,))
        c.execute("SELECT warning_count FROM users WHERE user_id=%s", (user_id,))
        count = c.fetchone()[0]
        conn.commit()
        conn.close()
        return count
    except:
        return 0

def increment_messages(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            UPDATE users SET
            total_messages = total_messages + 1,
            coins = coins + %s
            WHERE user_id=%s
        """, (COINS_PER_MESSAGE, user_id))
        conn.commit()
        c.execute("SELECT total_messages, referred_by FROM users WHERE user_id=%s", (user_id,))
        row = c.fetchone()
        conn.close()
        if row:
            total, referred_by = row
            if referred_by:
                affiliate_bonus = int(COINS_PER_MESSAGE * AFFILIATE_PERCENT / 100)
                add_coins(referred_by, affiliate_bonus)
                try:
                    conn2 = get_conn()
                    c2 = conn2.cursor()
                    c2.execute("UPDATE users SET affiliate_earnings = affiliate_earnings + %s WHERE user_id=%s",
                              (affiliate_bonus, referred_by))
                    conn2.commit()
                    conn2.close()
                except:
                    pass
            return total
        return 0
    except:
        return 0

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
            "tts": 30, "imagine": 0, "translate": 30, "code": 30, "document": 30
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

async def generate_image_gemini(prompt):
    try:
        model = genai.GenerativeModel("gemini-2.0-flash-preview-image-generation")
        response = model.generate_content(
            f"Generate an image: {prompt}",
            generation_config=genai.GenerationConfig(response_modalities=["image", "text"])
        )
        for part in response.candidates[0].content.parts:
            if hasattr(part, 'inline_data') and part.inline_data:
                return part.inline_data.data
        return None
    except Exception as e:
        raise e

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    referred_by = None
    if context.args:
        try:
            ref_code = context.args[0]
            conn = get_conn()
            c = conn.cursor()
            c.execute("SELECT user_id FROM users WHERE referral_code = %s OR affiliate_code = %s", (ref_code, ref_code))
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
    await update.message.reply_text(get_text(lang, "welcome", plan_emoji=plan_emoji, plan=plan, coins=u["coins"]))

async def coins_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    keyboard = []
    if u["coins"] >= COINS_STANDARD:
        keyboard.append([InlineKeyboardButton(get_text(lang, "coins_standard_btn"), callback_data="coins_buy_standard")])
    if u["coins"] >= COINS_PREMIUM:
        keyboard.append([InlineKeyboardButton(get_text(lang, "coins_premium_btn"), callback_data="coins_buy_premium")])
    await update.message.reply_text(
        get_text(lang, "coins_info", coins=u["coins"], standard=COINS_STANDARD, premium=COINS_PREMIUM),
        reply_markup=InlineKeyboardMarkup(keyboard) if keyboard else None
    )

async def affiliate_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    bot = await context.bot.get_me()
    aff_code = u["affiliate_code"] or ""
    link = f"https://t.me/{bot.username}?start={aff_code}"
    await update.message.reply_text(
        get_text(lang, "affiliate_info", link=link, count=u["referral_count"], earned=u["affiliate_earnings"])
    )

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
            usage_chat = usage_chat + 5, coins = coins + 25
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
            SELECT username, full_name, total_messages, plan, coins
            FROM users WHERE is_blocked=FALSE
            ORDER BY total_messages DESC LIMIT 10
        """)
        rows = c.fetchall()
        conn.close()
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        text = f"{get_text(lang, 'top_title')}\n{'─' * 28}\n\n"
        medals = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"]
        for i, row in enumerate(rows):
            username, full_name, total_msgs, plan, coins = row
            name = f"@{username}" if username else (full_name or "Anonymous")
            emoji = plan_emoji.get(plan, "🆓")
            text += f"{medals[i]} {name} {emoji} — {total_msgs} msgs | 🪙{coins}\n"
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
            c.execute("SELECT reward_type, reward_value, max_uses, used_count, expires_at FROM promo_codes WHERE code=%s", (code,))
            promo = c.fetchone()
            if not promo:
                await update.message.reply_text(get_text(lang, "promo_invalid"))
                conn.close()
                return
            reward_type, reward_value, max_uses, used_count, expires_at = promo
            if (expires_at and datetime.now() > expires_at) or used_count >= max_uses:
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
            elif reward_type == "coins":
                add_coins(user.id, reward_value)
                reward_text = f"🪙 +{reward_value} coins!"
            else:
                reward_text = "✅"
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
        await update.message.reply_text("Usage: /createpromo [standard/premium/coins] [value] [max_uses]")
        return
    try:
        reward_type = context.args[0]
        reward_value = int(context.args[1])
        max_uses = int(context.args[2])
        code = generate_promo_code()
        conn = get_conn()
        c = conn.cursor()
        c.execute("INSERT INTO promo_codes (code, reward_type, reward_value, max_uses, expires_at) VALUES (%s, %s, %s, %s, %s)",
                  (code, reward_type, reward_value, max_uses, datetime.now() + timedelta(days=30)))
        conn.commit()
        conn.close()
        await update.message.reply_text(f"✅ Promo created!\nCode: `{code}`\nType: {reward_type}\nValue: {reward_value}\nUses: {max_uses}", parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def group_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    keyboard = [[InlineKeyboardButton(get_text(lang, "buy_group"), callback_data="stars_group")]]
    await update.message.reply_text(get_text(lang, "group_info", stars=STARS_GROUP), reply_markup=InlineKeyboardMarkup(keyboard))

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
        c.execute("SELECT usage_chat, usage_search, usage_image, usage_post, usage_biznes, usage_pdf, usage_cv, usage_email, usage_tts FROM users WHERE user_id = %s", (user.id,))
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
            f"👤 {plan_emoji} {plan.upper()}\n📅 {expires}\n"
            f"🪙 Coins: {u['coins']}\n👥 Referrals: {u['referral_count']}\n"
            f"🔥 Streak: {u['streak_count']} days\n💬 Total: {u['total_messages']}\n\n"
            f"📊 Today:\n💬{chat} 🌐{search} 🖼️{image} 📱{post}\n💼{biznes} 📄{pdf} 👤{cv} 📧{email} 🔊{tts}"
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
        get_text(lang, "referral_title", link=link, count=count, standard_status=standard_status, premium_status=premium_status),
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
        reply = await ai_generate(f"Translate accurately: {text}")
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
        reply = await ai_generate(f"Write clean, well-commented code for: {text}\nExplain how it works.")
        await update.message.reply_text(f"💻 {reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def document_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]: return
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
        reply = await ai_generate(f"Create a professional document: {text}\nFormat properly with all sections. Same language as request.")
        await update.message.reply_text(f"📋 {reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def imagine_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    if u["is_blocked"]: return
    limits = get_limits(u["plan"])
    prompt = " ".join(context.args)
    if not prompt:
        await update.message.reply_text(get_text(lang, "imagine_example"))
        return
    if limits["imagine"] == 0:
        keyboard = [
            [InlineKeyboardButton(get_text(lang, "upgrade_premium"), callback_data="buy_premium")],
            [InlineKeyboardButton(get_text(lang, "imagine_pay"), callback_data=f"imagine_pay_{prompt[:50]}")],
        ]
        await update.message.reply_text(get_text(lang, "imagine_locked", stars=STARS_IMAGINE), reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "imagine", limits["imagine"]):
        await update.message.reply_text(get_text(lang, "limit_reached"))
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
        await update.message.reply_text(f"❌ {str(e)}")

async def updateplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    lang = u["language"]
    plan = u["plan"]
    expires = u["expires_at"].strftime("%d.%m.%Y") if u.get("expires_at") else "—"
    plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}.get(plan, "🆓")
    keyboard = [
        [InlineKeyboardButton(get_text(lang, "buy_standard_stars"), callback_data="stars_standard"),
         InlineKeyboardButton(get_text(lang, "buy_premium_stars"), callback_data="stars_premium")],
        [InlineKeyboardButton(get_text(lang, "buy_standard_coins"), callback_data="coins_buy_standard"),
         InlineKeyboardButton(get_text(lang, "buy_premium_coins"), callback_data="coins_buy_premium")],
        [InlineKeyboardButton(get_text(lang, "buy_standard"), callback_data="buy_standard"),
         InlineKeyboardButton(get_text(lang, "buy_premium"), callback_data="buy_premium")],
        [InlineKeyboardButton(get_text(lang, "buy_group"), callback_data="stars_group")],
    ]
    await update.message.reply_text(
        f"💰 Plans\n{'─' * 28}\n\n"
        f"🆓 FREE: Chat 20/day, Search 20/day\n\n"
        f"⭐ STANDARD — 5 USDT / {STARS_STANDARD}⭐ / {COINS_STANDARD:,}🪙\n"
        f"Everything 30/day + PDF, CV, Email, Voice, Image, Document\n\n"
        f"💎 PREMIUM — 10 USDT / {STARS_PREMIUM}⭐ / {COINS_PREMIUM:,}🪙\n"
        f"Everything Unlimited + AI Image, PowerPoint, Word\n\n"
        f"👥 GROUP — {STARS_GROUP}⭐ ($99.9)\n\n"
        f"{'─' * 28}\n{plan_emoji} {plan.upper()} | 📅 {expires} | 🪙 {u['coins']}",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def precheckout_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.pre_checkout_query.answer(ok=True)

async def successful_payment_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    payload = update.message.successful_payment.invoice_payload
    if payload == "standard_stars":
        set_plan(user.id, "standard", 30)
        await update.message.reply_text("✅ ⭐ Standard activated for 30 days!")
    elif payload == "premium_stars":
        set_plan(user.id, "premium", 30)
        await update.message.reply_text("✅ 💎 Premium activated for 30 days!")
    elif payload == "group_stars":
        conn = get_conn()
        c = conn.cursor()
        c.execute("UPDATE users SET group_plan='active' WHERE user_id=%s", (user.id,))
        conn.commit()
        conn.close()
        await update.message.reply_text("✅ 👥 Group mode activated!")
    elif payload.startswith("imagine_"):
        prompt = payload[8:]
        await update.message.reply_text("⏳ Generating image...")
        try:
            img_bytes = await generate_image_gemini(prompt)
            if img_bytes:
                buf = io.BytesIO(img_bytes)
                buf.name = "image.png"
                await update.message.reply_photo(photo=buf, caption=f"🎨 {prompt}")
            else:
                await update.message.reply_text("❌ Could not generate. Try again.")
        except Exception as e:
            await update.message.reply_text(f"❌ {str(e)}")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    u = get_user(user_id)
    lang = u["language"]

    if data.startswith("imagine_pay_"):
        prompt = data[12:]
        await context.bot.send_invoice(
            chat_id=user_id,
            title="🎨 AI Image",
            description=f"Generate: {prompt}",
            payload=f"imagine_{prompt}",
            currency="XTR",
            prices=[LabeledPrice("AI Image", STARS_IMAGINE)],
        )

    elif data == "stars_standard":
        await context.bot.send_invoice(chat_id=user_id, title="⭐ Standard Plan",
            description="Standard for 30 days", payload="standard_stars",
            currency="XTR", prices=[LabeledPrice("Standard 30d", STARS_STANDARD)])

    elif data == "stars_premium":
        await context.bot.send_invoice(chat_id=user_id, title="💎 Premium Plan",
            description="Premium for 30 days", payload="premium_stars",
            currency="XTR", prices=[LabeledPrice("Premium 30d", STARS_PREMIUM)])

    elif data == "stars_group":
        await context.bot.send_invoice(chat_id=user_id, title="👥 Group Mode",
            description="Activate for your group", payload="group_stars",
            currency="XTR", prices=[LabeledPrice("Group Mode", STARS_GROUP)])

    elif data == "coins_buy_standard":
        if u["coins"] >= COINS_STANDARD:
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE users SET coins = coins - %s WHERE user_id=%s", (COINS_STANDARD, user_id))
            conn.commit()
            conn.close()
            set_plan(user_id, "standard", 30)
            await query.message.edit_text(get_text(lang, "coins_success_standard"))
        else:
            need = COINS_STANDARD - u["coins"]
            await query.answer(get_text(lang, "not_enough_coins", need=need), show_alert=True)

    elif data == "coins_buy_premium":
        if u["coins"] >= COINS_PREMIUM:
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE users SET coins = coins - %s WHERE user_id=%s", (COINS_PREMIUM, user_id))
            conn.commit()
            conn.close()
            set_plan(user_id, "premium", 30)
            await query.message.edit_text(get_text(lang, "coins_success_premium"))
        else:
            need = COINS_PREMIUM - u["coins"]
            await query.answer(get_text(lang, "not_enough_coins", need=need), show_alert=True)

    elif data.startswith("setlang_"):
        new_lang = data.split("_")[1]
        last_changed = u["language_changed_at"]
        if last_changed and datetime.now() - last_changed < timedelta(hours=24):
            await query.answer(get_text(lang, "language_cooldown"), show_alert=True)
            return
        set_language(user_id, new_lang)
        lang_name = LANGUAGES.get(new_lang, ("", "Unknown"))[1]
        await query.message.edit_text(f"✅ Language: {lang_name}")

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
            f"⭐ STANDARD — 5 USDT\n" + get_text(lang, "payment_info", card=CARD_NUMBER, amount=5),
            parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "buy_premium":
        keyboard = [[InlineKeyboardButton(get_text(lang, "contact_admin"), url=f"https://t.me/{ADMIN_USERNAME}")]]
        await query.message.reply_text(
            f"💎 PREMIUM — 10 USDT\n" + get_text(lang, "payment_info", card=CARD_NUMBER, amount=10),
            parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("ap_setplan_"):
        parts = data.split("_")
        target_id = int(parts[2])
        plan = parts[3]
        days = 30 if plan != "free" else None
        set_plan(target_id, plan, days)
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        await query.message.edit_text(f"✅ User {target_id} → {plan_emoji[plan]} {plan.upper()}")
        try:
            await context.bot.send_message(chat_id=target_id, text=f"{plan_emoji[plan]} Plan updated!")
        except: pass

    elif data.startswith("ap_block_"):
        target_id = int(data.split("_")[2])
        set_blocked(target_id, True)
        await query.message.edit_text(f"🚫 {target_id} blocked!")

    elif data.startswith("ap_unblock_"):
        target_id = int(data.split("_")[2])
        set_blocked(target_id, False)
        await query.message.edit_text(f"✅ {target_id} unblocked!")

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
        c.execute("SELECT SUM(coins) FROM users")
        total_coins = c.fetchone()[0] or 0
        conn.close()
        revenue = standard * 5 + premium * 10
    except:
        total = standard = premium = blocked = revenue = total_coins = 0
    await update.message.reply_text(
        f"🔧 Admin Panel\n{'═'*28}\n\n"
        f"👥 Total: {total}\n🆓 Free: {total-standard-premium}\n"
        f"⭐ Standard: {standard}\n💎 Premium: {premium}\n"
        f"🚫 Blocked: {blocked}\n💰 Revenue: ~${revenue}\n"
        f"🪙 Total coins: {total_coins:,}\n\n"
        f"/users /find [id/@user] /broadcast [msg]\n"
        f"/createpromo [type] [days] [uses]\n"
        f"Types: standard, premium, coins"
    )

async def users_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT user_id, username, plan, is_blocked FROM users ORDER BY CASE plan WHEN 'premium' THEN 1 WHEN 'standard' THEN 2 ELSE 3 END, joined_at DESC")
        rows = c.fetchall()
        conn.close()
        if not rows:
            await update.message.reply_text("No users.")
            return
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        text = f"👥 Users: {len(rows)}\n{'═'*30}\n\n"
        for row in rows:
            uid, username, plan, is_blocked = row
            uname = f"@{username}" if username else "no_username"
            blocked = " 🚫" if is_blocked else ""
            text += f"{plan_emoji.get(plan,'🆓')} {uname} {uid}{blocked}\n"
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
        await update.message.reply_text("Usage: /find [id or @username]")
        return
    try:
        arg = context.args[0]
        conn = get_conn()
        c = conn.cursor()
        if arg.startswith("@"):
            c.execute("SELECT plan, expires_at, is_blocked, full_name, username, user_id FROM users WHERE username=%s", (arg[1:],))
        else:
            c.execute("SELECT plan, expires_at, is_blocked, full_name, username, user_id FROM users WHERE user_id=%s", (int(arg),))
        row = c.fetchone()
        conn.close()
        if not row:
            await update.message.reply_text(f"❌ Not found: {arg}")
            return
        plan, expires_at, is_blocked, full_name, username, target_id = row
        expires = expires_at.strftime("%d.%m.%Y") if expires_at else "—"
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        keyboard = [
            [InlineKeyboardButton("🆓", callback_data=f"ap_setplan_{target_id}_free"),
             InlineKeyboardButton("⭐", callback_data=f"ap_setplan_{target_id}_standard"),
             InlineKeyboardButton("💎", callback_data=f"ap_setplan_{target_id}_premium")],
            [InlineKeyboardButton("✅ Unblock" if is_blocked else "🚫 Block",
             callback_data=f"ap_unblock_{target_id}" if is_blocked else f"ap_block_{target_id}")]
        ]
        await update.message.reply_text(
            f"👤 {full_name or '—'} | @{username or '—'}\n"
            f"🆔 {target_id}\n{plan_emoji.get(plan,'🆓')} {plan.upper()} | 📅 {expires}\n"
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
        c.execute("SELECT user_id FROM users WHERE is_blocked=FALSE")
        rows = c.fetchall()
        conn.close()
        sent = 0
        for row in rows:
            try:
                await context.bot.send_message(chat_id=row[0], text=f"📢 {message}")
                sent += 1
            except: pass
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
        reply = await ai_generate(f"Write professional CV for: {info}\nSections: Summary, Experience, Skills, Education. ATS-friendly.")
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
        reply = await ai_generate(f"Write professional email about: {topic}\nInclude: Subject, greeting, body, closing.")
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
        reply = await ai_generate(f"Write engaging social media post about: {topic}\nInclude: Hook, content, CTA, hashtags.")
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
        reply = await ai_generate(f"Write detailed business plan for: {idea}\nInclude: Executive Summary, Market Analysis, Products, Marketing, Financial Plan.")
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

    user_text = update.message.text

    if is_bad_message(user_text):
        warning_count = add_warning(user_id)
        if warning_count >= MAX_WARNINGS:
            set_blocked(user_id, True)
            await update.message.reply_text(get_text(lang, "banned_for_abuse"))
        else:
            await update.message.reply_text(get_text(lang, "warning", count=warning_count, max=MAX_WARNINGS))
        return

    limits = get_limits(u["plan"])
    if not check_limit(user_id, "chat", limits["chat"]):
        keyboard = [[InlineKeyboardButton(get_text(lang, "upgrade_plan"), callback_data="buy_standard")]]
        await update.message.reply_text(get_text(lang, "limit_reached"), reply_markup=InlineKeyboardMarkup(keyboard))
        return

    total = increment_messages(user_id)

    if total == 100000 and not u["achievement_100k"]:
        set_plan(user_id, "standard", 5)
        try:
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE users SET achievement_100k=TRUE WHERE user_id=%s", (user_id,))
            conn.commit()
            conn.close()
        except: pass
        await update.message.reply_text(get_text(lang, "achievement_100k"))

    memory = get_memory(user_id)
    memory_context = ""
    if memory and (memory.get("name") or memory.get("facts")):
        memory_context = f"User info — name: {memory['name']}, facts: {memory['facts']}. "

    system_prompt = (
        "You are a helpful AI assistant. "
        "Always reply in the SAME language as the user's current message. "
        "Uzbek→Uzbek, English→English, Russian→Russian. "
        "If asked to switch language, do it immediately. "
        "Be accurate, helpful and concise. "
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
            message_content = f"Question: {user_text}\n\nWeb results:\n{search_content}\n\nAnswer in same language."
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
            user_histories[user.id] = [{"role": "system", "content": "You are a helpful AI. Reply in user's language."}]
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
        caption = update.message.caption or "What is in this image? Describe in detail."
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
        caption = update.message.caption or "Summarize and explain key points."
        if user.id not in user_histories:
            user_histories[user.id] = [{"role": "system", "content": "You are a helpful AI. Reply in user's language."}]
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
        BotCommand("start", "Start"),
        BotCommand("updateplan", "Plans & pricing"),
        BotCommand("pptx", "PowerPoint (Premium)"),
        BotCommand("word", "Word (Premium)"),
        BotCommand("cv", "CV (Standard+)"),
        BotCommand("email", "Email (Standard+)"),
        BotCommand("post", "Marketing post"),
        BotCommand("biznes", "Business plan"),
        BotCommand("ai_sound", "AI Voice (Standard+)"),
        BotCommand("imagine", "AI Image (Premium)"),
        BotCommand("translate", "Translator"),
        BotCommand("code", "Code writer"),
        BotCommand("document", "Document (Standard+)"),
        BotCommand("coins", "My coins"),
        BotCommand("referral", "Invite friends"),
        BotCommand("affiliate", "Affiliate program"),
        BotCommand("gift", "Daily bonus"),
        BotCommand("top", "Top users"),
        BotCommand("stats", "Statistics"),
        BotCommand("promo", "Promo code"),
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
    app.add_handler(CommandHandler("coins", coins_command))
    app.add_handler(CommandHandler("referral", referral_command))
    app.add_handler(CommandHandler("affiliate", affiliate_command))
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
