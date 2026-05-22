import os
import io
import requests
import fitz
import json
import psycopg2
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

ADMIN_ID = 8230883785
CARD_NUMBER = "48547002151326"
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
    "nl": "nl", "pl": "pl", "uk": "uk", "kk": "ru"
}

def detect_lang(text):
    try:
        lang = detect(text)
        return GTTS_LANG_MAP.get(lang, "en")
    except:
        return "en"

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

def ensure_user(user_id, username=None, full_name=None):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO users (user_id, username, full_name)
            VALUES (%s, %s, %s)
            ON CONFLICT (user_id) DO UPDATE SET
                username = EXCLUDED.username,
                full_name = EXCLUDED.full_name
        """, (user_id, username, full_name))
        conn.commit()
        conn.close()
    except:
        pass

def get_user(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            SELECT plan, expires_at, is_blocked, full_name, username 
            FROM users WHERE user_id = %s
        """, (user_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return {"plan": "free", "expires_at": None, "is_blocked": False, "full_name": None, "username": None}
        plan, expires_at, is_blocked, full_name, username = row
        if expires_at and datetime.now() > expires_at and plan != 'free':
            set_plan(user_id, "free", None)
            plan = "free"
            expires_at = None
        return {
            "plan": plan,
            "expires_at": expires_at,
            "is_blocked": is_blocked,
            "full_name": full_name,
            "username": username
        }
    except:
        return {"plan": "free", "expires_at": None, "is_blocked": False, "full_name": None, "username": None}

def set_plan(user_id, plan, days=30):
    try:
        conn = get_conn()
        c = conn.cursor()
        expires_at = datetime.now() + timedelta(days=days) if days else None
        c.execute("""
            UPDATE users SET plan=%s, expires_at=%s WHERE user_id=%s
        """, (plan, expires_at, user_id))
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
                last_reset=%s
                WHERE user_id=%s
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
            "pdf": 30, "cv": 30, "email": 30, "voice": 0, "pptx": 0, "word": 0,
            "tts": 30
        }
    else:
        return {
            "chat": 20, "search": 20, "image": 20, "post": 20, "biznes": 20,
            "pdf": 0, "cv": 0, "email": 0, "voice": 0, "pptx": 0, "word": 0,
            "tts": 0
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
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    plan = u["plan"].upper()
    plan_emoji = {"FREE": "🆓", "STANDARD": "⭐", "PREMIUM": "💎"}.get(plan, "🆓")
    await update.message.reply_text(
        f"Hello! I am Chatbot 🤖\n"
        f"Your plan: {plan_emoji} {plan}\n\n"
        f"💬 Chat with me\n"
        f"🌐 Current news, prices, weather\n"
        f"📄 Send PDF to analyze\n"
        f"🖼️ Send image\n"
        f"🎤 Send voice message\n\n"
        f"📊 /pptx — PowerPoint\n"
        f"📝 /word — Word document\n"
        f"👤 /cv — Write CV\n"
        f"📧 /email — Write email\n"
        f"📱 /post — Marketing post\n"
        f"🔊 /ai_sound — AI Voice (Standard+)\n\n"
        f"💰 /updateplan — Update plan\n"
        f"/help — Help\n"
        f"/reset — Clear history"
    )

async def updateplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    plan = u["plan"]
    expires = u["expires_at"].strftime("%d.%m.%Y") if u.get("expires_at") else "—"
    plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}.get(plan, "🆓")
    keyboard = [
        [InlineKeyboardButton("⭐ Buy Standard — 5 USDT/month", callback_data="buy_standard")],
        [InlineKeyboardButton("💎 Buy Premium — 10 USDT/month", callback_data="buy_premium")],
    ]
    await update.message.reply_text(
        f"💰 Plans & Pricing\n"
        f"{'─' * 28}\n\n"
        f"🆓 FREE — Free\n"
        f"• Chat: 20/day\n"
        f"• Internet search: 20/day\n"
        f"• Image analysis: 20/day\n"
        f"• Marketing post: 20/day\n"
        f"• Business plan: 20/day\n"
        f"• Memory: Unlimited\n\n"
        f"⭐ STANDARD — 5 USDT/month\n"
        f"• Everything in Free: 30/day\n"
        f"• PDF analysis: 30/day\n"
        f"• CV writing: 30/day\n"
        f"• Email writing: 30/day\n"
        f"• AI Voice: 30/day\n"
        f"• Memory: Unlimited\n\n"
        f"💎 PREMIUM — 10 USDT/month\n"
        f"• Everything: Unlimited\n"
        f"• Voice messages ✓\n"
        f"• PowerPoint creation ✓\n"
        f"• Word documents ✓\n"
        f"• Priority support ✓\n\n"
        f"{'─' * 28}\n"
        f"👤 Your plan: {plan_emoji} {plan.upper()}\n"
        f"📅 Expires: {expires}",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id

    if data == "buy_standard":
        keyboard = [[InlineKeyboardButton("💬 Contact Admin", url=f"https://t.me/{ADMIN_USERNAME}")]]
        await query.message.reply_text(
            f"⭐ STANDARD — 5 USDT/month\n"
            f"{'─' * 30}\n\n"
            f"💳 Pay to card:\n"
            f"`{CARD_NUMBER}`\n\n"
            f"💵 Amount: 5 USDT equivalent\n\n"
            f"📋 After payment:\n"
            f"1. Take a screenshot\n"
            f"2. Send it to admin\n"
            f"3. Your plan will be activated within 1 hour ✅",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data == "buy_premium":
        keyboard = [[InlineKeyboardButton("💬 Contact Admin", url=f"https://t.me/{ADMIN_USERNAME}")]]
        await query.message.reply_text(
            f"💎 PREMIUM — 10 USDT/month\n"
            f"{'─' * 30}\n\n"
            f"💳 Pay to card:\n"
            f"`{CARD_NUMBER}`\n\n"
            f"💵 Amount: 10 USDT equivalent\n\n"
            f"📋 After payment:\n"
            f"1. Take a screenshot\n"
            f"2. Send it to admin\n"
            f"3. Your plan will be activated within 1 hour ✅",
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
        await query.message.edit_text(
            f"✅ Done!\n"
            f"User {target_id} → {plan_emoji[plan]} {plan.upper()}"
        )
        try:
            await context.bot.send_message(
                chat_id=target_id,
                text=f"{plan_emoji[plan]} Your plan has been updated to {plan.upper()}!\n\nUse /updateplan to see details."
            )
        except:
            pass

    elif data.startswith("ap_block_"):
        target_id = int(data.split("_")[2])
        set_blocked(target_id, True)
        await query.message.edit_text(f"🚫 User {target_id} has been blocked!")

    elif data.startswith("ap_unblock_"):
        target_id = int(data.split("_")[2])
        set_blocked(target_id, False)
        await query.message.edit_text(f"✅ User {target_id} has been unblocked!")

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
        f"🔧 Admin Panel\n"
        f"{'═' * 28}\n\n"
        f"📊 Statistics:\n"
        f"👥 Total users: {total}\n"
        f"🆓 Free: {total - standard - premium}\n"
        f"⭐ Standard: {standard}\n"
        f"💎 Premium: {premium}\n"
        f"🚫 Blocked: {blocked}\n"
        f"💰 Monthly revenue: ~${revenue}\n\n"
        f"{'─' * 28}\n"
        f"Commands:\n"
        f"/users — All users list\n"
        f"/find [user\\_id] — Manage user\n"
        f"/broadcast [text] — Message all"
    )

async def users_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            SELECT user_id, username, plan, is_blocked
            FROM users
            ORDER BY
                CASE plan
                    WHEN 'premium' THEN 1
                    WHEN 'standard' THEN 2
                    ELSE 3
                END,
                joined_at DESC
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
            emoji = plan_emoji.get(plan, "🆓")
            text += f"{emoji} Username: {uname}    ID: {uid}{blocked}\n"

        if len(text) > 4000:
            parts = [text[i:i+4000] for i in range(0, len(text), 4000)]
            for part in parts:
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
        c.execute("""
            SELECT plan, expires_at, is_blocked, full_name, username 
            FROM users WHERE user_id = %s
        """, (target_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            await update.message.reply_text(f"❌ User {target_id} not found!")
            return

        plan, expires_at, is_blocked, full_name, username = row
        expires = expires_at.strftime("%d.%m.%Y") if expires_at else "—"
        status = "🚫 Blocked" if is_blocked else "✅ Active"
        uname = f"@{username}" if username else "—"
        name = full_name or "—"
        plan_emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}

        keyboard = [
            [
                InlineKeyboardButton("🆓 Free", callback_data=f"ap_setplan_{target_id}_free"),
                InlineKeyboardButton("⭐ Standard", callback_data=f"ap_setplan_{target_id}_standard"),
                InlineKeyboardButton("💎 Premium", callback_data=f"ap_setplan_{target_id}_premium"),
            ],
            [
                InlineKeyboardButton(
                    "✅ Unblock" if is_blocked else "🚫 Block",
                    callback_data=f"ap_unblock_{target_id}" if is_blocked else f"ap_block_{target_id}"
                )
            ]
        ]

        await update.message.reply_text(
            f"👤 User Info\n"
            f"{'─' * 25}\n"
            f"🆔 ID: {target_id}\n"
            f"👤 Name: {name}\n"
            f"📱 Username: {uname}\n"
            f"📋 Plan: {plan_emoji.get(plan, '🆓')} {plan.upper()}\n"
            f"📅 Expires: {expires}\n"
            f"🔰 Status: {status}\n"
            f"{'─' * 25}\n"
            f"Select action:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    except ValueError:
        await update.message.reply_text("❌ Invalid user ID!")
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
    await update.message.reply_text(
        "📌 Commands:\n\n"
        "/pptx — PowerPoint 💎\n"
        "/word — Word document 💎\n"
        "/cv — CV/Resume ⭐\n"
        "/email — Email ⭐\n"
        "/post — Marketing post\n"
        "/biznes — Business plan\n"
        "/ai_sound — AI Voice 🔊 ⭐\n"
        "/updateplan — Plans & pricing\n"
        "/reset — Clear chat history\n"
        "/help — Help\n\n"
        "⭐ = Standard or Premium\n"
        "💎 = Premium only"
    )

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_histories[update.effective_user.id] = []
    await update.message.reply_text("Chat history cleared ✅")

async def pptx_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    if u["is_blocked"]:
        return
    if get_limits(u["plan"])["pptx"] == 0:
        keyboard = [[InlineKeyboardButton("💎 Upgrade to Premium", callback_data="buy_premium")]]
        await update.message.reply_text("💎 PowerPoint requires Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /pptx artificial intelligence")
        return
    await update.message.reply_text(f"⏳ Creating presentation on '{topic}'...")
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
            await update.message.reply_document(document=f, filename=f"{topic}.pptx", caption="✅ Ready!")
        os.remove(path)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def word_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    if u["is_blocked"]:
        return
    if get_limits(u["plan"])["word"] == 0:
        keyboard = [[InlineKeyboardButton("💎 Upgrade to Premium", callback_data="buy_premium")]]
        await update.message.reply_text("💎 Word requires Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /word business plan")
        return
    await update.message.reply_text(f"⏳ Creating document on '{topic}'...")
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
            await update.message.reply_document(document=f, filename=f"{topic}.docx", caption="✅ Ready!")
        os.remove(path)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def cv_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["cv"] == 0:
        keyboard = [[InlineKeyboardButton("⭐ Upgrade to Standard", callback_data="buy_standard")]]
        await update.message.reply_text("⭐ CV requires Standard or Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "cv", limits["cv"]):
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.")
        return
    info = " ".join(context.args)
    if not info:
        await update.message.reply_text("Example: /cv Python developer, 3 years experience")
        return
    await update.message.reply_text("⏳ Writing your CV...")
    try:
        reply = await ai_generate(f"Write a professional CV/Resume for: {info}\nFormat: Summary, Experience, Skills, Education. ATS-friendly.")
        await update.message.reply_text(f"✅ Your CV:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def email_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["email"] == 0:
        keyboard = [[InlineKeyboardButton("⭐ Upgrade to Standard", callback_data="buy_standard")]]
        await update.message.reply_text("⭐ Email requires Standard or Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user.id, "email", limits["email"]):
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.")
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /email follow up after interview")
        return
    await update.message.reply_text("⏳ Writing email...")
    try:
        reply = await ai_generate(f"Write a professional email about: {topic}\nInclude: Subject, greeting, body, closing.")
        await update.message.reply_text(f"✅ Your email:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def post_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if not check_limit(user.id, "post", limits["post"]):
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.")
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /post new coffee shop opening")
        return
    await update.message.reply_text("⏳ Writing marketing post...")
    try:
        reply = await ai_generate(f"Write an engaging social media post about: {topic}\nInclude: Hook, content, call to action, hashtags.")
        await update.message.reply_text(f"✅ Your post:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def biznes_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.username, user.full_name)
    u = get_user(user.id)
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if not check_limit(user.id, "biznes", limits["biznes"]):
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.")
        return
    idea = " ".join(context.args)
    if not idea:
        await update.message.reply_text("Example: /biznes online clothing store")
        return
    await update.message.reply_text("⏳ Writing business plan...")
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
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits.get("tts") == 0:
        keyboard = [[InlineKeyboardButton("⭐ Upgrade to Standard", callback_data="buy_standard")]]
        await update.message.reply_text("⭐ AI Sound requires Standard or Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user_id, "tts", limits["tts"]):
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.")
        return
    text = " ".join(context.args)
    if not text:
        await update.message.reply_text("Example: /ai_sound Hello, how are you?")
        return
    await update.message.reply_text("⏳ Generating voice...")
    try:
        lang = detect_lang(text)
        tts = gTTS(text=text, lang=lang)
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

    if u["is_blocked"]:
        await update.message.reply_text("🚫 You are blocked. Contact admin.")
        return

    limits = get_limits(u["plan"])
    if not check_limit(user_id, "chat", limits["chat"]):
        keyboard = [[InlineKeyboardButton("💰 Upgrade Plan", callback_data="buy_standard")]]
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return

    user_text = update.message.text
    memory = get_memory(user_id)
    memory_context = ""
    if memory and (memory.get("name") or memory.get("facts")):
        memory_context = f"User info — name: {memory['name']}, facts: {memory['facts']}. "

    if user_id not in user_histories:
        user_histories[user_id] = [{"role": "system", "content": (
            "You are a professional AI assistant. "
            "VERY IMPORTANT: Always reply in the SAME language the user uses in their message. "
            "If user writes in Uzbek, reply in Uzbek. "
            "If user writes in English, reply in English. "
            "If user writes in Russian, reply in Russian. "
            "Never switch languages on your own. "
            "Keep answers short, clear and natural. "
            + memory_context
        )}]
    else:
        user_histories[user_id][0]["content"] = (
            "You are a professional AI assistant. "
            "VERY IMPORTANT: Always reply in the SAME language the user uses in their message. "
            "If user writes in Uzbek, reply in Uzbek. "
            "If user writes in English, reply in English. "
            "If user writes in Russian, reply in Russian. "
            "Never switch languages on your own. "
            "Keep answers short, clear and natural. "
            + memory_context
        )

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")

    try:
        if needs_search(user_text) and limits["search"] != 0:
            search_results = tavily.search(query=user_text, max_results=3)
            search_content = "\n\n".join([f"Source: {r['url']}\n{r['content']}" for r in search_results.get("results", [])])
            message_content = f"User question: {user_text}\n\nWeb results:\n{search_content}\n\nGive a clear answer in the same language as the user's question."
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
    if u["is_blocked"]:
        return
    if get_limits(u["plan"])["voice"] == 0:
        keyboard = [[InlineKeyboardButton("💎 Upgrade to Premium", callback_data="buy_premium")]]
        await update.message.reply_text("💎 Voice requires Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
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
        await update.message.reply_text(f"🎤 You said: {user_text}")
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
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if not check_limit(user_id, "image", limits["image"]):
        keyboard = [[InlineKeyboardButton("💰 Upgrade Plan", callback_data="buy_standard")]]
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
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
    if u["is_blocked"]:
        return
    limits = get_limits(u["plan"])
    if limits["pdf"] == 0:
        keyboard = [[InlineKeyboardButton("⭐ Upgrade to Standard", callback_data="buy_standard")]]
        await update.message.reply_text("⭐ PDF requires Standard or Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user_id, "pdf", limits["pdf"]):
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.")
        return
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        doc = update.message.document
        if not doc.file_name.endswith(".pdf"):
            await update.message.reply_text("Please send a PDF file! 📄")
            return
        await update.message.reply_text("⏳ Reading PDF...")
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
        user_histories[user_id].append({"role": "user", "content": f"PDF:\n\n{text}\n\nRequest: {caption}"}),
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
