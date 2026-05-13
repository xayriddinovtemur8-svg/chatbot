import os
import requests
import fitz
import json
import psycopg2
from datetime import datetime, timedelta
from groq import Groq
from tavily import TavilyClient
from dotenv import load_dotenv
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

ADMIN_ID = 8230883785
CARD_NUMBER = "4916990345412073"
ADMIN_USERNAME = "temur_uzb7779"

client = Groq(api_key=GROQ_API_KEY)
tavily = TavilyClient(api_key=TAVILY_API_KEY)

user_histories = {}

SEARCH_KEYWORDS = [
    "today", "now", "current", "latest", "news", "price", "rate", "weather",
    "bugun", "hozir", "narx", "kurs", "yangilik", "ob-havo", "oxirgi",
    "сегодня", "сейчас", "курс", "цена", "новости"
]

# ===== DATABASE =====
def get_conn():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS user_memory (
            user_id BIGINT PRIMARY KEY,
            name TEXT,
            facts TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS subscriptions (
            user_id BIGINT PRIMARY KEY,
            plan TEXT DEFAULT 'free',
            expires_at TIMESTAMP,
            usage_chat INTEGER DEFAULT 0,
            usage_search INTEGER DEFAULT 0,
            usage_image INTEGER DEFAULT 0,
            usage_post INTEGER DEFAULT 0,
            usage_biznes INTEGER DEFAULT 0,
            usage_pdf INTEGER DEFAULT 0,
            usage_cv INTEGER DEFAULT 0,
            usage_email INTEGER DEFAULT 0,
            last_reset DATE DEFAULT CURRENT_DATE
        )
    """)
    conn.commit()
    conn.close()

def get_subscription(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT plan, expires_at, last_reset FROM subscriptions WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return {"plan": "free", "expires_at": None}
        plan, expires_at, last_reset = row
        if expires_at and datetime.now() > expires_at:
            set_plan(user_id, "free", None)
            return {"plan": "free", "expires_at": None}
        return {"plan": plan, "expires_at": expires_at}
    except:
        return {"plan": "free", "expires_at": None}

def ensure_user(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO subscriptions (user_id) VALUES (%s)
            ON CONFLICT (user_id) DO NOTHING
        """, (user_id,))
        conn.commit()
        conn.close()
    except:
        pass

def set_plan(user_id, plan, days=30):
    try:
        conn = get_conn()
        c = conn.cursor()
        expires_at = datetime.now() + timedelta(days=days) if days else None
        c.execute("""
            INSERT INTO subscriptions (user_id, plan, expires_at)
            VALUES (%s, %s, %s)
            ON CONFLICT (user_id) DO UPDATE SET plan=%s, expires_at=%s
        """, (user_id, plan, expires_at, plan, expires_at))
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
        c.execute("SELECT last_reset FROM subscriptions WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        if row and row[0] < today:
            c.execute("""
                UPDATE subscriptions SET
                usage_chat=0, usage_search=0, usage_image=0,
                usage_post=0, usage_biznes=0, usage_pdf=0,
                usage_cv=0, usage_email=0, last_reset=%s
                WHERE user_id=%s
            """, (today, user_id))
            conn.commit()
        c.execute(f"SELECT usage_{feature} FROM subscriptions WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        usage = row[0] if row else 0
        if usage >= limit:
            conn.close()
            return False
        c.execute(f"UPDATE subscriptions SET usage_{feature} = usage_{feature} + 1 WHERE user_id = %s", (user_id,))
        conn.commit()
        conn.close()
        return True
    except:
        return True

def get_limits(plan):
    if plan == "premium":
        return {k: -1 for k in ["chat","search","image","post","biznes","pdf","cv","email","voice","pptx","word"]}
    elif plan == "standard":
        return {
            "chat": 30, "search": 30, "image": 30, "post": 30, "biznes": 30,
            "pdf": 30, "cv": 30, "email": 30, "voice": 0, "pptx": 0, "word": 0
        }
    else:
        return {
            "chat": 20, "search": 20, "image": 20, "post": 20, "biznes": 20,
            "pdf": 0, "cv": 0, "email": 0, "voice": 0, "pptx": 0, "word": 0
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
    prompt = f"""Extract user info from this conversation.
Current name: {current_name}
Current facts: {current_facts}
User: {user_text}
Assistant: {reply}
Reply ONLY in JSON: {{"name": "name or empty", "facts": "short summary of facts"}}"""
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

# ===== HANDLERS =====
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ensure_user(user_id)
    sub = get_subscription(user_id)
    plan = sub["plan"].upper()
    await update.message.reply_text(
        f"Hello! I am Chatbot 🤖\n"
        f"Your plan: {plan}\n\n"
        f"💬 Chat with me\n"
        f"🌐 Current news, prices, weather\n"
        f"📄 Send PDF to analyze\n"
        f"🖼️ Send image\n"
        f"🎤 Send voice message\n\n"
        f"📊 /pptx — PowerPoint\n"
        f"📝 /word — Word document\n"
        f"👤 /cv — Write CV\n"
        f"📧 /email — Write email\n"
        f"📱 /post — Marketing post\n\n"
        f"💰 /updateplan — Update plan\n"
        f"/help — Help\n"
        f"/reset — Clear history"
    )

async def updateplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    plan = sub["plan"]
    expires = sub["expires_at"].strftime("%d.%m.%Y") if sub.get("expires_at") else "—"

    keyboard = [
        [InlineKeyboardButton("⭐ Buy Standard — 5 USDT/month", callback_data="buy_standard")],
        [InlineKeyboardButton("💎 Buy Premium — 10 USDT/month", callback_data="buy_premium")],
    ]

    await update.message.reply_text(
        f"💰 Plans & Pricing:\n\n"
        f"🆓 FREE — Free\n"
        f"• Chat: 20/day\n"
        f"• Internet search: 20/day\n"
        f"• Image analysis: 20/day\n"
        f"• Marketing post: 20/day\n\n"
        f"⭐ STANDARD — 5 USDT/month\n"
        f"• Everything in Free: 30/day\n"
        f"• PDF analysis: 30/day\n"
        f"• CV writing: 30/day\n"
        f"• Email writing: 30/day\n"
        f"• Memory: Unlimited\n\n"
        f"💎 PREMIUM — 10 USDT/month\n"
        f"• Everything unlimited\n"
        f"• Voice messages\n"
        f"• PowerPoint creation\n"
        f"• Word documents\n"
        f"• Priority support\n\n"
        f"👤 Your plan: {plan.upper()}\n"
        f"📅 Expires: {expires}",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "buy_standard":
        keyboard = [[InlineKeyboardButton("💬 Contact Admin", url=f"https://t.me/{ADMIN_USERNAME}")]]
        await query.message.reply_text(
            f"⭐ STANDARD — 5 USDT/month\n\n"
            f"💳 Pay to card:\n"
            f"`{CARD_NUMBER}`\n\n"
            f"💵 Amount: 5 USDT equivalent in UZS\n\n"
            f"📸 After payment:\n"
            f"1. Screenshot oling\n"
            f"2. Admin ga yuboring\n"
            f"3. Your ID: `{user_id}`\n\n"
            f"✅ Admin 1 soat ichida aktivlashtiradi!",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif query.data == "buy_premium":
        keyboard = [[InlineKeyboardButton("💬 Contact Admin", url=f"https://t.me/{ADMIN_USERNAME}")]]
        await query.message.reply_text(
            f"💎 PREMIUM — 10 USDT/month\n\n"
            f"💳 Pay to card:\n"
            f"`{CARD_NUMBER}`\n\n"
            f"💵 Amount: 10 USDT equivalent in UZS\n\n"
            f"📸 After payment:\n"
            f"1. Screenshot oling\n"
            f"2. Admin ga yuboring\n"
            f"3. Your ID: `{user_id}`\n\n"
            f"✅ Admin 1 soat ichida aktivlashtiradi!",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

# ===== ADMIN =====
async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    await update.message.reply_text(
        "🔧 Admin Panel\n\n"
        "/setplan [user_id] [plan] — Set plan\n"
        "  free | standard | premium\n"
        "  Example: /setplan 123456 premium\n\n"
        "/users — All subscribers\n"
        "/stats — Bot statistics\n"
        "/broadcast [text] — Message to all"
    )

async def setplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    if len(context.args) < 2:
        await update.message.reply_text("Usage: /setplan [user_id] [plan]")
        return
    try:
        target_id = int(context.args[0])
        plan = context.args[1].lower()
        if plan not in ["free", "standard", "premium"]:
            await update.message.reply_text("❌ Plan: free, standard, premium")
            return
        days = 30 if plan != "free" else None
        set_plan(target_id, plan, days)
        await update.message.reply_text(f"✅ User {target_id} → {plan.upper()}")
        emoji = {"free": "🆓", "standard": "⭐", "premium": "💎"}
        try:
            await context.bot.send_message(
                chat_id=target_id,
                text=f"{emoji[plan]} Your plan upgraded to {plan.upper()}! 🎉\n\nUse /updateplan to see details."
            )
        except:
            pass
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def users_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT user_id, plan, expires_at FROM subscriptions ORDER BY plan DESC")
        rows = c.fetchall()
        conn.close()
        if not rows:
            await update.message.reply_text("No users yet.")
            return
        text = f"👥 Total users: {len(rows)}\n\n"
        paid = [r for r in rows if r[1] != "free"]
        text += f"💰 Paid: {len(paid)}\n"
        text += f"🆓 Free: {len(rows) - len(paid)}\n\n"
        for row in paid:
            uid, plan, expires = row
            exp = expires.strftime("%d.%m.%Y") if expires else "—"
            text += f"ID: {uid} | {plan.upper()} | {exp}\n"
        await update.message.reply_text(text)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM subscriptions")
        total = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM subscriptions WHERE plan = 'standard'")
        standard = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM subscriptions WHERE plan = 'premium'")
        premium = c.fetchone()[0]
        conn.close()
        await update.message.reply_text(
            f"📊 Bot Statistics\n\n"
            f"👥 Total users: {total}\n"
            f"🆓 Free: {total - standard - premium}\n"
            f"⭐ Standard: {standard}\n"
            f"💎 Premium: {premium}\n"
            f"💰 Monthly revenue: ~${standard * 5 + premium * 10}"
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
        c.execute("SELECT user_id FROM subscriptions")
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
        "/pptx [topic] — PowerPoint 💎\n"
        "/word [topic] — Word document 💎\n"
        "/cv [info] — CV/Resume ⭐\n"
        "/email [topic] — Email ⭐\n"
        "/post [topic] — Marketing post\n"
        "/biznes [idea] — Business plan\n"
        "/updateplan — Plans & pricing\n"
        "/reset — Clear chat history\n"
        "/help — Help\n\n"
        "⭐ = Standard+\n"
        "💎 = Premium only"
    )

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_histories[update.effective_user.id] = []
    await update.message.reply_text("Chat history cleared ✅")

async def pptx_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    if get_limits(sub["plan"])["pptx"] == 0:
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
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    if get_limits(sub["plan"])["word"] == 0:
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
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    limits = get_limits(sub["plan"])
    if limits["cv"] == 0:
        keyboard = [[InlineKeyboardButton("⭐ Upgrade to Standard", callback_data="buy_standard")]]
        await update.message.reply_text("⭐ CV requires Standard or Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user_id, "cv", limits["cv"]):
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
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    limits = get_limits(sub["plan"])
    if limits["email"] == 0:
        keyboard = [[InlineKeyboardButton("⭐ Upgrade to Standard", callback_data="buy_standard")]]
        await update.message.reply_text("⭐ Email requires Standard or Premium!\nUse /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    if not check_limit(user_id, "email", limits["email"]):
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
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    limits = get_limits(sub["plan"])
    if not check_limit(user_id, "post", limits["post"]):
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
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    limits = get_limits(sub["plan"])
    if not check_limit(user_id, "biznes", limits["biznes"]):
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

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_text = update.message.text
    ensure_user(user_id)
    sub = get_subscription(user_id)
    limits = get_limits(sub["plan"])

    if not check_limit(user_id, "chat", limits["chat"]):
        keyboard = [[InlineKeyboardButton("💰 Upgrade Plan", callback_data="buy_standard")]]
        await update.message.reply_text("❌ Daily limit reached! Use /updateplan to upgrade.", reply_markup=InlineKeyboardMarkup(keyboard))
        return

    memory = get_memory(user_id)
    memory_context = ""
    if memory and (memory.get("name") or memory.get("facts")):
        memory_context = f"User info — name: {memory['name']}, facts: {memory['facts']}. "

    if user_id not in user_histories:
        user_histories[user_id] = [{"role": "system", "content": (
            "You are a professional AI assistant. "
            "Always reply in the same language the user writes in. "
            "Keep answers short, clear and natural. "
            + memory_context
        )}]

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")

    try:
        if needs_search(user_text) and limits["search"] != 0:
            search_results = tavily.search(query=user_text, max_results=3)
            search_content = "\n\n".join([f"Source: {r['url']}\n{r['content']}" for r in search_results.get("results", [])])
            message_content = f"User question: {user_text}\n\nWeb search results:\n{search_content}\n\nGive a clear answer based on results."
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
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    if get_limits(sub["plan"])["voice"] == 0:
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
    user_id = update.effective_user.id
    ensure_user(user_id)
    sub = get_subscription(user_id)
    limits = get_limits(sub["plan"])
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
    user_id = update.effective_user.id
    ensure_user(user_id)
    sub = get_subscription(user_id)
    limits = get_limits(sub["plan"])
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
        user_histories[user_id].append({"role": "user", "content": f"PDF content:\n\n{text}\n\nRequest: {caption}"})
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
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("setplan", setplan_command))
    app.add_handler(CommandHandler("users", users_command))
    app.add_handler(CommandHandler("stats", stats_command))
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