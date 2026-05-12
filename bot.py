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
ADMIN_USERNAME = "@temur_uzb7779"
STANDARD_PRICE = 5
PREMIUM_PRICE = 10

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
            started_at TIMESTAMP DEFAULT NOW(),
            expires_at TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

def get_subscription(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT plan, started_at, expires_at FROM subscriptions WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        conn.close()

        now = datetime.now()

        if not row:
            # Yangi foydalanuvchi — 20 kunlik FREE berish
            expires_at = now + timedelta(days=20)
            conn = get_conn()
            c = conn.cursor()
            c.execute("""
                INSERT INTO subscriptions (user_id, plan, started_at, expires_at)
                VALUES (%s, 'free', %s, %s)
            """, (user_id, now, expires_at))
            conn.commit()
            conn.close()
            return {"plan": "free", "started_at": now, "expires_at": expires_at}

        plan, started_at, expires_at = row

        # Muddati tugagan
        if expires_at and now > expires_at:
            if plan in ("standard", "premium"):
                # Pullik plan tugadi → FREE ga qaytadi, 20 kunlik
                new_expires = now + timedelta(days=20)
                conn = get_conn()
                c = conn.cursor()
                c.execute("""
                    UPDATE subscriptions SET plan='free', started_at=%s, expires_at=%s WHERE user_id=%s
                """, (now, new_expires, user_id))
                conn.commit()
                conn.close()
                return {"plan": "free", "started_at": now, "expires_at": new_expires}
            else:
                # FREE tugadi → avtomatik yana 20 kunlik FREE
                new_expires = now + timedelta(days=20)
                conn = get_conn()
                c = conn.cursor()
                c.execute("""
                    UPDATE subscriptions SET started_at=%s, expires_at=%s WHERE user_id=%s
                """, (now, new_expires, user_id))
                conn.commit()
                conn.close()
                return {"plan": "free", "started_at": now, "expires_at": new_expires}

        return {"plan": plan, "started_at": started_at, "expires_at": expires_at}
    except Exception as e:
        return {"plan": "free", "started_at": datetime.now(), "expires_at": datetime.now() + timedelta(days=20)}

def set_plan(user_id, plan, days):
    try:
        now = datetime.now()
        expires_at = now + timedelta(days=days)
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO subscriptions (user_id, plan, started_at, expires_at)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (user_id) DO UPDATE SET plan=%s, started_at=%s, expires_at=%s
        """, (user_id, plan, now, expires_at, plan, now, expires_at))
        conn.commit()
        conn.close()
    except:
        pass

def get_plan_limits(plan, feature):
    limits = {
        "free": {
            "chat": -1, "search": -1, "image": -1, "post": -1, "biznes": -1,
            "pdf": 0, "cv": 0, "email": 0, "voice": 0, "pptx": 0, "word": 0
        },
        "standard": {
            "chat": -1, "search": -1, "image": -1, "post": -1, "biznes": -1,
            "pdf": -1, "cv": -1, "email": -1, "voice": 0, "pptx": 0, "word": 0
        },
        "premium": {
            "chat": -1, "search": -1, "image": -1, "post": -1, "biznes": -1,
            "pdf": -1, "cv": -1, "email": -1, "voice": -1, "pptx": -1, "word": -1
        }
    }
    return limits.get(plan, limits["free"]).get(feature, 0)

def get_memory(user_id):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT name, facts FROM user_memory WHERE user_id = %s", (user_id,))
        row = c.fetchone()
        conn.close()
        if row:
            return {"name": row[0], "facts": row[1]}
        return None
    except:
        return None

def save_memory(user_id, name, facts):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO user_memory (user_id, name, facts)
            VALUES (%s, %s, %s)
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
    prompt = f"""Based on this conversation, extract and update user information.
Current known facts: {current_facts}
Current name: {current_name}
User: {user_text}
Assistant: {reply}
Reply ONLY in JSON:
{{"name": "user name or empty string", "facts": "updated facts as a short summary"}}"""
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

def days_left(expires_at):
    if not expires_at:
        return 0
    delta = expires_at - datetime.now()
    return max(0, delta.days)

# ===== PPTX & DOCX =====
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
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()
        for i, point in enumerate(s["points"]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"• {point}"
            p.font.color.rgb = RGBColor(0xCD, 0xD6, 0xF4)
            p.font.size = Pt(18)
    path = "presentation.pptx"
    prs.save(path)
    return path

def create_docx(title, sections):
    doc = Document()
    h = doc.add_heading(title, 0)
    h.runs[0].font.color.rgb = DocRGB(0x1E, 0x3A, 0x8A)
    for section in sections:
        doc.add_heading(section["title"], level=1)
        for point in section["points"]:
            doc.add_paragraph(point, style="List Bullet")
        doc.add_paragraph()
    path = "document.docx"
    doc.save(path)
    return path

async def generate_content(topic, doc_type):
    if doc_type == "pptx":
        prompt = f"""Create presentation content for: '{topic}'.
Reply ONLY in JSON:
{{"title": "title", "slides": [{{"title": "s1", "points": ["p1","p2","p3"]}}]}}
At least 5 slides. Same language as topic."""
    else:
        prompt = f"""Create Word document content for: '{topic}'.
Reply ONLY in JSON:
{{"title": "title", "sections": [{{"title": "s1", "points": ["p1","p2","p3"]}}]}}
At least 4 sections. Same language as topic."""
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

async def ai_generate(prompt):
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

# ===== COMMANDS =====
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    plan = sub["plan"].upper()
    left = days_left(sub["expires_at"])

    plan_icons = {"FREE": "🆓", "STANDARD": "⭐️", "PREMIUM": "💎"}
    icon = plan_icons.get(plan, "🆓")

    await update.message.reply_text(
        f"Hello! I am Chatbot 🤖\n"
        f"Your plan: {icon} {plan} — {left} days left\n\n"
        "💬 Chat with me\n"
        "🌐 Current news, prices, weather\n"
        "📄 Send PDF to analyze\n"
        "🖼️ Send image\n"
        "🎤 Send voice message\n\n"
        "📊 /pptx — PowerPoint\n"
        "📝 /word — Word document\n"
        "👤 /cv — Write CV\n"
        "📧 /email — Write email\n"
        "📱 /post — Marketing post\n\n"
        "💰 /pricing — Plans & pricing\n"
        "👤 /myplan — My current plan\n"
        "/help — Help\n"
        "/reset — Clear history"
    )

async def pricing_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    plan = sub["plan"]

    keyboard = []
    if plan != "standard":
        keyboard.append([InlineKeyboardButton(f"⭐️ Buy STANDARD — {STANDARD_PRICE} USDT/month", callback_data="plan_buy_standard")])
    if plan != "premium":
        keyboard.append([InlineKeyboardButton(f"💎 Buy PREMIUM — {PREMIUM_PRICE} USDT/month", callback_data="plan_buy_premium")])

    reply_markup = InlineKeyboardMarkup(keyboard)

    await update.message.reply_text(
        "💰 Plans & Pricing:\n\n"
        "🆓 FREE — Free (20 days, then auto-renews)\n"
        "• Chat: Unlimited\n"
        "• Internet search: Unlimited\n"
        "• Image analysis: Unlimited\n"
        "• Marketing post: Unlimited\n"
        "• Business plan: Unlimited\n\n"
        f"⭐️ STANDARD — {STANDARD_PRICE} USDT/month\n"
        "• Everything in Free\n"
        "• PDF analysis\n"
        "• CV writing\n"
        "• Email writing\n"
        "• Memory: Unlimited\n\n"
        f"💎 PREMIUM — {PREMIUM_PRICE} USDT/month\n"
        "• Everything in Standard\n"
        "• Voice messages\n"
        "• PowerPoint creation\n"
        "• Word documents\n"
        "• Priority support\n",
        reply_markup=reply_markup
    )

async def myplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    plan = sub["plan"].upper()
    left = days_left(sub["expires_at"])
    expires = sub["expires_at"].strftime("%d.%m.%Y") if sub.get("expires_at") else "—"

    plan_icons = {"FREE": "🆓", "STANDARD": "⭐️", "PREMIUM": "💎"}
    icon = plan_icons.get(plan, "🆓")

    await update.message.reply_text(
        f"👤 Your plan: {icon} {plan}\n"
        f"📅 Expires: {expires}\n"
        f"⏳ Days left: {left}\n\n"
        "Use /pricing to upgrade!"
    )

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "plan_buy_standard":
        await query.message.reply_text(
            f"⭐️ STANDARD plan — {STANDARD_PRICE} USDT/month\n\n"
            f"💳 Pay to card: `{CARD_NUMBER}`\n"
            f"💵 Amount: {STANDARD_PRICE} USDT equivalent\n\n"
            f"After payment, send screenshot to {ADMIN_USERNAME}\n"
            f"Your ID: `{user_id}`\n\n"
            "Admin will activate your plan within 1 hour! ✅",
            parse_mode="Markdown"
        )
    elif query.data == "plan_buy_premium":
        await query.message.reply_text(
            f"💎 PREMIUM plan — {PREMIUM_PRICE} USDT/month\n\n"
            f"💳 Pay to card: `{CARD_NUMBER}`\n"
            f"💵 Amount: {PREMIUM_PRICE} USDT equivalent\n\n"
            f"After payment, send screenshot to {ADMIN_USERNAME}\n"
            f"Your ID: `{user_id}`\n\n"
            "Admin will activate your plan within 1 hour! ✅",
            parse_mode="Markdown"
        )

# ===== ADMIN =====
async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    await update.message.reply_text(
        "🔧 Admin Panel\n\n"
        "/setplan [user_id] [plan] — Set user plan\n"
        "  Plans: free, standard, premium\n"
        "  Example: /setplan 123456789 premium\n\n"
        "/users — Show all subscribers\n"
        "/broadcast [message] — Send message to all users"
    )

async def setplan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    if len(context.args) < 2:
        await update.message.reply_text("Usage: /setplan [user_id] [plan]")
        return
    target_id = int(context.args[0])
    plan = context.args[1].lower()
    if plan not in ["free", "standard", "premium"]:
        await update.message.reply_text("❌ Plan must be: free, standard, or premium")
        return
    days = 20 if plan == "free" else 30
    set_plan(target_id, plan, days)
    await update.message.reply_text(f"✅ User {target_id} plan set to {plan.upper()}!")
    try:
        plan_emoji = {"free": "🆓", "standard": "⭐️", "premium": "💎"}
        await context.bot.send_message(
            chat_id=target_id,
            text=f"{plan_emoji[plan]} Your plan has been upgraded to {plan.upper()}! Thank you! 🎉\nUse /myplan to see details."
        )
    except:
        pass

async def users_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("❌ Access denied!")
        return
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT user_id, plan, expires_at FROM subscriptions WHERE plan != 'free' ORDER BY plan")
        rows = c.fetchall()
        conn.close()
        if not rows:
            await update.message.reply_text("No paid subscribers yet.")
            return
        text = "👥 Paid subscribers:\n\n"
        for row in rows:
            uid, plan, expires = row
            expires_str = expires.strftime("%d.%m.%Y") if expires else "—"
            text += f"ID: {uid} | {plan.upper()} | until {expires_str}\n"
        await update.message.reply_text(text)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
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
        "/pptx [topic] — PowerPoint (Premium)\n"
        "/word [topic] — Word document (Premium)\n"
        "/cv [info] — CV/Resume (Standard+)\n"
        "/email [topic] — Email (Standard+)\n"
        "/post [topic] — Marketing post\n"
        "/biznes [idea] — Business plan\n"
        "/pricing — Plans & pricing\n"
        "/myplan — My current plan\n"
        "/reset — Clear chat history\n"
        "/help — Help"
    )

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_histories[user_id] = []
    await update.message.reply_text("Chat history cleared ✅")

async def pptx_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    if get_plan_limits(sub["plan"], "pptx") == 0:
        await update.message.reply_text("💎 This feature requires Premium plan!\nUse /pricing to upgrade.")
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
    if get_plan_limits(sub["plan"], "word") == 0:
        await update.message.reply_text("💎 This feature requires Premium plan!\nUse /pricing to upgrade.")
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
    if get_plan_limits(sub["plan"], "cv") == 0:
        await update.message.reply_text("⭐️ This feature requires Standard or Premium plan!\nUse /pricing to upgrade.")
        return
    info = " ".join(context.args)
    if not info:
        await update.message.reply_text("Example: /cv Python developer, 3 years experience")
        return
    await update.message.reply_text("⏳ Writing your CV...")
    try:
        prompt = f"Write a professional CV/Resume for: {info}\nFormat: Summary, Experience, Skills, Education. ATS-friendly."
        reply = await ai_generate(prompt)
        await update.message.reply_text(f"✅ Your CV:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def email_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    if get_plan_limits(sub["plan"], "email") == 0:
        await update.message.reply_text("⭐️ This feature requires Standard or Premium plan!\nUse /pricing to upgrade.")
        return
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /email follow up after job interview")
        return
    await update.message.reply_text("⏳ Writing email...")
    try:
        prompt = f"Write a professional email about: {topic}\nInclude: Subject, greeting, body, closing."
        reply = await ai_generate(prompt)
        await update.message.reply_text(f"✅ Your email:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def post_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /post new coffee shop opening")
        return
    await update.message.reply_text("⏳ Writing marketing post...")
    try:
        prompt = f"Write an engaging social media marketing post about: {topic}\nInclude: Hook, content, call to action, hashtags."
        reply = await ai_generate(prompt)
        await update.message.reply_text(f"✅ Your post:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def biznes_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    idea = " ".join(context.args)
    if not idea:
        await update.message.reply_text("Example: /biznes online clothing store")
        return
    await update.message.reply_text("⏳ Writing business plan...")
    try:
        prompt = f"Write a detailed business plan for: {idea}\nInclude: Executive Summary, Market Analysis, Products/Services, Marketing Strategy, Financial Plan."
        reply = await ai_generate(prompt)
        await update.message.reply_text(f"✅ Business plan:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_text = update.message.text
    sub = get_subscription(user_id)

    memory = get_memory(user_id)
    memory_context = ""
    if memory and (memory["name"] or memory["facts"]):
        memory_context = f"User info — name: {memory['name']}, facts: {memory['facts']}. "

    if user_id not in user_histories:
        user_histories[user_id] = [
            {
                "role": "system",
                "content": (
                    "You are a professional AI assistant. "
                    "Always reply in the same language the user writes in. "
                    "Keep answers short, clear and natural. "
                    + memory_context
                )
            }
        ]

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")

    try:
        if needs_search(user_text):
            try:
                search_results = tavily.search(query=user_text, max_results=3)
                search_content = "\n\n".join([
                    f"Source: {r['url']}\n{r['content']}"
                    for r in search_results.get("results", [])
                ])
                message_content = (
                    f"User question: {user_text}\n\n"
                    f"Web search results:\n{search_content}\n\n"
                    f"Based on these results, give a clear and accurate answer."
                )
            except:
                message_content = user_text
        else:
            message_content = user_text

        user_histories[user_id].append({"role": "user", "content": message_content})
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=user_histories[user_id]
        )
        reply = response.choices[0].message.content
        user_histories[user_id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
        await update_memory(user_id, user_text, reply)

    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sub = get_subscription(user_id)
    if get_plan_limits(sub["plan"], "voice") == 0:
        await update.message.reply_text("💎 Voice messages require Premium plan!\nUse /pricing to upgrade.")
        return
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        voice = update.message.voice
        file = await context.bot.get_file(voice.file_id)
        audio_data = requests.get(file.file_path).content
        with open("voice.ogg", "wb") as f:
            f.write(audio_data)
        with open("voice.ogg", "rb") as f:
            transcription = client.audio.transcriptions.create(
                file=("voice.ogg", f.read()),
                model="whisper-large-v3",
            )
        user_text = transcription.text
        await update.message.reply_text(f"🎤 You said: {user_text}")
        if user_id not in user_histories:
            user_histories[user_id] = [{"role": "system", "content": "You are a professional AI assistant. Always reply in the same language the user writes in."}]
        user_histories[user_id].append({"role": "user", "content": user_text})
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=user_histories[user_id]
        )
        reply = response.choices[0].message.content
        user_histories[user_id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
    sub = get_subscription(user_id)
    if get_plan_limits(sub["plan"], "pdf") == 0:
        await update.message.reply_text("⭐️ PDF analysis requires Standard or Premium plan!\nUse /pricing to upgrade.")
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
        text = ""
        for page in pdf:
            text += page.get_text()
        pdf.close()
        os.remove("temp.pdf")
        if len(text) > 12000:
            text = text[:12000] + "..."
        caption = update.message.caption or "Summarize this document and explain the key points."
        if user_id not in user_histories:
            user_histories[user_id] = [{"role": "system", "content": "You are a professional AI assistant. Always reply in the same language the user writes in."}]
        user_histories[user_id].append({"role": "user", "content": f"PDF content:\n\n{text}\n\nUser request: {caption}"})
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=user_histories[user_id]
        )
        reply = response.choices[0].message.content
        user_histories[user_id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def post_init(app):
    init_db()
    await app.bot.set_my_commands([
        BotCommand("start", "Start the bot"),
        BotCommand("pricing", "Plans & pricing"),
        BotCommand("myplan", "My current plan"),
        BotCommand("pptx", "Create PowerPoint (Premium)"),
        BotCommand("word", "Create Word document (Premium)"),
        BotCommand("cv", "Write CV (Standard+)"),
        BotCommand("email", "Write email (Standard+)"),
        BotCommand("post", "Marketing post"),
        BotCommand("biznes", "Business plan"),
        BotCommand("reset", "Clear chat history"),
        BotCommand("help", "Help"),
    ])

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("pricing", pricing_command))
    app.add_handler(CommandHandler("myplan", myplan_command))
    app.add_handler(CommandHandler("pptx", pptx_command))
    app.add_handler(CommandHandler("word", word_command))
    app.add_handler(CommandHandler("cv", cv_command))
    app.add_handler(CommandHandler("email", email_command))
    app.add_handler(CommandHandler("post", post_command))
    app.add_handler(CommandHandler("biznes", biznes_command))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("setplan", setplan_command))
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