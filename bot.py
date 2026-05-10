import os
import requests
import fitz
import json
import psycopg2
from groq import Groq
from tavily import TavilyClient
from dotenv import load_dotenv
from telegram import Update, BotCommand
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
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

client = Groq(api_key=GROQ_API_KEY)
tavily = TavilyClient(api_key=TAVILY_API_KEY)

user_histories = {}

SEARCH_KEYWORDS = [
    "today", "now", "current", "latest", "news", "price", "rate", "weather",
    "bugun", "hozir", "narx", "kurs", "yangilik", "ob-havo", "oxirgi",
    "сегодня", "сейчас", "курс", "цена", "новости"
]

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
    conn.commit()
    conn.close()

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

    prompt = f"""
Based on this conversation, extract and update user information.
Current known facts: {current_facts}
Current name: {current_name}

New conversation:
User: {user_text}
Assistant: {reply}

Reply ONLY in JSON:
{{"name": "user name or empty string", "facts": "updated facts as a short summary"}}
"""
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

    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(0x1E, 0x1E, 0x2E)
    title_box = slide.shapes.title
    title_box.text = title
    title_box.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    title_box.text_frame.paragraphs[0].font.size = Pt(40)
    title_box.text_frame.paragraphs[0].font.bold = True

    for s in slides_data:
        sl = prs.slide_layouts[1]
        slide = prs.slides.add_slide(sl)
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

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Hello! I am your AI assistant 🤖\n\n"
        "💬 Chat with me\n"
        "🌐 Current news, prices, weather\n"
        "📄 Send PDF to analyze\n"
        "🖼️ Send image\n"
        "🎤 Send voice message\n\n"
        "📊 /pptx — PowerPoint\n"
        "📝 /word — Word document\n"
        "👤 /cv — Write CV/Resume\n"
        "📧 /email — Write email\n"
        "📱 /post — Marketing post\n"
        "💼 /biznes — Business plan\n\n"
        "/help — Help\n"
        "/reset — Clear history"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📌 Commands:\n\n"
        "/pptx [topic] — PowerPoint\n"
        "/word [topic] — Word document\n"
        "/cv [your info] — CV/Resume\n"
        "/email [topic] — Professional email\n"
        "/post [topic] — Marketing post\n"
        "/biznes [idea] — Business plan\n"
        "/reset — Clear chat history\n"
        "/help — Help\n\n"
        "Examples:\n"
        "/cv Python developer, 3 years experience\n"
        "/email follow up after interview\n"
        "/post new coffee shop opening\n"
        "/biznes online clothing store"
    )

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_histories[user_id] = []
    await update.message.reply_text("Chat history cleared ✅")

async def pptx_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /pptx artificial intelligence")
        return
    await update.message.reply_text(f"⏳ Creating presentation on '{topic}'...")
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="upload_document")
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
            await update.message.reply_document(document=f, filename=f"{topic}.pptx", caption=f"✅ Ready!")
        os.remove(path)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def word_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /word business plan")
        return
    await update.message.reply_text(f"⏳ Creating document on '{topic}'...")
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="upload_document")
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
            await update.message.reply_document(document=f, filename=f"{topic}.docx", caption=f"✅ Ready!")
        os.remove(path)
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def cv_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    info = " ".join(context.args)
    if not info:
        await update.message.reply_text("Example: /cv Python developer, 3 years experience, Tashkent")
        return
    await update.message.reply_text("⏳ Writing your CV...")
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        prompt = f"""Write a professional CV/Resume for: {info}
Format it nicely with sections: Summary, Experience, Skills, Education.
Make it ATS-friendly and professional."""
        reply = await ai_generate(prompt)
        await update.message.reply_text(f"✅ Your CV:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def email_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /email follow up after job interview")
        return
    await update.message.reply_text("⏳ Writing email...")
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        prompt = f"""Write a professional email about: {topic}
Include: Subject line, greeting, body, closing.
Make it professional and concise."""
        reply = await ai_generate(prompt)
        await update.message.reply_text(f"✅ Your email:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def post_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Example: /post new coffee shop opening in Tashkent")
        return
    await update.message.reply_text("⏳ Writing marketing post...")
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        prompt = f"""Write an engaging social media marketing post about: {topic}
Include: Hook, main content, call to action, relevant hashtags.
Make it catchy and professional."""
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
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    try:
        prompt = f"""Write a detailed business plan for: {idea}
Include: Executive Summary, Market Analysis, Products/Services, 
Marketing Strategy, Financial Plan, Operations.
Make it professional and realistic."""
        reply = await ai_generate(prompt)
        await update.message.reply_text(f"✅ Business plan:\n\n{reply}")
    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_text = update.message.text

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
                    "Never add unnecessary information about yourself. "
                    + memory_context
                )
            }
        ]

    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")

    try:
        if needs_search(user_text):
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
        user_id = update.effective_user.id
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
        BotCommand("pptx", "Create PowerPoint"),
        BotCommand("word", "Create Word document"),
        BotCommand("cv", "Write CV/Resume"),
        BotCommand("email", "Write professional email"),
        BotCommand("post", "Write marketing post"),
        BotCommand("biznes", "Write business plan"),
        BotCommand("reset", "Clear chat history"),
        BotCommand("help", "Help"),
    ])

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("pptx", pptx_command))
    app.add_handler(CommandHandler("word", word_command))
    app.add_handler(CommandHandler("cv", cv_command))
    app.add_handler(CommandHandler("email", email_command))
    app.add_handler(CommandHandler("post", post_command))
    app.add_handler(CommandHandler("biznes", biznes_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.PHOTO, handle_image))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_document))
    print("Bot is running... ✅")
    app.run_polling()

if __name__ == "__main__":
    main()