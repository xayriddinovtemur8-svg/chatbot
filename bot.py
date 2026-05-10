import os
import requests
import fitz  # pymupdf
from groq import Groq
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
import json

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

client = Groq(api_key=GROQ_API_KEY)

user_histories = {}

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
        prompt = f"""
Create presentation content for the topic: '{topic}'.
Reply ONLY in JSON format, nothing else:
{{
  "title": "Presentation title",
  "slides": [
    {{"title": "Slide title", "points": ["point 1", "point 2", "point 3"]}},
    {{"title": "Slide title", "points": ["point 1", "point 2", "point 3"]}}
  ]
}}
Create at least 5 slides. Use the same language as the topic."""
    else:
        prompt = f"""
Create Word document content for the topic: '{topic}'.
Reply ONLY in JSON format, nothing else:
{{
  "title": "Document title",
  "sections": [
    {{"title": "Section title", "points": ["sentence 1", "sentence 2", "sentence 3"]}},
    {{"title": "Section title", "points": ["sentence 1", "sentence 2", "sentence 3"]}}
  ]
}}
Create at least 4 sections. Use the same language as the topic."""

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Hello! I am your AI assistant 🤖\n\n"
        "📊 /pptx — Create PowerPoint presentation\n"
        "📝 /word — Create Word document\n"
        "🎤 Send a voice message\n"
        "🖼️ Send an image\n\n"
        "/help — Help\n"
        "/reset — Clear chat history"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📌 Commands:\n\n"
        "/pptx [topic] — Create PowerPoint\n"
        "  Example: /pptx artificial intelligence\n\n"
        "/word [topic] — Create Word document\n"
        "  Example: /word business plan\n\n"
        "/reset — Clear chat history\n"
        "/help — Help"
    )

async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_histories[user_id] = []
    await update.message.reply_text("Chat history cleared ✅")

async def pptx_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Please provide a topic!\nExample: /pptx artificial intelligence")
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
            await update.message.reply_document(
                document=f,
                filename=f"{topic}.pptx",
                caption=f"✅ Presentation on '{topic}' is ready!"
            )
        os.remove(path)

    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def word_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = " ".join(context.args)
    if not topic:
        await update.message.reply_text("Please provide a topic!\nExample: /word business plan")
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
            await update.message.reply_document(
                document=f,
                filename=f"{topic}.docx",
                caption=f"✅ Document on '{topic}' is ready!"
            )
        os.remove(path)

    except Exception as e:
        await update.message.reply_text(f"Error: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_text = update.message.text

    if user_id not in user_histories:
        user_histories[user_id] = [
            {
                "role": "system",
                "content": (
                    "You are a professional AI assistant. "
                    "Always reply in the same language the user writes in. "
                    "Keep answers short, clear and natural. "
                    "Never add unnecessary information about yourself. "
                    "If there is no question, do not add extra information."
                )
            }
        ]

    user_histories[user_id].append({"role": "user", "content": user_text})
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=user_histories[user_id]
        )
        reply = response.choices[0].message.content
        user_histories[user_id].append({"role": "assistant", "content": reply})
        await update.message.reply_text(reply)
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
            user_histories[user_id] = [
                {
                    "role": "system",
                    "content": (
                        "You are a professional AI assistant. "
                        "Always reply in the same language the user writes in. "
                        "Keep answers short, clear and natural."
                    )
                }
            ]

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
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": file.file_path}},
                    {"type": "text", "text": caption}
                ]
            }]
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
            user_histories[user_id] = [
                {
                    "role": "system",
                    "content": (
                        "You are a professional AI assistant. "
                        "Always reply in the same language the user writes in. "
                        "Keep answers short, clear and natural."
                    )
                }
            ]

        user_histories[user_id].append({
            "role": "user",
            "content": f"PDF content:\n\n{text}\n\nUser request: {caption}"
        })

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
    await app.bot.set_my_commands([
        BotCommand("start", "Start the bot"),
        BotCommand("pptx", "Create PowerPoint"),
        BotCommand("word", "Create Word document"),
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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.PHOTO, handle_image))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_document))
    print("Bot is running... ✅")
    app.run_polling()

if __name__ == "__main__":
    main()