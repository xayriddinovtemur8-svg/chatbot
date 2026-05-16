import os
import asyncio
import schedule
import time
import requests
from groq import Groq
from tavily import TavilyClient
from dotenv import load_dotenv
from telegram import Bot
from datetime import datetime
from youtubesearchpython import VideosSearch

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
CHANNEL_ID = os.getenv("CHANNEL_ID")
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY")

groq_client = Groq(api_key=GROQ_API_KEY)
tavily = TavilyClient(api_key=TAVILY_API_KEY)
bot = Bot(token=TELEGRAM_TOKEN)

TOPICS = [
    ("world breaking news today", "🌍", "world news", "#DunyoYangiliklari #BreakingNews"),
    ("global politics news today", "🏛️", "politics government", "#Siyosat #Xalqaro #Dunyo"),
    ("world economy financial news", "💹", "economy finance", "#Iqtisodiyot #Moliya #Dunyo"),
    ("technology AI news today", "🤖", "technology computer", "#Texnologiya #AI #Innovatsiya"),
    ("science space discovery news", "🔬", "science space", "#Fan #Kashfiyot #Kosmik"),
    ("climate environment news today", "🌱", "nature environment", "#Iqlim #Ekologiya #Tabiat"),
    ("world sports news today", "⚽", "sports stadium", "#Sport #Dunyo #Chempionat"),
    ("middle east asia news today", "🗺️", "middle east city", "#OrtaShark #Osiyo #Xalqaro"),
    ("europe america latest news", "🌐", "europe city", "#Yevropa #Amerika #Dunyo"),
    ("health medicine news today", "🏥", "hospital medicine", "#Soglik #Tibbiyot #Yangilik"),
]

topic_index = [0]


def get_pexels_image(query):
    try:
        headers = {"Authorization": PEXELS_API_KEY}
        url = f"https://api.pexels.com/v1/search?query={query}&per_page=5&orientation=landscape"
        resp = requests.get(url, headers=headers, timeout=10)
        data = resp.json()
        photos = data.get("photos", [])
        if photos:
            return photos[0]["src"]["large"]
    except Exception as e:
        print(f"Pexels xato: {e}")
    return None


def generate_text(news_content, emoji, hashtags):
    prompt = f"""Siz professional o'zbek yangiliklar muharriri siz.
Quyidagi dunyo yangiliklarini o'zbek tilida Telegram kanal posti qilib yozing.

Yangiliklar:
{news_content}

QOIDALAR:
1. FAQAT O'ZBEK TILIDA yozing — hech qanday ingliz so'z YOZMANG
2. {emoji} emoji bilan SARLAVHA bilan boshlang
3. 2-3 qisqa paragraf yozing
4. Emoji ishlating
5. Oxirida: {hashtags}
6. Matn 800 belgidan OSHMASIN

FAQAT POSTNI YOZING:"""

    response = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=350,
    )
    text = response.choices[0].message.content.strip()
    if len(text) > 1020:
        text = text[:1020] + "..."
    return text


def find_youtube_video(query):
    try:
        search = VideosSearch(query + " news", limit=5)
        results = search.result().get("result", [])
        for video in results:
            duration = video.get("duration", "")
            if duration and ":" in duration:
                parts = duration.split(":")
                if len(parts) == 2 and int(parts[0]) <= 10:
                    return f"https://youtu.be/{video['id']}"
        if results:
            return f"https://youtu.be/{results[0]['id']}"
    except Exception as e:
        print(f"YouTube xato: {e}")
    return None


async def send_image_post():
    try:
        topic, emoji, pexels_query, hashtags = TOPICS[topic_index[0] % len(TOPICS)]
        topic_index[0] += 1

        print(f"[{datetime.now().strftime('%H:%M:%S')}] 📸 Rasm post: {topic}")

        results = tavily.search(query=topic, max_results=3)
        articles = results.get("results", [])
        if not articles:
            print("Maqola topilmadi!")
            return

        news_content = ""
        for i, a in enumerate(articles[:3]):
            news_content += f"{i+1}. {a.get('title', '')}\n{a.get('content', '')[:400]}\n\n"

        caption = generate_text(news_content, emoji, hashtags)

        # Pexels dan rasm olish
        image_url = get_pexels_image(pexels_query)

        if image_url:
            try:
                img_bytes = requests.get(image_url, timeout=10).content
                await bot.send_photo(
                    chat_id=CHANNEL_ID,
                    photo=img_bytes,
                    caption=caption,
                )
                print("✅ Rasm + matn yuborildi!")
                return
            except Exception as e:
                print(f"Rasm xato: {e}")

        # Rasm topilmasa faqat matn
        await bot.send_message(chat_id=CHANNEL_ID, text=caption)
        print("✅ Faqat matn yuborildi")

    except Exception as e:
        print(f"❌ Xato: {e}")


async def send_video_post():
    try:
        topic, emoji, pexels_query, hashtags = TOPICS[topic_index[0] % len(TOPICS)]
        topic_index[0] += 1

        print(f"[{datetime.now().strftime('%H:%M:%S')}] 🎥 Video post: {topic}")

        results = tavily.search(query=topic, max_results=3)
        articles = results.get("results", [])
        if not articles:
            print("Maqola topilmadi!")
            return

        news_content = ""
        for i, a in enumerate(articles[:3]):
            news_content += f"{i+1}. {a.get('title', '')}\n{a.get('content', '')[:400]}\n\n"

        caption = generate_text(news_content, emoji, hashtags)

        video_url = find_youtube_video(topic)

        if video_url:
            full_text = f"{caption}\n\n🎥 {video_url}"
            await bot.send_message(chat_id=CHANNEL_ID, text=full_text[:4096])
            print("✅ Video link + matn yuborildi!")
        else:
            await bot.send_message(chat_id=CHANNEL_ID, text=caption)
            print("✅ Faqat matn yuborildi (video topilmadi)")

    except Exception as e:
        print(f"❌ Xato: {e}")


def job_image():
    asyncio.run(send_image_post())


def job_video():
    asyncio.run(send_video_post())


def main():
    print("🤖 Dunyo Yangiliklari Boti ishga tushdi!")
    print(f"📢 Kanal: {CHANNEL_ID}")
    print("📅 Jadval: 09:00 📸 | 13:00 🎥 | 17:00 📸 | 21:00 🎥\n")

    print("🚀 Birinchi test post yuborilmoqda...")
    job_image()

    schedule.every().day.at("09:00").do(job_image)
    schedule.every().day.at("13:00").do(job_video)
    schedule.every().day.at("17:00").do(job_image)
    schedule.every().day.at("21:00").do(job_video)

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()