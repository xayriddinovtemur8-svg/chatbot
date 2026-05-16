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

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
CHANNEL_ID = os.getenv("CHANNEL_ID")

client = Groq(api_key=GROQ_API_KEY)
tavily = TavilyClient(api_key=TAVILY_API_KEY)
bot = Bot(token=TELEGRAM_TOKEN)

NEWS_TOPICS = [
    "O'zbekiston yangiliklari bugun",
    "Uzbekistan news today",
    "O'zbekiston iqtisodiyot yangiliklari",
    "Uzbekistan politics news today",
    "O'zbekiston sport yangiliklari",
    "Toshkent yangiliklari bugun",
    "O'zbekiston texnologiya yangiliklari",
    "Uzbekistan economy news today",
]

TOPIC_EMOJIS = {
    "O'zbekiston yangiliklari bugun": "🇺🇿",
    "Uzbekistan news today": "🌍",
    "O'zbekiston iqtisodiyot yangiliklari": "💰",
    "Uzbekistan politics news today": "🏛️",
    "O'zbekiston sport yangiliklari": "⚽",
    "Toshkent yangiliklari bugun": "🏙️",
    "O'zbekiston texnologiya yangiliklari": "💻",
    "Uzbekistan economy news today": "📈",
}

TOPIC_HASHTAGS = {
    "O'zbekiston yangiliklari bugun": "#Uzbekistan #Ozbekiston #Yangiliklar",
    "Uzbekistan news today": "#Uzbekistan #News #CentralAsia",
    "O'zbekiston iqtisodiyot yangiliklari": "#Iqtisodiyot #Economy #Uzbekistan",
    "Uzbekistan politics news today": "#Politics #Uzbekistan #Government",
    "O'zbekiston sport yangiliklari": "#Sport #Uzbekistan #Football",
    "Toshkent yangiliklari bugun": "#Toshkent #Tashkent #Uzbekistan",
    "O'zbekiston texnologiya yangiliklari": "#Tech #Texnologiya #Uzbekistan",
    "Uzbekistan economy news today": "#Economy #Finance #Uzbekistan",
}

post_counter = [0]

async def fetch_and_post_news():
    try:
        topic = NEWS_TOPICS[post_counter[0] % len(NEWS_TOPICS)]
        emoji = TOPIC_EMOJIS.get(topic, "🇺🇿")
        hashtags = TOPIC_HASHTAGS.get(topic, "#Uzbekistan #Yangiliklar")
        post_counter[0] += 1

        print(f"[{datetime.now().strftime('%H:%M:%S')}] Fetching: {topic}")

        results = tavily.search(query=topic, max_results=3)
        articles = results.get("results", [])

        if not articles:
            print("No articles found!")
            return

        news_content = ""
        for i, article in enumerate(articles[:3]):
            news_content += f"Article {i+1}: {article.get('title', '')}\n{article.get('content', '')[:500]}\n\n"

        prompt = f"""Siz professional yangiliklar muharriri siz. Quyidagi maqolalar asosida Telegram kanal uchun post yozing.

Maqolalar:
{news_content}

Talablar:
- {emoji} emoji bilan chiroyli sarlavha bilan boshlang
- 3-4 paragraf yozing — yangilikni batafsil tushuntiring
- O'zbek va ingliz tilida aralash yozing (avval o'zbek, keyin ingliz)
- Post davomida emoji ishlating
- Oxirida: {hashtags}
- Uzunlik: 300-500 so'z
- Faqat O'zbekistonga oid yangiliklar haqida yozing

Postni hozir yozing:"""

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}]
        )

        post_text = response.choices[0].message.content

        image_url = None
        for article in articles:
            if article.get("image"):
                image_url = article["image"]
                break

        print(f"Sending to {CHANNEL_ID}...")

        if image_url:
            try:
                await bot.send_photo(
                    chat_id=CHANNEL_ID,
                    photo=image_url,
                    caption=post_text,
                    parse_mode="Markdown"
                )
                print("✅ Post with image sent!")
                return
            except Exception:
                pass

        await bot.send_message(
            chat_id=CHANNEL_ID,
            text=post_text,
            parse_mode="Markdown"
        )
        print("✅ Post sent!")

    except Exception as e:
        print(f"❌ Error: {str(e)}")
        try:
            await bot.send_message(
                chat_id=CHANNEL_ID,
                text=post_text
            )
            print("✅ Post sent (no markdown)!")
        except Exception as e2:
            print(f"❌ Failed: {str(e2)}")

def run_async(coro):
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(coro)
    loop.close()

def job():
    run_async(fetch_and_post_news())

def main():
    print("🤖 O'zbekiston News Bot started!")
    print(f"📢 Channel: {CHANNEL_ID}")
    print("⏰ Kuniga 4 marta post\n")

    print("Birinchi post yuborilmoqda...")
    job()

    schedule.every().day.at("09:00").do(job)
    schedule.every().day.at("13:00").do(job)
    schedule.every().day.at("17:00").do(job)
    schedule.every().day.at("21:00").do(job)

    while True:
        schedule.run_pending()
        time.sleep(60)

if __name__ == "__main__":
    main()