import os, io, json, random, string, requests, psycopg2, fitz, base64, asyncio
from datetime import datetime, timedelta, date
from groq import Groq
from dotenv import load_dotenv
from gtts import gTTS
from langdetect import detect
from telegram import Update, BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, LabeledPrice, WebAppInfo
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, CallbackQueryHandler, PreCheckoutQueryHandler, filters, ContextTypes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import Pt as DocPt, RGBColor as DocRGB

load_dotenv()

TELEGRAM_TOKEN    = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY      = os.getenv("GROQ_API_KEY")
DATABASE_URL      = os.getenv("DATABASE_URL")
CARD_NUMBER       = os.getenv("CARD_NUMBER")
WEATHER_API_KEY   = os.getenv("WEATHER_API_KEY")
COINGECKO_API_KEY = os.getenv("COINGECKO_API_KEY")
MINIAPP_URL       = os.getenv("MINIAPP_URL", "")

ADMIN_ID       = 8230883785
ADMIN_USERNAME = "temur_uzb7779"
BOT_NAME       = "Emerland AI"

STARS_STANDARD = 770
STARS_PREMIUM  = 1150
STARS_GROUP    = 8547
STARS_IMAGINE  = 50
USDT_STANDARD  = 10
USDT_PREMIUM   = 15
COINS_PER_MSG  = 5
COINS_STANDARD = 10000
COINS_PREMIUM  = 100000
AFFILIATE_PCT  = 25
MAX_WARNINGS   = 10

BAD_WORDS = [
    "ублюдок","сука","пизда","хуй","мудак","залупа","пиздец","блядь","ёбаный",
    "fuck","shit","asshole","bitch","bastard","dick","cunt",
    "ahmoq","tentak","yaramas","eshak","haromzoda","onangni","otangni","egangni"
]

SEARCH_KEYWORDS = [
    "today","now","current","latest","news","price","weather","rate",
    "bugun","hozir","narx","kurs","yangilik","ob-havo",
    "сегодня","сейчас","курс","цена","новости","погода"
]

GTTS_LANG_MAP = {
    "uz":"ru","en":"en","ru":"ru","tr":"tr","de":"de","fr":"fr",
    "es":"es","ar":"ar","ko":"ko","ja":"ja","it":"it","pt":"pt","hi":"hi"
}

LANGUAGES = {
    "en":("🇬🇧","English"), "ru":("🇷🇺","Русский"), "uz":("🇺🇿","O'zbek"),
    "tr":("🇹🇷","Türkçe"),  "de":("🇩🇪","Deutsch"), "fr":("🇫🇷","Français"),
    "es":("🇪🇸","Español"), "ar":("🇸🇦","العربية"), "ko":("🇰🇷","한국어"),
    "ja":("🇯🇵","日本語"),  "zh":("🇨🇳","中文"),     "it":("🇮🇹","Italiano"),
    "pt":("🇵🇹","Português"),"hi":("🇮🇳","हिंदी"),
}

ai = Groq(api_key=GROQ_API_KEY)
user_histories = {}

def db():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    conn = db(); c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY, user_id BIGINT UNIQUE NOT NULL,
        username TEXT, full_name TEXT, plan TEXT DEFAULT 'free',
        expires_at TIMESTAMP, is_blocked BOOLEAN DEFAULT FALSE,
        is_admin BOOLEAN DEFAULT FALSE, joined_at TIMESTAMP DEFAULT NOW(),
        language TEXT DEFAULT 'uz', language_changed_at TIMESTAMP,
        coins INTEGER DEFAULT 0, warning_count INTEGER DEFAULT 0,
        referral_code TEXT, affiliate_code TEXT, referred_by BIGINT,
        referral_count INTEGER DEFAULT 0, affiliate_earnings INTEGER DEFAULT 0,
        claimed_standard BOOLEAN DEFAULT FALSE, claimed_premium BOOLEAN DEFAULT FALSE,
        streak_count INTEGER DEFAULT 0, last_gift_claim DATE,
        streak_claimed BOOLEAN DEFAULT FALSE,
        total_messages INTEGER DEFAULT 0, achievement_100k BOOLEAN DEFAULT FALSE,
        usage_chat INTEGER DEFAULT 0, usage_image INTEGER DEFAULT 0,
        usage_pdf INTEGER DEFAULT 0, usage_cv INTEGER DEFAULT 0,
        usage_email INTEGER DEFAULT 0, usage_tts INTEGER DEFAULT 0,
        usage_post INTEGER DEFAULT 0, usage_biznes INTEGER DEFAULT 0,
        usage_translate INTEGER DEFAULT 0, usage_code INTEGER DEFAULT 0,
        last_reset DATE DEFAULT CURRENT_DATE
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS user_memory (
        user_id BIGINT PRIMARY KEY, name TEXT, facts TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS group_modes (
        chat_id BIGINT PRIMARY KEY, activated_by BIGINT,
        activated_at TIMESTAMP DEFAULT NOW()
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS promo_codes (
        id SERIAL PRIMARY KEY, code TEXT UNIQUE,
        reward_type TEXT, reward_value INTEGER,
        max_uses INTEGER DEFAULT 1, used_count INTEGER DEFAULT 0,
        expires_at TIMESTAMP, created_at TIMESTAMP DEFAULT NOW()
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS promo_uses (
        user_id BIGINT, code TEXT, PRIMARY KEY(user_id, code)
    )""")
    c.execute("UPDATE users SET is_admin=TRUE WHERE user_id=%s",(ADMIN_ID,))
    conn.commit(); conn.close()

def rnd(n=8):
    return ''.join(random.choices(string.ascii_uppercase+string.digits, k=n))

def ensure_user(uid, username=None, full_name=None, referred_by=None):
    try:
        conn=db(); c=conn.cursor()
        c.execute("""INSERT INTO users(user_id,username,full_name,referral_code,affiliate_code,referred_by,is_admin)
            VALUES(%s,%s,%s,%s,%s,%s,%s) ON CONFLICT(user_id) DO UPDATE SET
            username=EXCLUDED.username,full_name=EXCLUDED.full_name""",
            (uid,username,full_name,rnd(8),rnd(10),referred_by,uid==ADMIN_ID))
        conn.commit()
        if referred_by:
            c.execute("UPDATE users SET referral_count=referral_count+1 WHERE user_id=%s",(referred_by,))
            conn.commit()
        conn.close()
    except: pass

def get_user(uid):
    D={"plan":"free","expires_at":None,"is_blocked":False,"is_admin":False,
       "full_name":None,"username":None,"language":"uz","language_changed_at":None,
       "coins":0,"warning_count":0,"referral_code":None,"affiliate_code":None,
       "referred_by":None,"referral_count":0,"affiliate_earnings":0,
       "claimed_standard":False,"claimed_premium":False,"streak_count":0,
       "last_gift_claim":None,"streak_claimed":False,"total_messages":0,"achievement_100k":False}
    try:
        conn=db(); c=conn.cursor()
        c.execute("""SELECT plan,expires_at,is_blocked,is_admin,full_name,username,
            language,language_changed_at,coins,warning_count,referral_code,affiliate_code,
            referred_by,referral_count,affiliate_earnings,claimed_standard,claimed_premium,
            streak_count,last_gift_claim,streak_claimed,total_messages,achievement_100k
            FROM users WHERE user_id=%s""",(uid,))
        row=c.fetchone(); conn.close()
        if not row: return D
        keys=["plan","expires_at","is_blocked","is_admin","full_name","username",
              "language","language_changed_at","coins","warning_count","referral_code",
              "affiliate_code","referred_by","referral_count","affiliate_earnings",
              "claimed_standard","claimed_premium","streak_count","last_gift_claim",
              "streak_claimed","total_messages","achievement_100k"]
        u=dict(zip(keys,row))
        for k in ["is_blocked","is_admin","claimed_standard","claimed_premium","streak_claimed","achievement_100k"]:
            u[k]=bool(u.get(k))
        for k in ["coins","warning_count","referral_count","affiliate_earnings","streak_count","total_messages"]:
            u[k]=u.get(k) or 0
        u["language"]=u.get("language") or "uz"
        if not u["is_admin"] and u["expires_at"] and datetime.now()>u["expires_at"] and u["plan"]!="free":
            set_plan(uid,"free",None); u["plan"]="free"; u["expires_at"]=None
        return u
    except: return D

def set_plan(uid,plan,days=30):
    try:
        conn=db(); c=conn.cursor()
        exp=datetime.now()+timedelta(days=days) if days else None
        c.execute("UPDATE users SET plan=%s,expires_at=%s WHERE user_id=%s",(plan,exp,uid))
        conn.commit(); conn.close()
    except: pass

def set_blocked(uid,v):
    try:
        conn=db(); c=conn.cursor()
        c.execute("UPDATE users SET is_blocked=%s WHERE user_id=%s",(v,uid))
        conn.commit(); conn.close()
    except: pass

def set_language(uid,lang):
    try:
        conn=db(); c=conn.cursor()
        c.execute("UPDATE users SET language=%s,language_changed_at=%s WHERE user_id=%s",(lang,datetime.now(),uid))
        conn.commit(); conn.close()
    except: pass

def add_coins(uid,n):
    try:
        conn=db(); c=conn.cursor()
        c.execute("UPDATE users SET coins=coins+%s WHERE user_id=%s",(n,uid))
        conn.commit(); conn.close()
    except: pass

def add_warning(uid):
    try:
        conn=db(); c=conn.cursor()
        c.execute("UPDATE users SET warning_count=warning_count+1 WHERE user_id=%s",(uid,))
        c.execute("SELECT warning_count FROM users WHERE user_id=%s",(uid,))
        n=c.fetchone()[0]; conn.commit(); conn.close(); return n
    except: return 0

def inc_msg(uid,is_admin=False):
    if is_admin: return 0
    try:
        conn=db(); c=conn.cursor()
        c.execute("UPDATE users SET total_messages=total_messages+1,coins=coins+%s WHERE user_id=%s",(COINS_PER_MSG,uid))
        conn.commit()
        c.execute("SELECT total_messages,referred_by FROM users WHERE user_id=%s",(uid,))
        row=c.fetchone(); conn.close()
        if row:
            total,ref=row
            if ref:
                bonus=int(COINS_PER_MSG*AFFILIATE_PCT/100)
                add_coins(ref,bonus)
                try:
                    conn2=db(); c2=conn2.cursor()
                    c2.execute("UPDATE users SET affiliate_earnings=affiliate_earnings+%s WHERE user_id=%s",(bonus,ref))
                    conn2.commit(); conn2.close()
                except: pass
            return total
        return 0
    except: return 0

def check_limit(uid,feature,limit):
    if limit==-1: return True
    if limit==0: return False
    try:
        conn=db(); c=conn.cursor()
        today=datetime.now().date()
        c.execute("SELECT last_reset FROM users WHERE user_id=%s",(uid,))
        row=c.fetchone()
        if row and row[0]<today:
            c.execute("""UPDATE users SET usage_chat=0,usage_image=0,usage_pdf=0,
                usage_cv=0,usage_email=0,usage_tts=0,usage_post=0,usage_biznes=0,
                usage_translate=0,usage_code=0,last_reset=%s WHERE user_id=%s""",(today,uid))
            conn.commit()
        c.execute(f"SELECT usage_{feature} FROM users WHERE user_id=%s",(uid,))
        row=c.fetchone(); usage=row[0] if row else 0
        if usage>=limit: conn.close(); return False
        c.execute(f"UPDATE users SET usage_{feature}=usage_{feature}+1 WHERE user_id=%s",(uid,))
        conn.commit(); conn.close(); return True
    except: return True

def get_limits(plan,is_admin=False):
    ALL={k:-1 for k in ["chat","image","pdf","cv","email","voice","pptx","word","tts","imagine","translate","code","biznes","post","document"]}
    if is_admin or plan=="premium": return ALL
    if plan=="standard":
        return {"chat":30,"image":30,"pdf":30,"cv":30,"email":30,"voice":0,"pptx":0,"word":0,"tts":30,"imagine":0,"translate":30,"code":30,"biznes":30,"post":30,"document":30}
    return {"chat":20,"image":20,"pdf":0,"cv":0,"email":0,"voice":0,"pptx":0,"word":0,"tts":0,"imagine":0,"translate":20,"code":20,"biznes":20,"post":20,"document":0}

def get_memory(uid):
    try:
        conn=db(); c=conn.cursor()
        c.execute("SELECT name,facts FROM user_memory WHERE user_id=%s",(uid,))
        row=c.fetchone(); conn.close()
        return {"name":row[0],"facts":row[1]} if row else None
    except: return None

def save_memory(uid,name,facts):
    try:
        conn=db(); c=conn.cursor()
        c.execute("""INSERT INTO user_memory(user_id,name,facts) VALUES(%s,%s,%s)
            ON CONFLICT(user_id) DO UPDATE SET name=%s,facts=%s""",(uid,name,facts,name,facts))
        conn.commit(); conn.close()
    except: pass

def is_group_active(chat_id):
    try:
        conn=db(); c=conn.cursor()
        c.execute("SELECT 1 FROM group_modes WHERE chat_id=%s",(chat_id,))
        r=c.fetchone(); conn.close(); return r is not None
    except: return False

def activate_group(chat_id,uid):
    try:
        conn=db(); c=conn.cursor()
        c.execute("INSERT INTO group_modes(chat_id,activated_by) VALUES(%s,%s) ON CONFLICT DO NOTHING",(chat_id,uid))
        conn.commit(); conn.close()
    except: pass

def plan_info(u):
    if u["is_admin"]: return "👑","ADMIN"
    p=u["plan"]
    em={"free":"🆓","standard":"⭐","premium":"💎"}.get(p,"🆓")
    nm={"free":"FREE","standard":"STANDARD","premium":"PREMIUM"}.get(p,"FREE")
    return em,nm

def exp_str(u):
    if u["is_admin"]: return "∞"
    if u.get("expires_at"): return u["expires_at"].strftime("%d.%m.%Y")
    return "—"

def get_weather(city):
    try:
        r=requests.get(f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={WEATHER_API_KEY}&units=metric",timeout=10).json()
        if r.get("cod")!=200: return None
        w=r["weather"][0]; m=r["main"]; wind=r["wind"]
        em={"Clear":"☀️","Clouds":"☁️","Rain":"🌧️","Snow":"❄️","Thunderstorm":"⛈️","Drizzle":"🌦️","Mist":"🌫️"}.get(w["main"],"🌤")
        return (f"🌍 *{r['name']}, {r['sys']['country']}*\n━━━━━━━━━━━━━━━━━━━━\n"
                f"{em} *{w['description'].capitalize()}*\n\n"
                f"🌡 {m['temp']:.1f}°C | 💧 {m['humidity']}% | 💨 {wind['speed']}m/s\n"
                f"🔼 {m['temp_max']:.1f}° | 🔽 {m['temp_min']:.1f}°")
    except: return None

def get_crypto(coin):
    try:
        h={"x-cg-demo-api-key":COINGECKO_API_KEY} if COINGECKO_API_KEY else {}
        r=requests.get(f"https://api.coingecko.com/api/v3/coins/{coin.lower()}?localization=false&tickers=false&community_data=false&developer_data=false",headers=h,timeout=10).json()
        if "error" in r: return None
        md=r["market_data"]
        price=md["current_price"]["usd"]
        c24=md["price_change_percentage_24h"] or 0
        c7=md["price_change_percentage_7d"] or 0
        mc=md["market_cap"]["usd"]
        h24=md["high_24h"]["usd"]; l24=md["low_24h"]["usd"]
        arrow="📈" if c24>=0 else "📉"
        return (f"💰 *{r['name']} ({r['symbol'].upper()})*\n━━━━━━━━━━━━━━━━━━━━\n"
                f"💵 ${price:,.4f}\n{arrow} 24h: {c24:+.2f}% | 7d: {c7:+.2f}%\n"
                f"📈 ${h24:,.4f} | 📉 ${l24:,.4f}\n"
                f"🏦 ${mc:,.0f} | 🏆 #{r.get('market_cap_rank','N/A')}")
    except: return None

def gen_image(prompt):
    try:
        for seed in [random.randint(1,999999) for _ in range(3)]:
            url=f"https://image.pollinations.ai/prompt/{requests.utils.quote(prompt)}?seed={seed}&width=1024&height=1024&nologo=true&enhance=true"
            r=requests.get(url,timeout=90)
            if r.status_code==200 and len(r.content)>5000:
                return r.content
        return None
    except: return None

def detect_gtts(text):
    try: return GTTS_LANG_MAP.get(detect(text),"en")
    except: return "en"

def is_bad(text):
    return any(w in text.lower() for w in BAD_WORDS)

def make_pptx(title,slides):
    prs=Presentation()
    prs.slide_width=Inches(13.33); prs.slide_height=Inches(7.5)
    sl=prs.slides.add_slide(prs.slide_layouts[0])
    sl.background.fill.solid(); sl.background.fill.fore_color.rgb=RGBColor(0x0F,0x0F,0x1A)
    tb=sl.shapes.title; tb.text=title
    tb.text_frame.paragraphs[0].font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
    tb.text_frame.paragraphs[0].font.size=Pt(40); tb.text_frame.paragraphs[0].font.bold=True
    for s in slides:
        sl=prs.slides.add_slide(prs.slide_layouts[1])
        sl.background.fill.solid(); sl.background.fill.fore_color.rgb=RGBColor(0x0F,0x0F,0x1A)
        ti=sl.shapes.title; ti.text=s["title"]
        ti.text_frame.paragraphs[0].font.color.rgb=RGBColor(0x89,0xB4,0xFA)
        ti.text_frame.paragraphs[0].font.size=Pt(28); ti.text_frame.paragraphs[0].font.bold=True
        tf=sl.placeholders[1].text_frame; tf.clear()
        for i,pt in enumerate(s["points"]):
            p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
            p.text=f"▸ {pt}"; p.font.color.rgb=RGBColor(0xCD,0xD6,0xF4); p.font.size=Pt(18)
    prs.save("pres.pptx"); return "pres.pptx"

def make_docx(title,sections):
    doc=Document()
    h=doc.add_heading(title,0)
    if h.runs: h.runs[0].font.color.rgb=DocRGB(0x1E,0x3A,0x8A)
    for sec in sections:
        doc.add_heading(sec["title"],level=1)
        for pt in sec["points"]: doc.add_paragraph(pt,style="List Bullet")
        doc.add_paragraph()
    doc.save("doc.docx"); return "doc.docx"

async def ai_chat(uid,text,do_search=False,limits=None):
    mem=get_memory(uid)
    mc=f"User: name={mem['name']}, facts={mem['facts']}. " if mem and (mem.get("name") or mem.get("facts")) else ""
    sys_p=(
        "You are Emerland AI — a professional, intelligent AI assistant. "
        "CRITICAL: Always reply in the EXACT same language as the user's message. "
        "Uzbek→Uzbek, English→English, Russian→Russian. "
        "If asked who created you: 'I was created by @temur_uzb7779.' "
        "Never say you are ChatGPT, Claude or any other AI. You are Emerland AI. "
        "Be precise, helpful and professional. "+mc
    )
    if uid not in user_histories:
        user_histories[uid]=[{"role":"system","content":sys_p}]
    else:
        user_histories[uid][0]["content"]=sys_p
    if do_search and limits and limits.get("chat")!=-1:
        if any(k in text.lower() for k in SEARCH_KEYWORDS):
            text=f"{text}\n[Provide most accurate current knowledge]"
    user_histories[uid].append({"role":"user","content":text})
    if len(user_histories[uid])>21:
        user_histories[uid]=[user_histories[uid][0]]+user_histories[uid][-20:]
    r=ai.chat.completions.create(model="llama-3.3-70b-versatile",messages=user_histories[uid],max_tokens=2048)
    reply=r.choices[0].message.content
    user_histories[uid].append({"role":"assistant","content":reply})
    return reply

async def ai_once(prompt):
    r=ai.chat.completions.create(model="llama-3.3-70b-versatile",messages=[{"role":"user","content":prompt}],max_tokens=2048)
    return r.choices[0].message.content

async def update_mem(uid,user_text,reply):
    mem=get_memory(uid)
    cn=mem["name"] if mem else ""; cf=mem["facts"] if mem else ""
    try:
        r=await ai_once(f"Extract name/facts. Current: name={cn},facts={cf}\nUser:{user_text}\nJSON only:{{\"name\":\"...\",\"facts\":\"...\"}}")
        r=r.strip()
        if "```" in r: r=r.split("```")[1]; r=r[4:] if r.startswith("json") else r
        d=json.loads(r); save_memory(uid,d.get("name",cn),d.get("facts",cf))
    except: pass

async def broadcast_all(app,message,parse_mode="Markdown"):
    try:
        conn=db(); c=conn.cursor()
        c.execute("SELECT user_id,language FROM users WHERE is_blocked=FALSE")
        rows=c.fetchall(); conn.close()
        sent=failed=0
        for uid,lang in rows:
            try:
                await app.bot.send_message(chat_id=uid,text=message,parse_mode=parse_mode)
                sent+=1
                await asyncio.sleep(0.05)
            except: failed+=1
        return sent,failed
    except: return 0,0

async def start(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ref=None
    if ctx.args:
        try:
            code=ctx.args[0]; conn=db(); c=conn.cursor()
            c.execute("SELECT user_id FROM users WHERE referral_code=%s OR affiliate_code=%s",(code,code))
            row=c.fetchone(); conn.close()
            if row and row[0]!=user.id: ref=row[0]
        except: pass
    ensure_user(user.id,user.username,user.full_name,ref)
    u=get_user(user.id)
    pe,pn=plan_info(u)
    exp=exp_str(u)

    text=(
        f"✨ *{BOT_NAME}* ga xush kelibsiz!\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"🤖 *Men nimalар qila olaman:*\n\n"
        f"💬 *AI Suhbat* — Istalgan savolga javob\n"
        f"🎨 *AI Rasm* — So'z bilan rasm yaratish\n"
        f"🌐 *Tarjima* — 100+ til\n"
        f"💻 *Kod* — Dastur yozish va tushuntirish\n"
        f"📄 *PDF Tahlil* — Hujjatni o'qib tahlil\n"
        f"🎤 *Ovoz* — Audio xabarni matnga\n"
        f"📊 *PowerPoint* — Prezentatsiya yaratish\n"
        f"📝 *Word* — Hujjat yaratish\n"
        f"👤 *CV* — Professional rezyume\n"
        f"📧 *Email* — Biznes xat\n"
        f"🌦 *Ob-havo* — Istalgan shahar\n"
        f"💰 *Crypto* — Real vaqt narxlar\n"
        f"📰 *Yangiliklar* — Dunyo yangiliklari\n\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"📋 *Sizning tarifingiz:* {pe} {pn}\n"
        f"🪙 *Coinlar:* {u['coins']:,}\n"
        f"📅 *Tugash:* {exp}\n\n"
        f"👇 *Quyidagi tugmani bosing:*"
    )

    kb=[]
    if MINIAPP_URL:
        kb.append([InlineKeyboardButton("🚀 Emerland AI ni Ochish",web_app=WebAppInfo(url=MINIAPP_URL))])
    kb.append([
        InlineKeyboardButton("💰 Tariflar",callback_data="show_plans"),
        InlineKeyboardButton("🎁 Bonus",callback_data="show_gift")
    ])
    kb.append([
        InlineKeyboardButton("👥 Referal",callback_data="show_referral"),
        InlineKeyboardButton("📊 Statistika",callback_data="show_stats")
    ])

    await update.message.reply_text(text,reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")

async def callback(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    q=update.callback_query; await q.answer()
    data=q.data; uid=q.from_user.id
    u=get_user(uid)

    if data=="show_plans":
        pe,pn=plan_info(u); exp=exp_str(u)
        kb=[
            [InlineKeyboardButton(f"⭐ Standart — {STARS_STANDARD}⭐",callback_data="stars_std"),
             InlineKeyboardButton(f"💎 Premium — {STARS_PREMIUM}⭐",callback_data="stars_prm")],
            [InlineKeyboardButton(f"💳 Standart — {USDT_STANDARD}$",callback_data="card_std"),
             InlineKeyboardButton(f"💳 Premium — {USDT_PREMIUM}$",callback_data="card_prm")],
            [InlineKeyboardButton(f"🪙 Standart — {COINS_STANDARD:,}🪙",callback_data="coins_std"),
             InlineKeyboardButton(f"🪙 Premium — {COINS_PREMIUM:,}🪙",callback_data="coins_prm")],
            [InlineKeyboardButton(f"👥 Guruh — {STARS_GROUP}⭐",callback_data="stars_grp")],
            [InlineKeyboardButton("🔙 Orqaga",callback_data="back_start")]
        ]
        await q.message.edit_text(
            f"💰 *Tariflar*\n━━━━━━━━━━━━━━━━━━━━\n\n"
            f"🆓 *BEPUL* — Suhbat 20/kun • Tarjima • Kod\n\n"
            f"⭐ *STANDART — {USDT_STANDARD}$ / {STARS_STANDARD}⭐ / {COINS_STANDARD:,}🪙*\n"
            f"• 30/kun + PDF • CV • Email • AI Ovoz • Hujjat\n\n"
            f"💎 *PREMIUM — {USDT_PREMIUM}$ / {STARS_PREMIUM}⭐ / {COINS_PREMIUM:,}🪙*\n"
            f"• Cheksiz + AI Rasm • PowerPoint • Word\n\n"
            f"👥 *GURUH — {STARS_GROUP}⭐*\n• Butun guruh uchun\n\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"{pe} *{pn}* | 📅 {exp} | 🪙 {u['coins']:,}",
            reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")

    elif data=="show_gift":
        today=date.today()
        if u["last_gift_claim"] and u["last_gift_claim"]>=today:
            await q.answer("🎁 Bugun bonus allaqachon olindi! Ertaga keling.",show_alert=True)
            return
        streak=u["streak_count"]
        streak=streak+1 if(u["last_gift_claim"] and(today-u["last_gift_claim"]).days==1) else 1
        try:
            conn=db(); c=conn.cursor()
            c.execute("UPDATE users SET streak_count=%s,last_gift_claim=%s,usage_chat=usage_chat+5,coins=coins+25 WHERE user_id=%s",(streak,today,uid))
            conn.commit(); conn.close()
        except: pass
        if streak>=10 and not u["streak_claimed"]:
            set_plan(uid,"standard",10)
            try:
                conn=db(); c=conn.cursor()
                c.execute("UPDATE users SET streak_claimed=TRUE WHERE user_id=%s",(uid,))
                conn.commit(); conn.close()
            except: pass
            await q.answer("🎉 10 kun streak! ⭐ Standart 10 kunga yoqildi!",show_alert=True)
        else:
            await q.answer(f"🎁 Bonus olindi! Streak: {streak} kun 🔥 +25 coin",show_alert=True)

    elif data=="show_referral":
        bot=await ctx.bot.get_me()
        link=f"https://t.me/{bot.username}?start={u['referral_code']}"
        count=u["referral_count"]
        ss="✅" if u["claimed_standard"] else (f"🔓 Tayyor!" if count>=10 else f"{count}/10")
        ps="✅" if u["claimed_premium"] else (f"🔓 Tayyor!" if count>=30 else f"{count}/30")
        kb=[]
        if count>=10 and not u["claimed_standard"]:
            kb.append([InlineKeyboardButton("🎁 Standart 15 kun olish",callback_data="claim_std")])
        if count>=30 and not u["claimed_premium"]:
            kb.append([InlineKeyboardButton("🎁 Premium 15 kun olish",callback_data="claim_prm")])
        kb.append([InlineKeyboardButton("🔙 Orqaga",callback_data="back_start")])
        await q.message.edit_text(
            f"👥 *Referal Dasturi*\n━━━━━━━━━━━━━━━━━━━━\n\n"
            f"🔗 *Sizning havolangiz:*\n`{link}`\n\n"
            f"📊 Taklif qilingan: *{count}* do'st\n\n"
            f"🎁 10 do'st → ⭐ Standart 15 kun — {ss}\n"
            f"🎁 30 do'st → 💎 Premium 15 kun — {ps}\n\n"
            f"🪙 Har referaldan *{AFFILIATE_PCT}%* coin olasiz!",
            reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")

    elif data=="show_stats":
        try:
            conn=db(); c=conn.cursor()
            c.execute("SELECT usage_chat,usage_image,usage_pdf,usage_cv,usage_email,usage_tts,usage_post,usage_biznes FROM users WHERE user_id=%s",(uid,))
            row=c.fetchone(); conn.close()
            chat,image,pdf,cv,email,tts,post,biznes=row if row else (0,)*8
        except: chat=image=pdf=cv=email=tts=post=biznes=0
        pe,pn=plan_info(u); exp=exp_str(u)
        await q.message.edit_text(
            f"📊 *Statistika*\n━━━━━━━━━━━━━━━━━━━━\n\n"
            f"{pe} *{pn}* | 📅 {exp}\n"
            f"🪙 {u['coins']:,} | 👥 {u['referral_count']} | 🔥 {u['streak_count']}k | 💬 {u['total_messages']:,}\n\n"
            f"📈 *Bugun:*\n"
            f"💬{chat} 🖼️{image} 📄{pdf} 👤{cv} 📧{email} 🔊{tts} 📱{post} 💼{biznes}",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Orqaga",callback_data="back_start")]]),
            parse_mode="Markdown")

    elif data=="back_start":
        pe,pn=plan_info(u); exp=exp_str(u)
        text=(
            f"✨ *{BOT_NAME}* ga xush kelibsiz!\n"
            f"━━━━━━━━━━━━━━━━━━━━\n\n"
            f"🤖 *Men nimalar qila olaman:*\n\n"
            f"💬 AI Suhbat • 🎨 AI Rasm • 🌐 Tarjima\n"
            f"💻 Kod • 📄 PDF • 🎤 Ovoz • 📊 PowerPoint\n"
            f"📝 Word • 👤 CV • 📧 Email\n"
            f"🌦 Ob-havo • 💰 Crypto • 📰 Yangiliklar\n\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"{pe} *{pn}* | 🪙 {u['coins']:,} | 📅 {exp}\n\n"
            f"👇 *Quyidagi tugmani bosing:*"
        )
        kb=[]
        if MINIAPP_URL:
            kb.append([InlineKeyboardButton("🚀 Emerland AI ni Ochish",web_app=WebAppInfo(url=MINIAPP_URL))])
        kb.append([InlineKeyboardButton("💰 Tariflar",callback_data="show_plans"),InlineKeyboardButton("🎁 Bonus",callback_data="show_gift")])
        kb.append([InlineKeyboardButton("👥 Referal",callback_data="show_referral"),InlineKeyboardButton("📊 Statistika",callback_data="show_stats")])
        await q.message.edit_text(text,reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")

    elif data=="stars_std":
        await ctx.bot.send_invoice(chat_id=uid,title="⭐ Standart",description="30 kun",payload="std_stars",currency="XTR",prices=[LabeledPrice("Standart",STARS_STANDARD)])
    elif data=="stars_prm":
        await ctx.bot.send_invoice(chat_id=uid,title="💎 Premium",description="30 kun",payload="prm_stars",currency="XTR",prices=[LabeledPrice("Premium",STARS_PREMIUM)])
    elif data=="stars_grp":
        if uid==ADMIN_ID:
            await q.answer("👑 Admin uchun bepul! Guruhga qo'shing va /activate_group yuboring.",show_alert=True)
        else:
            await ctx.bot.send_invoice(chat_id=uid,title="👥 Guruh",description="Guruh uchun",payload="grp_stars",currency="XTR",prices=[LabeledPrice("Guruh",STARS_GROUP)])
    elif data.startswith("img_pay_"):
        prompt=data[8:]
        await ctx.bot.send_invoice(chat_id=uid,title="🎨 AI Rasm",description=f"{prompt}",payload=f"img_{prompt}",currency="XTR",prices=[LabeledPrice("Rasm",STARS_IMAGINE)])
    elif data=="card_std":
        if uid==ADMIN_ID:
            set_plan(uid,"standard",36500)
            await q.answer("👑 Admin: Standart yoqildi!",show_alert=True)
        else:
            kb=[[InlineKeyboardButton("💬 Admin bilan bog'lanish",url=f"https://t.me/{ADMIN_USERNAME}")],[InlineKeyboardButton("🔙 Orqaga",callback_data="show_plans")]]
            await q.message.edit_text(f"💳 *Karta orqali to'lov*\n━━━━━━━━━━━━━━━━━━━━\n\n`{CARD_NUMBER}`\n\n💵 *{USDT_STANDARD} USDT*\n\n1. To'lang\n2. Skrinshot oling\n3. Adminga yuboring\n4. 1 soatda yoqiladi ✅",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")
    elif data=="card_prm":
        if uid==ADMIN_ID:
            set_plan(uid,"premium",36500)
            await q.answer("👑 Admin: Premium yoqildi!",show_alert=True)
        else:
            kb=[[InlineKeyboardButton("💬 Admin bilan bog'lanish",url=f"https://t.me/{ADMIN_USERNAME}")],[InlineKeyboardButton("🔙 Orqaga",callback_data="show_plans")]]
            await q.message.edit_text(f"💳 *Karta orqali to'lov*\n━━━━━━━━━━━━━━━━━━━━\n\n`{CARD_NUMBER}`\n\n💵 *{USDT_PREMIUM} USDT*\n\n1. To'lang\n2. Skrinshot oling\n3. Adminga yuboring\n4. 1 soatda yoqiladi ✅",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")
    elif data=="coins_std":
        if u["coins"]>=COINS_STANDARD:
            conn=db(); c=conn.cursor()
            c.execute("UPDATE users SET coins=coins-%s WHERE user_id=%s",(COINS_STANDARD,uid))
            conn.commit(); conn.close(); set_plan(uid,"standard",30)
            await q.answer("✅ Standart 30 kunga yoqildi!",show_alert=True)
        else:
            await q.answer(f"❌ {COINS_STANDARD-u['coins']:,} coin yetarli emas!",show_alert=True)
    elif data=="coins_prm":
        if u["coins"]>=COINS_PREMIUM:
            conn=db(); c=conn.cursor()
            c.execute("UPDATE users SET coins=coins-%s WHERE user_id=%s",(COINS_PREMIUM,uid))
            conn.commit(); conn.close(); set_plan(uid,"premium",30)
            await q.answer("✅ Premium 30 kunga yoqildi!",show_alert=True)
        else:
            await q.answer(f"❌ {COINS_PREMIUM-u['coins']:,} coin yetarli emas!",show_alert=True)
    elif data=="claim_std":
        if u["referral_count"]>=10 and not u["claimed_standard"]:
            set_plan(uid,"standard",15)
            conn=db(); c=conn.cursor()
            c.execute("UPDATE users SET claimed_standard=TRUE WHERE user_id=%s",(uid,))
            conn.commit(); conn.close()
            await q.answer("🎉 Standart 15 kunga yoqildi!",show_alert=True)
        else:
            await q.answer("❌ Allaqachon olingan yoki referal yetarli emas!",show_alert=True)
    elif data=="claim_prm":
        if u["referral_count"]>=30 and not u["claimed_premium"]:
            set_plan(uid,"premium",15)
            conn=db(); c=conn.cursor()
            c.execute("UPDATE users SET claimed_premium=TRUE WHERE user_id=%s",(uid,))
            conn.commit(); conn.close()
            await q.answer("🎉 Premium 15 kunga yoqildi!",show_alert=True)
        else:
            await q.answer("❌ Allaqachon olingan yoki referal yetarli emas!",show_alert=True)
    elif data.startswith("ap_free_") or data.startswith("ap_std_") or data.startswith("ap_prm_"):
        parts=data.split("_"); pm={"free":"free","std":"standard","prm":"premium"}
        plan_key=parts[1]; tid=int(parts[2])
        plan=pm[plan_key]; days=30 if plan!="free" else None
        set_plan(tid,plan,days)
        pem={"free":"🆓","standard":"⭐","premium":"💎"}
        await q.message.edit_text(f"✅ `{tid}` → {pem[plan]} *{plan.upper()}*",parse_mode="Markdown")
        try: await ctx.bot.send_message(tid,f"{pem[plan]} Tarifingiz *{plan.upper()}*ga yangilandi!",parse_mode="Markdown")
        except: pass
    elif data.startswith("ap_block_"):
        tid=int(data.split("_")[2]); set_blocked(tid,True)
        await q.message.edit_text(f"🚫 `{tid}` bloklandi!",parse_mode="Markdown")
    elif data.startswith("ap_unblock_"):
        tid=int(data.split("_")[2]); set_blocked(tid,False)
        await q.message.edit_text(f"✅ `{tid}` blokdan chiqarildi!",parse_mode="Markdown")
    elif data.startswith("lang_"):
        nl=data[5:]
        lca=u["language_changed_at"]
        if lca and datetime.now()-lca<timedelta(hours=24):
            await q.answer("⏳ 24 soatda 1 marta o'zgartirish mumkin.",show_alert=True); return
        set_language(uid,nl)
        flag,name=LANGUAGES.get(nl,("🌐","Unknown"))
        await q.message.edit_text(f"✅ Til *{flag} {name}*ga o'zgartirildi!",parse_mode="Markdown")

async def handle_text(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    chat=update.effective_chat
    text=update.message.text or ""
    if chat.type in["group","supergroup"]:
        if not text.lower().startswith("bot "): return
        if not is_group_active(chat.id): return
        question=text[4:].strip()
        if not question: return
        try:
            await ctx.bot.send_chat_action(chat_id=chat.id,action="typing")
            reply=await ai_chat(user.id,question)
            await update.message.reply_text(reply)
        except Exception as e: await update.message.reply_text(f"Xato: {e}")
        return
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]:
        await update.message.reply_text(f"🚫 Bloklangansiz. @{ADMIN_USERNAME} bilan bog'laning."); return
    if is_bad(text):
        wc=add_warning(user.id)
        if wc>=MAX_WARNINGS:
            set_blocked(user.id,True)
            await update.message.reply_text("🚫 Haqorat uchun bloklandi.")
        else:
            await update.message.reply_text(f"⚠️ *Ogohlantirish {wc}/{MAX_WARNINGS}:* Hurmat bilan muomala qiling!",parse_mode="Markdown")
        return
    limits=get_limits(u["plan"],u["is_admin"])
    if not u["is_admin"] and not check_limit(user.id,"chat",limits["chat"]):
        kb=[[InlineKeyboardButton("💰 Tarifni yangilash",callback_data="show_plans")]]
        await update.message.reply_text("❌ *Kunlik limit tugadi!*\nTarifni yangilang.",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    total=inc_msg(user.id,u["is_admin"])
    if total==100000 and not u["achievement_100k"] and not u["is_admin"]:
        set_plan(user.id,"standard",5)
        try:
            conn=db(); c=conn.cursor()
            c.execute("UPDATE users SET achievement_100k=TRUE WHERE user_id=%s",(user.id,))
            conn.commit(); conn.close()
        except: pass
        await update.message.reply_text("🏆 *100,000 xabar! ⭐ Standart 5 kunga yoqildi!*",parse_mode="Markdown")
    await ctx.bot.send_chat_action(chat_id=chat.id,action="typing")
    try:
        reply=await ai_chat(user.id,text,do_search=True,limits=limits)
        await update.message.reply_text(reply)
        await update_mem(user.id,text,reply)
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_voice(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if limits["voice"]==0:
        kb=[[InlineKeyboardButton("💎 Premiumga o'tish",callback_data="show_plans")]]
        await update.message.reply_text("💎 *Ovoz xabarlari Premium talab qiladi!*",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    await ctx.bot.send_chat_action(chat_id=update.effective_chat.id,action="typing")
    try:
        f=await ctx.bot.get_file(update.message.voice.file_id)
        data=requests.get(f.file_path).content
        with open("v.ogg","wb") as fp: fp.write(data)
        with open("v.ogg","rb") as fp:
            tr=ai.audio.transcriptions.create(file=("v.ogg",fp.read()),model="whisper-large-v3")
        text=tr.text
        await update.message.reply_text(f"🎤 *Siz dedingiz:*\n{text}",parse_mode="Markdown")
        reply=await ai_chat(user.id,text)
        await update.message.reply_text(reply)
        try: os.remove("v.ogg")
        except: pass
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_photo(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if not u["is_admin"] and not check_limit(user.id,"image",limits["image"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    await ctx.bot.send_chat_action(chat_id=update.effective_chat.id,action="typing")
    try:
        photo=update.message.photo[-1]
        f=await ctx.bot.get_file(photo.file_id)
        # To'g'ridan URL orqali - base64 shart emas Groq uchun
        img_url=f.file_path
        caption=update.message.caption or "Bu rasmda nima bor? Batafsil tasvirlab ber."
        r=ai.chat.completions.create(
            model="meta-llama/llama-4-scout-17b-16e-instruct",
            messages=[{"role":"user","content":[
                {"type":"image_url","image_url":{"url":img_url}},
                {"type":"text","text":caption}
            ]}])
        await update.message.reply_text(r.choices[0].message.content)
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_document(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if limits["pdf"]==0:
        kb=[[InlineKeyboardButton("⭐ Standartga o'tish",callback_data="show_plans")]]
        await update.message.reply_text("⭐ *PDF Standart yoki Premium talab qiladi!*",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    if not u["is_admin"] and not check_limit(user.id,"pdf",limits["pdf"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    doc=update.message.document
    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("📄 Faqat PDF fayl yuboring!"); return
    await update.message.reply_text("⏳ PDF o'qilmoqda...")
    await ctx.bot.send_chat_action(chat_id=update.effective_chat.id,action="typing")
    try:
        f=await ctx.bot.get_file(doc.file_id)
        data=requests.get(f.file_path).content
        with open("tmp.pdf","wb") as fp: fp.write(data)
        pdf=fitz.open("tmp.pdf")
        text="".join(p.get_text() for p in pdf); pdf.close()
        try: os.remove("tmp.pdf")
        except: pass
        if len(text)>12000: text=text[:12000]+"..."
        caption=update.message.caption or "Bu hujjatni batafsil tahlil qilib, asosiy fikrlarini tushuntir."
        reply=await ai_chat(user.id,f"PDF:\n\n{text}\n\nSo'rov: {caption}",do_search=False)
        await update.message.reply_text(reply)
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_tts(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if limits["tts"]==0:
        kb=[[InlineKeyboardButton("⭐ Standartga o'tish",callback_data="show_plans")]]
        await update.message.reply_text("⭐ *AI Ovoz Standart talab qiladi!*\nMisol: `/ai_sound Salom dunyo`",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    if not check_limit(user.id,"tts",limits["tts"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    text=" ".join(ctx.args)
    if not text: await update.message.reply_text("Misol: `/ai_sound Salom dunyo`",parse_mode="Markdown"); return
    await update.message.reply_text("⏳ Ovoz yaratilmoqda...")
    try:
        tts_lang=detect_gtts(text)
        tts=gTTS(text=text,lang=tts_lang)
        buf=io.BytesIO(); tts.write_to_fp(buf); buf.seek(0)
        await update.message.reply_voice(voice=buf)
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_imagine(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    prompt=" ".join(ctx.args)
    if not prompt: await update.message.reply_text("Misol: `/imagine chiroyli tog' manzarasi`",parse_mode="Markdown"); return
    if limits["imagine"]==0:
        kb=[[InlineKeyboardButton("💎 Premiumga o'tish",callback_data="show_plans")],[InlineKeyboardButton(f"🎨 {STARS_IMAGINE}⭐ to'lab rasm",callback_data=f"img_pay_{prompt[:50]}")]]
        await update.message.reply_text("💎 *AI Rasm Premium talab qiladi!*\nYoki {STARS_IMAGINE}⭐ to'lang.",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    if not check_limit(user.id,"image",limits["image"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    await update.message.reply_text("⏳ AI rasm yaratilmoqda...")
    img=gen_image(prompt)
    if img:
        buf=io.BytesIO(img); buf.name="image.png"
        await update.message.reply_photo(photo=buf,caption=f"🎨 {prompt}")
    else:
        await update.message.reply_text("❌ Rasm yaratib bo'lmadi. Boshqa prompt sinab ko'ring.")

async def handle_pptx(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if limits["pptx"]==0:
        kb=[[InlineKeyboardButton("💎 Premiumga o'tish",callback_data="show_plans")]]
        await update.message.reply_text("💎 *PowerPoint Premium talab qiladi!*",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    topic=" ".join(ctx.args)
    if not topic: await update.message.reply_text("Misol: `/pptx sun'iy intellekt`",parse_mode="Markdown"); return
    await update.message.reply_text(f"⏳ *{topic}* prezentatsiyasi yaratilmoqda...",parse_mode="Markdown")
    try:
        raw=await ai_once(f"Create presentation for '{topic}'. JSON only:\n{{\"title\":\"...\",\"slides\":[{{\"title\":\"...\",\"points\":[\"...\",\"...\",\"...\"]}}]}}\nMin 6 slides. Same language as topic.")
        raw=raw.strip()
        if "```" in raw: raw=raw.split("```")[1]; raw=raw[4:] if raw.startswith("json") else raw
        data=json.loads(raw)
        path=make_pptx(data["title"],data["slides"])
        with open(path,"rb") as f:
            await update.message.reply_document(f,filename=f"{topic}.pptx",caption="✅ Tayyor!")
        os.remove(path)
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_word(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if limits["word"]==0:
        kb=[[InlineKeyboardButton("💎 Premiumga o'tish",callback_data="show_plans")]]
        await update.message.reply_text("💎 *Word Premium talab qiladi!*",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    topic=" ".join(ctx.args)
    if not topic: await update.message.reply_text("Misol: `/word biznes taklif`",parse_mode="Markdown"); return
    await update.message.reply_text(f"⏳ *{topic}* hujjati yaratilmoqda...",parse_mode="Markdown")
    try:
        raw=await ai_once(f"Create Word document for '{topic}'. JSON only:\n{{\"title\":\"...\",\"sections\":[{{\"title\":\"...\",\"points\":[\"...\",\"...\",\"...\"]}}]}}\nMin 5 sections. Same language.")
        raw=raw.strip()
        if "```" in raw: raw=raw.split("```")[1]; raw=raw[4:] if raw.startswith("json") else raw
        data=json.loads(raw)
        path=make_docx(data["title"],data["sections"])
        with open(path,"rb") as f:
            await update.message.reply_document(f,filename=f"{topic}.docx",caption="✅ Tayyor!")
        os.remove(path)
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_cv(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if limits["cv"]==0:
        kb=[[InlineKeyboardButton("⭐ Standartga o'tish",callback_data="show_plans")]]
        await update.message.reply_text("⭐ *CV Standart talab qiladi!*",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    if not check_limit(user.id,"cv",limits["cv"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    info=" ".join(ctx.args)
    if not info: await update.message.reply_text("Misol: `/cv Python dasturchi, 3 yil tajriba`",parse_mode="Markdown"); return
    await update.message.reply_text("⏳ CV yozilmoqda...")
    try:
        reply=await ai_once(f"Write professional ATS-friendly CV for: {info}\nSections: Summary, Experience, Skills, Education, Achievements.")
        await update.message.reply_text(f"👤 *CV:*\n\n{reply}",parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_email(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if limits["email"]==0:
        kb=[[InlineKeyboardButton("⭐ Standartga o'tish",callback_data="show_plans")]]
        await update.message.reply_text("⭐ *Email Standart talab qiladi!*",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown"); return
    if not check_limit(user.id,"email",limits["email"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    topic=" ".join(ctx.args)
    if not topic: await update.message.reply_text("Misol: `/email intervyudan keyin minnatdorchilik`",parse_mode="Markdown"); return
    await update.message.reply_text("⏳ Email yozilmoqda...")
    try:
        reply=await ai_once(f"Write professional email about: {topic}\nSubject, greeting, body, closing. Same language.")
        await update.message.reply_text(f"📧 *Email:*\n\n{reply}",parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_post(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if not check_limit(user.id,"post",limits["post"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    topic=" ".join(ctx.args)
    if not topic: await update.message.reply_text("Misol: `/post do'kon ochilishi`",parse_mode="Markdown"); return
    await update.message.reply_text("⏳ Post yozilmoqda...")
    try:
        reply=await ai_once(f"Write viral social media post about: {topic}\nHook, content, CTA, hashtags. Same language.")
        await update.message.reply_text(f"📱 *Post:*\n\n{reply}",parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_biznes(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if u["is_blocked"]: return
    limits=get_limits(u["plan"],u["is_admin"])
    if not check_limit(user.id,"biznes",limits["biznes"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    idea=" ".join(ctx.args)
    if not idea: await update.message.reply_text("Misol: `/biznes online do'kon`",parse_mode="Markdown"); return
    await update.message.reply_text("⏳ Biznes reja yozilmoqda...")
    try:
        reply=await ai_once(f"Write comprehensive business plan for: {idea}\nExecutive Summary, Market, Products, Marketing, Finance.")
        await update.message.reply_text(f"💼 *Biznes Reja:*\n\n{reply}",parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_translate(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    limits=get_limits(u["plan"],u["is_admin"])
    if not check_limit(user.id,"translate",limits["translate"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    text=" ".join(ctx.args)
    if not text: await update.message.reply_text("Misol: `/translate Hello → O'zbek`",parse_mode="Markdown"); return
    await update.message.reply_text("⏳ Tarjima qilinmoqda...")
    try:
        reply=await ai_once(f"Translate accurately: {text}")
        await update.message.reply_text(f"🌐 {reply}")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_code(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    limits=get_limits(u["plan"],u["is_admin"])
    if not check_limit(user.id,"code",limits["code"]):
        await update.message.reply_text("❌ Kunlik limit tugadi!"); return
    text=" ".join(ctx.args)
    if not text: await update.message.reply_text("Misol: `/code Python fibonacci funksiyasi`",parse_mode="Markdown"); return
    await update.message.reply_text("⏳ Kod yozilmoqda...")
    try:
        reply=await ai_once(f"Write clean, well-commented code for: {text}\nExplain how it works.")
        await update.message.reply_text(f"💻 {reply}")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_weather(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    city=" ".join(ctx.args)
    if not city: await update.message.reply_text("Misol: `/weather Toshkent`",parse_mode="Markdown"); return
    result=get_weather(city)
    await update.message.reply_text(result if result else "❌ Shahar topilmadi.",parse_mode="Markdown")

async def handle_crypto(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    coin=ctx.args[0] if ctx.args else ""
    if not coin: await update.message.reply_text("Misol: `/crypto bitcoin`",parse_mode="Markdown"); return
    result=get_crypto(coin)
    await update.message.reply_text(result if result else "❌ Topilmadi.",parse_mode="Markdown")

async def handle_news(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    await update.message.reply_text("⏳ Yangiliklar yuklanmoqda...")
    result=await ai_once("Give 5 important world news headlines with brief descriptions. Format: 📌 **Title**\nDescription\n")
    await update.message.reply_text(f"📰 *Dunyo Yangiliklari*\n━━━━━━━━━━━━━━━━━━━━\n\n{result}",parse_mode="Markdown")

async def handle_reset(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    user_histories[user.id]=[]
    await update.message.reply_text("✅ Suhbat tarixi tozalandi!")

async def handle_language(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    lca=u["language_changed_at"]
    if lca and datetime.now()-lca<timedelta(hours=24):
        await update.message.reply_text("⏳ 24 soatda 1 marta o'zgartirish mumkin."); return
    kb=[]; row=[]
    for code,(flag,name) in LANGUAGES.items():
        row.append(InlineKeyboardButton(f"{flag} {name}",callback_data=f"lang_{code}"))
        if len(row)==2: kb.append(row); row=[]
    if row: kb.append(row)
    await update.message.reply_text("🌐 *Tilni tanlang:*",reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")

async def handle_activate_group(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    chat=update.effective_chat
    if chat.type=="private": await update.message.reply_text("⚠️ Bu buyruq faqat guruhlarda ishlaydi!"); return
    user=update.effective_user
    activate_group(chat.id,user.id)
    await update.message.reply_text("✅ *Bot bu guruh uchun yoqildi!*\nA'zolar `bot [savol]` deb yoza oladi.",parse_mode="Markdown")

async def handle_promo(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    ensure_user(user.id,user.username,user.full_name)
    u=get_user(user.id)
    if not ctx.args: await update.message.reply_text("Kodni kiriting: `/promo KOD`",parse_mode="Markdown"); return
    code=ctx.args[0].upper()
    try:
        conn=db(); c=conn.cursor()
        c.execute("SELECT reward_type,reward_value,max_uses,used_count,expires_at FROM promo_codes WHERE code=%s",(code,))
        promo=c.fetchone()
        if not promo: await update.message.reply_text("❌ Noto'g'ri promo kod!"); conn.close(); return
        rtype,rval,maxu,used,exp=promo
        if(exp and datetime.now()>exp) or used>=maxu:
            await update.message.reply_text("❌ Muddati o'tgan yoki tugagan!"); conn.close(); return
        c.execute("SELECT 1 FROM promo_uses WHERE user_id=%s AND code=%s",(user.id,code))
        if c.fetchone(): await update.message.reply_text("❌ Allaqachon ishlatilgan!"); conn.close(); return
        c.execute("INSERT INTO promo_uses(user_id,code) VALUES(%s,%s)",(user.id,code))
        c.execute("UPDATE promo_codes SET used_count=used_count+1 WHERE code=%s",(code,))
        conn.commit(); conn.close()
        if rtype=="standard": set_plan(user.id,"standard",rval); reward=f"⭐ Standart {rval} kun!"
        elif rtype=="premium": set_plan(user.id,"premium",rval); reward=f"💎 Premium {rval} kun!"
        elif rtype=="coins": add_coins(user.id,rval); reward=f"🪙 +{rval:,} coin!"
        else: reward="✅"
        await update.message.reply_text(f"✅ *Promo kod qo'llandi!*\n🎁 {reward}",parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

# ADMIN
async def handle_admin(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id!=ADMIN_ID: await update.message.reply_text("❌ Ruxsat yo'q!"); return
    try:
        conn=db(); c=conn.cursor()
        c.execute("SELECT COUNT(*) FROM users"); total=c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM users WHERE plan='standard'"); std=c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM users WHERE plan='premium'"); prm=c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM users WHERE is_blocked=TRUE"); blk=c.fetchone()[0]
        c.execute("SELECT COALESCE(SUM(coins),0) FROM users"); tc=c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM group_modes"); grps=c.fetchone()[0]
        conn.close(); rev=std*USDT_STANDARD+prm*USDT_PREMIUM
    except: total=std=prm=blk=tc=grps=rev=0
    await update.message.reply_text(
        f"🔧 *Admin Panel*\n━━━━━━━━━━━━━━━━━━━━\n\n"
        f"👥 Jami: *{total}* | 🚫 Bloklangan: *{blk}*\n"
        f"🆓 Bepul: *{total-std-prm}* | ⭐ Standart: *{std}* | 💎 Premium: *{prm}*\n"
        f"👥 Guruhlar: *{grps}* | 💰 Daromad: ~*${rev}*\n"
        f"🪙 Jami coin: *{tc:,}*\n\n"
        f"`/users` — Foydalanuvchilar\n"
        f"`/find [@username yoki id]` — Topish\n"
        f"`/broadcast [xabar]` — Hammaga yuborish\n"
        f"`/createpromo [type] [qiymat] [son] [kod?]`\n"
        f"`/addcoins [uid] [miqdor]`\n"
        f"`/maintenance` — Texnik xizmat xabari",parse_mode="Markdown")

async def handle_users(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id!=ADMIN_ID: await update.message.reply_text("❌ Ruxsat yo'q!"); return
    try:
        conn=db(); c=conn.cursor()
        c.execute("SELECT user_id,username,full_name,plan,is_blocked FROM users ORDER BY CASE plan WHEN 'premium' THEN 1 WHEN 'standard' THEN 2 ELSE 3 END,joined_at DESC")
        rows=c.fetchall(); conn.close()
        if not rows: await update.message.reply_text("Foydalanuvchilar yo'q."); return
        pem={"free":"🆓","standard":"⭐","premium":"💎"}
        text=f"👥 *Jami foydalanuvchilar: {len(rows)}*\n━━━━━━━━━━━━━━━━━━━━\n\n"
        for uid,un,fn,plan,blk in rows:
            n=f"@{un}" if un else (fn or "Anonim")
            b=" 🚫" if blk else ""
            text+=f"{pem.get(plan,'🆓')} {n} `{uid}`{b}\n"
        if len(text)>4000:
            for p in[text[i:i+4000] for i in range(0,len(text),4000)]:
                await update.message.reply_text(p,parse_mode="Markdown")
        else: await update.message.reply_text(text,parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_find(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id!=ADMIN_ID: await update.message.reply_text("❌ Ruxsat yo'q!"); return
    if not ctx.args: await update.message.reply_text("Foydalanish: /find [id yoki @username]"); return
    try:
        arg=ctx.args[0]; conn=db(); c=conn.cursor()
        if arg.startswith("@"):
            c.execute("SELECT plan,expires_at,is_blocked,full_name,username,user_id,coins,total_messages,warning_count FROM users WHERE username=%s",(arg[1:],))
        else:
            c.execute("SELECT plan,expires_at,is_blocked,full_name,username,user_id,coins,total_messages,warning_count FROM users WHERE user_id=%s",(int(arg),))
        row=c.fetchone(); conn.close()
        if not row: await update.message.reply_text(f"❌ Topilmadi: {arg}"); return
        plan,exp,blk,fn,un,tid,coins,msgs,warns=row
        pem={"free":"🆓","standard":"⭐","premium":"💎"}
        exp_s=exp.strftime("%d.%m.%Y") if exp else "—"
        kb=[
            [InlineKeyboardButton("🆓",callback_data=f"ap_free_{tid}"),
             InlineKeyboardButton("⭐",callback_data=f"ap_std_{tid}"),
             InlineKeyboardButton("💎",callback_data=f"ap_prm_{tid}")],
            [InlineKeyboardButton("✅ Blokdan chiqarish" if blk else "🚫 Bloklash",
             callback_data=f"ap_unblock_{tid}" if blk else f"ap_block_{tid}")]
        ]
        await update.message.reply_text(
            f"👤 *Foydalanuvchi*\n━━━━━━━━━━━━━━━━━━━━\n"
            f"🆔 `{tid}`\n👤 {fn or '—'} | @{un or '—'}\n"
            f"{pem.get(plan,'🆓')} *{plan.upper()}* | 📅 {exp_s}\n"
            f"🪙 {coins:,} | 💬 {msgs:,} | ⚠️ {warns}\n"
            f"{'🚫 BLOKLANGAN' if blk else '✅ Faol'}",
            reply_markup=InlineKeyboardMarkup(kb),parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_broadcast(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id!=ADMIN_ID: await update.message.reply_text("❌ Ruxsat yo'q!"); return
    msg=" ".join(ctx.args)
    if not msg: await update.message.reply_text("Foydalanish: /broadcast [xabar]"); return
    await update.message.reply_text("⏳ Yuborilmoqda...")
    sent,failed=await broadcast_all(ctx.application,f"📢 {msg}")
    await update.message.reply_text(f"✅ Yuborildi: *{sent}* | ❌ Xato: *{failed}*",parse_mode="Markdown")

async def handle_maintenance(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id!=ADMIN_ID: await update.message.reply_text("❌ Ruxsat yo'q!"); return
    await update.message.reply_text("⏳ Texnik xizmat xabari yuborilmoqda...")
    msg="🔧 *Texnik Xizmat*\n━━━━━━━━━━━━━━━━━━━━\n\nBot vaqtinchalik ishlamaydi.\nTez orada qayta ishga tushadi! Uzr so'raymiz 🙏"
    sent,failed=await broadcast_all(ctx.application,msg)
    await update.message.reply_text(f"✅ Yuborildi: *{sent}* | ❌ Xato: *{failed}*",parse_mode="Markdown")

async def handle_createpromo(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id!=ADMIN_ID: await update.message.reply_text("❌ Ruxsat yo'q!"); return
    if len(ctx.args)<3:
        await update.message.reply_text(
            "📋 *Promo Kod Yaratish*\n\n"
            "`/createpromo [type] [qiymat] [son] [kod?]`\n\n"
            "Turlar: `standard` `premium` `coins`\n\n"
            "`/createpromo standard 7 50` — avtomatik kod\n"
            "`/createpromo premium 30 10 VIP2025` — maxsus kod\n"
            "`/createpromo coins 5000 100 BONUS`",parse_mode="Markdown"); return
    try:
        rtype,rval,maxu=ctx.args[0],int(ctx.args[1]),int(ctx.args[2])
        code=ctx.args[3].upper() if len(ctx.args)>3 else rnd(10)
        conn=db(); c=conn.cursor()
        c.execute("INSERT INTO promo_codes(code,reward_type,reward_value,max_uses,expires_at) VALUES(%s,%s,%s,%s,%s)",
                  (code,rtype,rval,maxu,datetime.now()+timedelta(days=30)))
        conn.commit(); conn.close()
        await update.message.reply_text(
            f"✅ *Promo Kod Yaratildi!*\n━━━━━━━━━━━━━━━━━━━━\n\n"
            f"🎟️ Kod: `{code}`\n📦 Tur: *{rtype}*\n"
            f"💰 Qiymat: *{rval}*\n👥 Son: *{maxu}*\n⏰ 30 kun",parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def handle_addcoins(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id!=ADMIN_ID: await update.message.reply_text("❌ Ruxsat yo'q!"); return
    if len(ctx.args)<2: await update.message.reply_text("Foydalanish: /addcoins [uid] [miqdor]"); return
    try:
        uid,amount=int(ctx.args[0]),int(ctx.args[1])
        add_coins(uid,amount)
        await update.message.reply_text(f"✅ `{uid}` ga *{amount:,}* coin qo'shildi!",parse_mode="Markdown")
    except Exception as e: await update.message.reply_text(f"Xato: {e}")

async def precheckout(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    await update.pre_checkout_query.answer(ok=True)

async def payment_success(update:Update,ctx:ContextTypes.DEFAULT_TYPE):
    user=update.effective_user
    payload=update.message.successful_payment.invoice_payload
    if payload=="std_stars":
        set_plan(user.id,"standard",30)
        await update.message.reply_text("✅ ⭐ *Standart* 30 kunga yoqildi! 🎉",parse_mode="Markdown")
    elif payload=="prm_stars":
        set_plan(user.id,"premium",30)
        await update.message.reply_text("✅ 💎 *Premium* 30 kunga yoqildi! 🎉",parse_mode="Markdown")
    elif payload=="grp_stars":
        await update.message.reply_text("✅ *Guruh rejimi sotib olindi!*\nBotni guruhga qo'shing va /activate\\_group yuboring.",parse_mode="Markdown")
    elif payload.startswith("img_"):
        prompt=payload[4:]
        await update.message.reply_text("⏳ AI rasm yaratilmoqda...")
        img=gen_image(prompt)
        if img:
            buf=io.BytesIO(img); buf.name="image.png"
            await update.message.reply_photo(photo=buf,caption=f"🎨 {prompt}")
        else: await update.message.reply_text("❌ Rasm yaratib bo'lmadi.")

async def post_init(app):
    init_db()
    await app.bot.set_my_commands([
        BotCommand("start","🚀 Botni boshlash"),
        BotCommand("ai_sound","🔊 AI Ovoz (Standart+)"),
        BotCommand("imagine","🎨 AI Rasm (Premium)"),
        BotCommand("pptx","📊 PowerPoint (Premium)"),
        BotCommand("word","📝 Word (Premium)"),
        BotCommand("cv","👤 CV (Standart+)"),
        BotCommand("email","📧 Email (Standart+)"),
        BotCommand("post","📱 Marketing post"),
        BotCommand("biznes","💼 Biznes reja"),
        BotCommand("translate","🌐 Tarjima"),
        BotCommand("code","💻 Kod yozish"),
        BotCommand("weather","🌦 Ob-havo"),
        BotCommand("crypto","💰 Crypto narxlar"),
        BotCommand("news","📰 Yangiliklar"),
        BotCommand("promo","🎟️ Promo kod"),
        BotCommand("language","🌐 Tilni o'zgartirish"),
        BotCommand("reset","🗑️ Tarixni tozalash"),
        BotCommand("activate_group","👥 Guruhni yoqish"),
    ])

def main():
    app=ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(post_init).build()
    app.add_handler(CommandHandler("start",start))
    app.add_handler(CommandHandler("ai_sound",handle_tts))
    app.add_handler(CommandHandler("imagine",handle_imagine))
    app.add_handler(CommandHandler("pptx",handle_pptx))
    app.add_handler(CommandHandler("word",handle_word))
    app.add_handler(CommandHandler("cv",handle_cv))
    app.add_handler(CommandHandler("email",handle_email))
    app.add_handler(CommandHandler("post",handle_post))
    app.add_handler(CommandHandler("biznes",handle_biznes))
    app.add_handler(CommandHandler("translate",handle_translate))
    app.add_handler(CommandHandler("code",handle_code))
    app.add_handler(CommandHandler("weather",handle_weather))
    app.add_handler(CommandHandler("crypto",handle_crypto))
    app.add_handler(CommandHandler("news",handle_news))
    app.add_handler(CommandHandler("reset",handle_reset))
    app.add_handler(CommandHandler("language",handle_language))
    app.add_handler(CommandHandler("activate_group",handle_activate_group))
    app.add_handler(CommandHandler("promo",handle_promo))
    app.add_handler(CommandHandler("admin",handle_admin))
    app.add_handler(CommandHandler("users",handle_users))
    app.add_handler(CommandHandler("find",handle_find))
    app.add_handler(CommandHandler("broadcast",handle_broadcast))
    app.add_handler(CommandHandler("maintenance",handle_maintenance))
    app.add_handler(CommandHandler("createpromo",handle_createpromo))
    app.add_handler(CommandHandler("addcoins",handle_addcoins))
    app.add_handler(PreCheckoutQueryHandler(precheckout))
    app.add_handler(MessageHandler(filters.SUCCESSFUL_PAYMENT,payment_success))
    app.add_handler(CallbackQueryHandler(callback))
    app.add_handler(MessageHandler(filters.TEXT&~filters.COMMAND,handle_text))
    app.add_handler(MessageHandler(filters.VOICE,handle_voice))
    app.add_handler(MessageHandler(filters.PHOTO,handle_photo))
    app.add_handler(MessageHandler(filters.Document.PDF,handle_document))
    print(f"🚀 {BOT_NAME} ishga tushdi!")
    app.run_polling()

if __name__=="__main__":
    main()
