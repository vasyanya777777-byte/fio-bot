import logging
import re
from io import BytesIO
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes
import pandas as pd

BOT_TOKEN = "8638225936:AAGbVBzEa6suQWPlK5IRIsMPEFjUstR3hKo"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

MALE_NAMES = {"александр","алексей","андрей","антон","артём","артем","борис","вадим","валентин","валерий","василий","виктор","виталий","владимир","владислав","вячеслав","геннадий","георгий","григорий","даниил","дмитрий","денис","евгений","иван","игорь","илья","кирилл","константин","леонид","максим","михаил","николай","никита","олег","павел","пётр","петр","роман","руслан","сергей","степан","тимур","фёдор","федор","филипп","юрий","яков","ярослав","лев","глеб","егор","семён","семен","тимофей","матвей","арсений","богдан","захар","макар","назар","савелий","тарас"}

FEMALE_NAMES = {"александра","алёна","алена","алина","алла","анастасия","анна","антонина","валентина","валерия","варвара","вера","виктория","галина","дарья","диана","екатерина","елена","елизавета","жанна","зинаида","зоя","инна","ирина","карина","кристина","ксения","лариса","лидия","людмила","маргарита","марина","мария","надежда","наталья","нина","оксана","олеся","ольга","полина","светлана","снежана","софья","тамара","татьяна","ульяна","юлия","яна","вероника","евгения","любовь","арина","василиса","злата","кира","милана"}

SLAVIC = re.compile(r"^[А-ЯЁа-яё\s\-]+$")

def get_gender(fio):
    if not isinstance(fio, str):
        return None
    parts = fio.strip().split()
    if len(parts) < 2:
        return None
    for part in reversed(parts[1:]):
        p = part.lower().rstrip(".")
        if p.endswith(("ович","евич","ич")):
            return "М"
        if p.endswith(("овна","евна","ична","инична")):
            return "Ж"
    name = parts[1].lower().rstrip(".")
    if name in MALE_NAMES:
        return "М"
    if name in FEMALE_NAMES:
        return "Ж"
    if name.endswith(("а","я")):
        return "Ж"
    return "М"

def is_slavic(fio):
    if not isinstance(fio, str) or not fio.strip():
        return False
    return bool(SLAVIC.match(fio.strip()))

def process(data):
    df = pd.read_excel(BytesIO(data))
    col = None
    for c in df.columns:
        if str(c).strip().upper() == "ФИО":
            col = c
            break
    if col is None:
        raise ValueError("Колонка ФИО не найдена")
    total = len(df)
    df = df[df[col].apply(is_slavic)].copy()
    removed = total - len(df)
    df["Пол"] = df[col].apply(get_gender)
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    out.seek(0)
    stats = {"total": total, "removed": removed, "left": len(df), "m": int((df["Пол"]=="М").sum()), "f": int((df["Пол"]=="Ж").sum())}
    return out.read(), stats

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.endswith((".xlsx",".xls")):
        await update.message.reply_text("Отправь файл .xlsx")
        return
    await update.message.reply_text("Обрабатываю...")
    try:
        f = await context.bot.get_file(doc.file_id)
        data = await f.download_as_bytearray()
        result, s = process(bytes(data))
        caption = f"Готово!\nВсего: {s['total']}\nУдалено не-славянских: {s['removed']}\nОсталось: {s['left']}\nМужчин: {s['m']}\nЖенщин: {s['f']}"
        await update.message.reply_document(BytesIO(result), filename="result.xlsx", caption=caption)
    except ValueError as e:
        await update.message.reply_text(f"Ошибка: {e}")
    except Exception as e:
        await update.message.reply_text(f"Ошибка: {e}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отправь Excel файл с колонкой ФИО")

app = ApplicationBuilder().token(BOT_TOKEN).build()
app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
app.add_handler(MessageHandler(filters.TEXT, handle_text))
app.run_polling()
