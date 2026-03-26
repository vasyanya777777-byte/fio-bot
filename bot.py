import logging import os import re import asyncio from io import BytesIO
from telegram import Update from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextType import pandas as pd
# ───────────────────────────────────────────── #  Настройки
# ───────────────────────────────────────────── BOT_TOKEN = "8638225936:AAGbVBzEa6suQWPlK5IRIsMPEFjUstR3hKo"  # @BotFather
logging.basicConfig(level=logging.INFO) logger = logging.getLogger(__name__)
# ───────────────────────────────────────────── #  Словари имён
# ─────────────────────────────────────────────
MALE_NAMES = {
    "александр","алексей","андрей","антон","аркадий","артём","артем","борис",
    "вадим","валентин","валерий","василий","виктор","виталий","владимир",
    "владислав","вячеслав","геннадий","георгий","григорий","даниил","дмитрий",
    "денис","евгений","иван","игорь","илья","кирилл","константин","леонид",
    "максим","михаил","николай","никита","олег","павел","пётр","петр","роман",
    "руслан","сергей","степан","тимур","фёдор","федор","филипп","юрий","яков",
    "ярослав","лев","глеб","егор","евгений","семён","семен","тимофей","матвей",
    "мирослав","всеволод","вадим","святослав","станислав","анатолий","аркадий",
    "арсений","богдан","вениамин","виталий","гавриил","давид","захар","зиновий",
    "иннокентий","кузьма","лаврентий","макар","назар","нестор","прохор","родион",
    "савелий","тарас","трофим","ульян","харитон","эдуард","эдуард", }
FEMALE_NAMES = {
    "александра","алёна","алена","алина","алла","анастасия","анна","антонина",
    "валентина","валерия","варвара","вера","виктория","галина","дарья","диана",
    "екатерина","елена","елизавета","жанна","зинаида","зоя","инга","инна",
    "ирина","карина","кристина","ксения","лариса","лидия","людмила","маргарита",
    "марина","мария","надежда","наталья","нина","оксана","олеся","ольга",
    "полина","светлана","снежана","софья","тамара","татьяна","ульяна","юлия",     "яна","вероника","евгения","любовь","людмила","нелли","нелли","регина",
    "римма","роза","рита","станислава","стелла","тина","элла","эльвира",
    "ангелина","арина","богдана","василиса","глафира","даниэла","евдокия",
    "злата","кира","лукерья","мила","мирослава","наина","нонна","пелагея",     "прасковья","серафима","степанида","таисия","устинья","фаина","феодора", }
# Типичные окончания отчеств
MALE_PATRONYMIC_ENDS = ("ович", "евич", "ич")
FEMALE_PATRONYMIC_ENDS = ("овна", "евна", "ична", "инична")
# Славянские символы — только кириллица + дефис + пробел
SLAVIC_PATTERN = re.compile(r"^[А-ЯЁа-яё\s\-]+$")
# ─────────────────────────────────────────────
#  Логика определения пола # ─────────────────────────────────────────────
def detect_gender(fio: str) -> str | None:
    """
    Возвращает 'М', 'Ж' или None если не удалось определить.     Анализирует имя и отчество из ФИО.
    """     if not isinstance(fio, str):         return None
    parts = fio.strip().split()     if len(parts) < 2:         return None
    # Стандартный порядок: Фамилия Имя Отчество     # Но отчество — самый надёжный признак     candidates = parts[1:]  # всё кроме фамилии
    # 1. Проверяем отчество (последнее или третье слово)     for part in reversed(candidates):         p = part.lower().rstrip(".")         if p.endswith(MALE_PATRONYMIC_ENDS):
            return "М"         if p.endswith(FEMALE_PATRONYMIC_ENDS):             return "Ж"
    # 2. Проверяем имя по словарю     if len(candidates) >= 1:
        name = candidates[0].lower().rstrip(".")         if name in MALE_NAMES:
            return "М"
        if name in FEMALE_NAMES:             return "Ж"
    # 3. Эвристика по окончанию имени     if len(candidates) >= 1:
        name = candidates[0].lower().rstrip(".")         if name.endswith(("й", "н", "м", "р", "л", "д", "г", "к", "п", "т", "в",             return "М"         if name.endswith(("а", "я")):             return "Ж"     return None
def is_slavic(fio: str) -> bool:     """Возвращает True если ФИО содержит только кириллические символы."""     if not isinstance(fio, str) or not fio.strip():         return False     return bool(SLAVIC_PATTERN.match(fio.strip()))
# ───────────────────────────────────────────── #  Обработка файла
# ─────────────────────────────────────────────
def process_excel(file_bytes: bytes) -> tuple[bytes, dict]:     df = pd.read_excel(BytesIO(file_bytes))
    # Ищем колонку ФИО (нечувствительно к регистру)     fio_col = None     for col in df.columns:         if str(col).strip().upper() == "ФИО":
            fio_col = col             break
    if fio_col is None:         raise ValueError("Колонка 'ФИО' не найдена в файле.")     total = len(df)
    # Фильтрация не-славянских     slavic_mask = df[fio_col].apply(is_slavic)     removed_non_slavic = (~slavic_mask).sum()     df = df[slavic_mask].copy()
    # Определение пола     df["Пол"] = df[fio_col].apply(detect_gender)     unknown_gender = df["Пол"].isna().sum()
    stats = {
        "total": total,
        "removed_non_slavic": int(removed_non_slavic),
        "remaining": len(df),
        "male": int((df["Пол"] == "М").sum()),
        "female": int((df["Пол"] == "Ж").sum()),
        "unknown_gender": int(unknown_gender),
    }
    # Сохраняем в байты     output = BytesIO()     with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)     output.seek(0)     return output.read(), stats
# ───────────────────────────────────────────── #  Telegram-хендлеры
# ─────────────────────────────────────────────
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document     fname = doc.file_name or ""
    if not (fname.endswith(".xlsx") or fname.endswith(".xls")):
        await update.message.reply_text("  Пожалуйста, отправь файл в формате .x         return     await update.message.reply_text("  Обрабатываю файл, подожди немного...")
    try:
        file = await context.bot.get_file(doc.file_id)         file_bytes = await file.download_as_bytearray()         result_bytes, stats = process_excel(bytes(file_bytes))
        caption = (             f" Готово!\n\n"             f" Всего строк: {stats['total']}\n"             f" Удалено не-славянских: {stats['removed_non_slavic']}\n"             f" Осталось: {stats['remaining']}\n"             f" Мужчин: {stats['male']}\n"             f" Женщин: {stats['female']}\n"
            f"  Пол не определён: {stats['unknown_gender']}"
        )
        await update.message.reply_document(             document=BytesIO(result_bytes),             filename="result.xlsx",             caption=caption,
        )
    except ValueError as e:
        await update.message.reply_text(f"  Ошибка: {e}")     except Exception as e:
        logger.exception("Ошибка при обработке файла")         await update.message.reply_text(f"  Неожиданная ошибка: {e}")
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):     await update.message.reply_text(
        "  Привет! Отправь мне Excel-файл (.xlsx) с колонкой *ФИО*.\n\n"
        "Я:\n"
        "• Уберу не-славянские имена\n"
        "• Определю пол каждой строки\n"         "• Верну обработанный файл",         parse_mode="Markdown",
    )
# ───────────────────────────────────────────── #  Запуск
# ─────────────────────────────────────────────
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()     app.add_handler(MessageHandler(filters.Document.ALL, handle_document))     app.add_handler(MessageHandler(filters.TEXT, handle_text))     logger.info("Бот запущен...")     app.run_polling()
if __name__ == "__main__":
    main()
