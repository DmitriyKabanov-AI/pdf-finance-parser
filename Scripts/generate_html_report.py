# generate_html_report.py
# Запуск: python generate_html_report.py
# Требования: pip install pandas openpyxl

import pandas as pd
import numpy as np
import json
import re
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path

# ═══════════════════════════════════════════════════════
#  КОНФИГУРАЦИЯ
# ═══════════════════════════════════════════════════════
SCRIPT_DIR     = Path(__file__).parent.parent.absolute()
EXCEL_FILE     = SCRIPT_DIR / "reports" / "transactions_consolidated.xlsx"
TEMPLATE_FILE  = SCRIPT_DIR / "Template" / "dashboard_template.html"
OUTPUT_FILE    = SCRIPT_DIR / "reports" / "FinPulse_Interactive_Dashboard.html"

# ── Внутренние переводы — ИСКЛЮЧАЕМ ──
INTERNAL_TRANSFER_KEYWORDS = [
    "внутрибанковский перевод", "внутренний перевод",
    "перевод с договора", "перевод на ",
]
EXTERNAL_INCOME_WHITELIST = [
    "банк атб", "атб", "atb", "зарплата", "salary", "аванс", "оклад",
    "внесение наличных", "банкомат", "кэшбэк", "cashback", "возврат", "refund",
]

CATEGORY_RULES = [
    # ─── Доходы / Поступления ───
    ("Кэшбэк/Возврат", [
        "кэшбэк за обычные",
        "кэшбэк за покупки",
        "отмена операции",
        "cashback",
        "возврат",
        "refund",
    ]),
    ("Внесение наличных", [
        "внесение наличных",
        "банкомат т-банк",
        "внесение через банкомат",
    ]),

    # ─── Внутренние переводы (между своими счетами) ───
    ("Внутренний перевод", [
        "внутренний перевод на",
        "внутрибанковский перевод",
        "внутренний перевод",
    ]),

    # ─── Переводы исходящие (внешние) ───
    ("Переводы исходящие", [
        "внешний перевод по номеру телефона",
        "внешний перевод",
    ]),

    # ─── Транспорт ───
    ("Метро / Транспортная карта", [
        "mos.transport",
        "mosgortrans",
        "strelkacard",
        "avtobusnyj park",
        "автобусный парк",
        "тройка",
        "мосметро",
        "московский метрополитен",
    ]),
    ("Самокат / Велосипед (аренда)", [
        "whoosh",
        "whoosh.bike",
        "whoosh moskva",
        "ymsurent",
        "ymurent",
        "urent",
        "мтс юрент",
        "ymenergo",        # зарядка самоката
        "ym*energo",
        "ymenergo",
        "energo",
    ]),
    ("Такси / Каршеринг", [
        "яндекс такси",
        "yandex taxi",
        "яндекс go",
        "yandex go",
        "ситимобил",
        "citimobil",
        "gett",
        "uber",
        "wheely",
        "bolt",
        "делимобиль",
        "ситидрайв",
        "citydrive",
        "каршеринг",
    ]),

    # ─── Еда — доставка ───
    ("Доставка еды", [
        "яндекс еда",
        "yandex eda",
        "yandex5411edarit",   # Яндекс Еда (реальный мерчант)
        "delivery club",
        "самокат",
        "samokat",
        "лавка",
        "lavka",
        "sushibox",           # суши-бокс доставка
        "dodo pizza",
        "dodopizza",
        "dodopizzaspeterburg",
    ]),

    # ─── Еда — фастфуд / кафе / рестораны ───
    ("Фастфуд", [
        "burger king",
        "бургер кинг",
        "kfc",
        "kfs",                # реальное написание KFC в выписке
        "макдоналдс",
        "mcdonalds",
        "вкусно и точка",
        "vkusnoitochka",
        "ymvkusnoitochka",    # Яндекс Вкусно и точка
        "ymvkusnoitochka",
        "y.mvkusnoitochka",
        "teremok",
        "теремок",
        "subway",
        "шаурма",
        "evo_.shaverma",
        "главфудгрупп",
        "glavfudgrupp",
        "qsr",                # быстрое питание (QSR-терминалы)
        "stolovka",
        "stolovka.ru",
        "столовая",
    ]),
    ("Кафе / Рестораны", [
        "restoran",
        "ресторан",
        "кафе",
        "cafe",
        "coffee",
        "starbucks",
        "шоколадница",
        "coffeehouse",
        "буфет",
        "bar ",
        "паб",
        "pub",
        "кофейня",
        "coffee to go",       # реальный мерчант в выписке
        "evo_khochukofe",     # «Хочу кофе»
        "chay ine chay",      # «Чай не чай»
        "surf coffee",
        "kofeynaya kantata",  # Кофейная кантата
        "kofiks",
        "blok pitania",       # Блок питания (кофейня СПб)
        "co&lo",              # кофейня
        "moroshka",           # бар/кафе
        "moroshka.",
        "moroshka..",
        "prosto vasya",       # кафе/бар
        "gruzinskaya kukhnya",# грузинский ресторан
        "restoran botkinskaya",
        "varochnaya",
        "ramen&wok",
        "kikukhana",
        "restoran kikukhana",
        "chajkhona",          # чайхана
        "sever metropol",
        "pupuer yunnan",
        "chabrets",
        "tokyo city",
        "stejk davaj",        # стейк-бар
        "the bull",           # ресторан
        "bay bo",             # кафе
        "shoko",
        "shoko atrium",
        "shoko_ogorodnyi",
        "farsh",              # ресторан Фарш
        "milette",            # ресторан
        "restimaks",
        "myata",              # кальянная/ресторан
        "oo derkul",
        "ooo derkul",
        "cider",
        "luchsheye mesto",
        "ooo luchsheye",
        "cvetnoj bulvar",
        "lyubim",
        "ooo veles",
        "oo koro",
        "ooo koro",
        "rxt",                # бар/ресторан
        "skyfall",            # заведение
        "evo_art-kafe",
        "au1033",             # неопознанный мерчант кафе
        "ip petrov sa",       # ИП Петров (кафе/еда)
        "ip frolov",          # ИП Фролов (кафе/еда)
        "ip morozova",
        "m-resting",
        "ooo m-resting",      # М-Рестинг (клуб/ресторан)
        "krasnodars. paren",  # кафе
        "pavilon optima",
        "xplatip",            # мерчант
        "kofe",
        "чай",
    ]),

    # ─── Продукты ───
    ("Продукты", [
        "пятёрочка",
        "пятерочка",
        "pyaterochka",
        "перекрёсток",
        "перекресток",
        "perekrestok",
        "вкусвилл",
        "vkusvill",
        "ашан",
        "auchan",
        "лента",
        "lenta",
        "tk lenta",           # реальное написание Ленты в выписке
        "дикси",
        "dixy",
        "магнит",
        "magnit",
        "magnit mm",          # Магнит у дома
        "монетка",
        "monetka",
        "авоська",
        "avoska",
        "universam avoska",
        "универсам",
        "азбука вкуса",
        "azbukavkusa",
        "av azbukavkusa",
        "магнолия",
        "magnoliya",
        "мираторг",
        "верный",
        "vernyj",
        "billa",
        "spar",
        "atak",               # магазин у дома (Атак/Дикси)
        "fasol",              # Фасоль (магазин)
        "da s77",             # продуктовый
        "suhofrukty",         # сухофрукты
        "metro store",        # магазин Метро
        "metro cash",         # Метро оптовый
        "metro tpp",          # терминал оплаты Метро
        "norman",             # магазин
        "tdrealsalova",       # магазин (ТД Реал Салова, СПб)
        "ozon fresh",
        "глобус",
        "globus",
    ]),

    # ─── Онлайн-покупки / Маркетплейсы ───
    ("Онлайн-покупки", [
        "wildberries",
        "wbwildberries",
        "wb*",
        "ozon",
        "озон",
        "aliexpress",
        "алиэкспресс",
        "avito",
        "авито",
        "lamoda",
        "ym*",
        "yandex market",
        "яндекс маркет",
        "yandex4112rasp",     # Яндекс Расписание (ж/д билеты)
        "ymmmamos",           # Яндекс Маркет
        "ymsutochno",         # Суточно.ру (посуточная аренда)
        "md.aviasales",       # Авиасейлс
        "remanga",            # сервис манги/комиксов
        "author.today",       # Автор Тудей (книги/контент)
        "sokratic",           # онлайн-сервис
        "gptunnel",           # VPN/сервис
        "ibank.t2",           # мобильный банк Т2 (Tele2)
        "ibank.mts",          # МТС Банк/сервис
    ]),
    ("Одежда / Обувь", [
        "zara",
        "h&m",
        "uniqlo",
        "adidas",
        "nike",
        "спортмастер",
        "gloria jeans",
        "befree",
        "reserved",
        "mango",
        "magazin odezhda",    # «Магазин одежда» в выписке
        "детский мир",
        "detskiy mir",
    ]),

    # ─── Развлечения / Досуг ───
    ("Развлечения / Досуг", [
        "кино",
        "cinema",
        "театр",
        "concert",
        "концерт",
        "билет",
        "ticketland",
        "afisha",
        "парк",
        "аттракцион",
        "музей",
        "выставка",
        "tc okhotnyy",        # ТЦ Охотный ряд
        "cvetnoj bulvar",     # ТЦ Цветной бульвар
        "best western zoomhotel", # зоо-отель (развлечение/проживание)
    ]),
    ("Подписки / Контент", [
        "spotify",
        "netflix",
        "яндекс плюс",
        "yandex plus",
        "vkvk music",         # реальный мерчант VK Музыка в выписке
        "vk музыка",
        "apple music",
        "google play",
        "notion",
        "icloud",
        "подписка",
        "t-bank.t-bundle",
        "t-bundle",
        "bundle",
        "premium",
        "кинопоиск",
        "wink",
        "ivi",
        "okko",
        "more.tv",
        "youtube",
        "sber5732sberdevices", # Сбер устройства/подписки
    ]),

    # ─── Связь / Интернет ───
    ("Связь / Интернет", [
        "мтс",
        "mts",
        "ibank.mts",
        "ibank.t2",
        "tele2",
        "теле2",
        "t2",
        "билайн",
        "beeline",
        "мегафон",
        "megafon",
        "ростелеком",
        "internet",
        "интернет",
        "провайдер",
        "yota",
        "услуги ibank",       # оплата мобильной связи через iBank
    ]),

    # ─── Здоровье ───
    ("Здоровье / Аптека", [
        "аптека",
        "pharmacy",
        "apteka",
        "yandex5912apteki",   # Яндекс Аптека (реальный мерчант)
        "доктор",
        "клиника",
        "clinic",
        "больница",
        "hospital",
        "медицина",
        "стоматология",
        "denta",
        "лаборатория",
        "анализ",
        "shlend",             # аптека/медицина (ШЛЕНД)
    ]),

    # ─── Красота / Уход ───
    ("Красота / Уход", [
        "парикмахер",
        "салон красоты",
        "маникюр",
        "барбер",
        "cosmetic",
        "косметика",
        "л'этуаль",
        "letual",
        "рив гош",
        "golden apple",
    ]),

    # ─── Жильё / Проживание ───
    ("Жильё / Проживание", [
        "best western zoomhotel",  # отель/хостел
        "ymsutochno",              # Суточно.ру
        "hotel",
        "hostel",
        "хостел",
        "отель",
    ]),

    # ─── Защита карты / Страховка ───
    ("Защита карты", [
        "защита карты",
        "плата за предоставление услуги защита",
        "страховка карты",
    ]),

    # ─── Штрафы / Пени ───
    ("Штрафы / Пени", [
        "штраф",
        "гибдд",
        "мади",
        "ампп",
        "пеня",
        "пени",
        "госпошлина",
    ]),

    # ─── Комиссии банка ───
    ("Комиссии банка", [
        "комиссия",
        "банковская комиссия",
        "проценты за просрочку",
        "брокерская комиссия",
    ]),

    # ─── Образование / Контент ───
    ("Образование / Книги", [
        "курс",
        "обучение",
        "школа",
        "университет",
        "udemy",
        "coursera",
        "skillbox",
        "нетология",
        "geekbrains",
        "author.today",   # книги/авторский контент
        "remanga",        # манга/комиксы
        "sokratic",
    ]),

    # ─── Прочее ───
    ("Прочее", []),  # fallback — всё, что не попало выше
]

# ── Вспомогательные списки для аналитики ──

INTERNAL_TRANSFER_KEYWORDS = [
    "внутренний перевод на",
    "внутрибанковский перевод",
]

EXTERNAL_TRANSFER_KEYWORDS = [
    "внешний перевод по номеру телефона",
    "внешний перевод",
]

INCOME_KEYWORDS = [
    "кэшбэк",
    "возврат",
    "отмена операции",
    "внесение наличных",
    "с договора",  # входящий внутрибанковский перевод
]

TAXI_KEYWORDS = [
    "яндекс такси", "yandex taxi", "яндекс go", "yandex go",
    "ситимобил", "citimobil", "gett", "uber", "wheely", "bolt",
    "делимобиль", "ситидрайв", "citydrive", "каршеринг",
]

SCOOTER_KEYWORDS = [
    "whoosh", "whoosh.bike", "whoosh moskva",
    "ymurent", "ymsurent", "urent", "мтс юрент",
]

METRO_KEYWORDS = [
    "mos.transport", "mosgortrans", "strelkacard",
    "avtobusnyj park", "автобусный парк",
    "тройка", "мосметро", "московский метрополитен",
]

CHARGING_KEYWORDS = [
    "ymenergo", "ym*energo", "yMenergo",
    "energo", "charge", "зарядка",
]

FINE_KEYWORDS = [
    "штраф", "гибдд", "мади", "ампп", "пеня", "пени",
    "госпошлина", "защита карты",
    "проценты за просрочку",
]

IMPULSIVE_KEYWORDS = [
    # доставка
    "yandex5411edarit", "dodopizza", "dodo pizza",
    "самокат", "лавка",
    # фастфуд
    "kfs", "kfc", "vkusnoitochka", "ymvkusnoitochka", "teremok",
    # онлайн
    "wildberries", "wbwildberries", "ozon", "aliexpress",
    "author.today", "remanga",
    # развлечения
    "whoosh", "ymurent",
    # рестораны вне дома поздно ночью — доп. признак в коде
]


# ═══════════════════════════════════════════════════════
#  ПАРСЕР EXCEL
# ═══════════════════════════════════════════════════════
def load_excel(path: Path) -> pd.DataFrame:
    print(f"📂 Загружаю: {path}")
    df = pd.read_excel(path, header=0)
    print(f"  ✓ Строк: {len(df)}, колонок: {len(df.columns)}")
    return df


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    col_map = {}
    EXACT = {
        'source_file': 'source', 'datetime_operation': 'date_op',
        'datetime_writeoff': 'date_proc', 'amount_operation': 'amount_orig',
        'amount_card': 'amount', 'description': 'description',
        'card_number': 'card',
    }
    for col in cols:
        key = str(col).lower().strip()
        if key in EXACT:
            col_map[col] = EXACT[key]
    PATTERNS = [
        ('source',      ['source','файл','pdf','источник']),
        ('date_op',     ['datetime_op','date_op','дата опер','дата и время опер']),
        ('date_proc',   ['datetime_write','date_proc','дата обраб','дата списания']),
        ('amount_orig', ['amount_op','amount_orig','сумма в вал','сумма операции']),
        ('amount',      ['amount_card','amount_rub','сумма в руб','сумма по карте']),
        ('description', ['description','описание','назначение','наименование']),
        ('card',        ['card_number','card','карта','договор','account']),
    ]
    already = set(col_map.values())
    for col in cols:
        if col in col_map:
            continue
        key = str(col).lower().strip()
        for target, pats in PATTERNS:
            if target in already:
                continue
            if any(p in key for p in pats):
                col_map[col] = target
                already.add(target)
                break
    POS = {0:'source',1:'date_op',2:'date_proc',3:'amount_orig',
           4:'amount',5:'description',6:'card'}
    already = set(col_map.values())
    for i, col in enumerate(cols):
        if col in col_map:
            continue
        if i in POS and POS[i] not in already:
            col_map[col] = POS[i]
            already.add(POS[i])
    df = df.rename(columns=col_map)
    for req in ['date_op', 'amount', 'description']:
        if req not in df.columns:
            raise ValueError(f"Колонка '{req}' не найдена. Доступные: {list(df.columns)}")
    return df


def parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    for col in ['date_op', 'date_proc']:
        if col not in df.columns:
            continue
        if not pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = pd.to_datetime(df[col].astype(str), dayfirst=True, errors='coerce')
    if 'date_op' not in df.columns and 'date_proc' in df.columns:
        df['date_op'] = df['date_proc']
    df = df.dropna(subset=['date_op']).copy()
    df['hour']      = df['date_op'].dt.hour
    df['weekday']   = df['date_op'].dt.dayofweek
    df['dow_name']  = df['date_op'].dt.day_name()
    df['day']       = df['date_op'].dt.date
    df['month']     = df['date_op'].dt.to_period('M')
    df['date_str']  = df['date_op'].dt.strftime('%d.%m.%Y %H:%M')
    print(f"  ✓ Период: {df['date_op'].min().date()} — {df['date_op'].max().date()}")
    return df


def parse_amounts(df: pd.DataFrame) -> pd.DataFrame:
    def clean(v):
        if pd.isna(v): 
            return 0.0
        s = str(v).replace(' ','').replace('\xa0','').replace(',','.')
        s = re.sub(r'[^\d.\-+]', '', s)
        try: 
            return float(s)
        except ValueError: 
            return 0.0
    df['amount']     = df['amount'].apply(clean)
    df['amount_abs'] = df['amount'].abs()
    df['is_income']  = df['amount'] > 0
    df['is_expense'] = df['amount'] < 0
    return df


def filter_internal_transfers(df: pd.DataFrame) -> pd.DataFrame:
    def is_internal(desc):
        if pd.isna(desc): 
            return False
        d = str(desc).lower()
        for w in EXTERNAL_INCOME_WHITELIST:
            if w in d: 
                return False
        for kw in INTERNAL_TRANSFER_KEYWORDS:
            if kw in d: 
                return True
        return False
    mask = df['description'].apply(is_internal)
    removed = mask.sum()
    if removed:
        print(f"  ✓ Исключено внутренних переводов: {removed}")
    return df[~mask].copy()


def categorize(df: pd.DataFrame) -> pd.DataFrame:
    def get_cat(desc):
        if pd.isna(desc): 
            return "Прочее"
        d = str(desc).lower()
        for cat_name, kws in CATEGORY_RULES:
            if any(k in d for k in kws):
                return cat_name
        return "Прочее"

    def is_taxi(desc):
        if pd.isna(desc): 
            return False
        d = str(desc).lower()
        return any(k in d for k in TAXI_KEYWORDS)

    def is_metro(desc):
        if pd.isna(desc): 
            return False
        d = str(desc).lower()
        return any(k in d for k in METRO_KEYWORDS)

    def is_fine(desc):
        if pd.isna(desc): 
            return False
        d = str(desc).lower()
        return any(k in d for k in FINE_KEYWORDS)

    def is_imp(desc):
        if pd.isna(desc): 
            return False
        d = str(desc).lower()
        return any(k in d for k in IMPULSIVE_KEYWORDS)

    df['category']     = df['description'].apply(get_cat)
    df['is_taxi']      = df['description'].apply(is_taxi)
    df['is_metro']     = df['description'].apply(is_metro)
    df['is_fine']      = df['description'].apply(is_fine)
    df['is_impulsive'] = df['description'].apply(is_imp)

    # Проверим долю «Прочее»
    exp = df[df['is_expense']]
    other_pct = (exp[exp['category'] == 'Прочее']['amount_abs'].sum() /
                 exp['amount_abs'].sum() * 100) if len(exp) > 0 else 0
    print(f"  ✓ Доля «Прочее» в расходах: {other_pct:.1f}%")
    cats = exp['category'].value_counts()
    print(f"  ✓ Топ категорий: {dict(list(cats.items())[:8])}")
    return df


def classify_period(hour: int) -> str:
    if   0 <= hour <  8: 
        return "night"
    elif 8 <= hour < 16: 
        return "day"
    else:                
        return "evening"


# ═══════════════════════════════════════════════════════
#  АНАЛИТИКА
# ═══════════════════════════════════════════════════════
def fmt_r(v):
    """Форматирование числа с пробелами-разделителями."""
    try:
        v = float(v)
        if v == int(v): 
            return f"{int(v):,}".replace(",", " ")
        return f"{v:,.2f}".replace(",", " ")
    except Exception: 
        return str(v)


def compute_analytics(df: pd.DataFrame) -> dict:
    print("\n📊 Вычисляю аналитику...")
    A = {}

    exp = df[df['is_expense']].copy()
    inc = df[df['is_income']].copy()
    exp['amount_pos'] = exp['amount'].abs()

    months_order = sorted(df['month'].unique())
    month_labels = [str(m) for m in months_order]
    n_months = max(len(months_order), 1)

    # ── KPI ──
    A['total_income']   = round(float(inc['amount'].sum()), 2)
    A['total_expense']  = round(float(exp['amount_pos'].sum()), 2)
    A['net_balance']    = round(A['total_income'] - A['total_expense'], 2)
    A['save_rate']      = round(A['net_balance'] / A['total_income'] * 100, 1) if A['total_income'] > 0 else 0
    A['coverage_ratio'] = round(A['total_income'] / A['total_expense'], 2) if A['total_expense'] > 0 else 0
    imp_sum = float(exp[exp['is_impulsive']]['amount_pos'].sum())
    A['impulsive_sum']  = round(imp_sum, 2)
    A['impulsive_pct']  = round(imp_sum / A['total_expense'] * 100, 1) if A['total_expense'] > 0 else 0
    avg_daily_inc = A['total_income'] / max(df['day'].nunique(), 1)
    daily_exp = exp.groupby('day')['amount_pos'].sum()
    A['leaky_days'] = int((daily_exp > avg_daily_inc * 1.2).sum())

    A['date_from']   = df['date_op'].min().strftime('%d.%m.%Y')
    A['date_to']     = df['date_op'].max().strftime('%d.%m.%Y')
    A['months_count'] = n_months
    A['avg_monthly_income']  = round(A['total_income']  / n_months, 2)
    A['avg_monthly_expense'] = round(A['total_expense'] / n_months, 2)

    # ── Месячная динамика ──
    monthly_inc = inc.groupby('month')['amount'].sum().reindex(months_order, fill_value=0)
    monthly_exp = exp.groupby('month')['amount_pos'].sum().reindex(months_order, fill_value=0)
    # Формируем метки «Янв 25», «Фев 25» и т.д.
    RU_MONTHS = ['Янв','Фев','Мар','Апр','Май','Июн',
                 'Июл','Авг','Сен','Окт','Ноя','Дек']
    month_labels_ru = []
    for m in months_order:
        dt = m.to_timestamp()
        month_labels_ru.append(f"{RU_MONTHS[dt.month-1]} {dt.strftime('%y')}")

    A['months']          = month_labels_ru
    A['months_raw']      = month_labels
    A['monthly_income']  = [round(float(v), 2) for v in monthly_inc]
    A['monthly_expense'] = [round(float(v), 2) for v in monthly_exp]
    A['monthly_balance'] = [round(float(i - e), 2)
                            for i, e in zip(A['monthly_income'], A['monthly_expense'])]

    # ── Накопленный баланс (подневной) ──
    all_days = sorted(df['day'].unique())
    daily_net = df.groupby('day')['amount'].sum()
    cumbal, daily_bal = 0.0, []
    daily_dates = []
    for d in all_days:
        cumbal += float(daily_net.get(d, 0))
        daily_bal.append(round(cumbal, 2))
        daily_dates.append(str(d))
    A['daily_balance'] = daily_bal
    A['daily_dates']   = daily_dates
    # Медиана 30 дней
    med30 = []
    for i in range(len(daily_bal)):
        sl = daily_bal[max(0, i-30):i+1]
        med30.append(round(float(np.median(sl)), 2))
    A['daily_med30'] = med30

    # ── Тепловая карта 24ч x 7 дней ──
    weekday_names = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс']
    heatmap_z = []
    for d in range(7):
        row = []
        for h in range(24):
            mask = (exp['hour'] == h) & (exp['weekday'] == d)
            val = float(exp.loc[mask, 'amount_pos'].mean()) if mask.any() else 0
            row.append(round(val, 2))
        heatmap_z.append(row)
    A['heatmap_z']    = heatmap_z
    A['heatmap_days'] = weekday_names

    # ── Периоды суток ──
    exp = exp.copy()
    exp['time_period'] = exp['hour'].apply(classify_period)
    period_keys  = ['night', 'day', 'evening']
    period_names = ['Ночь (00-08)', 'День (08-16)', 'Вечер (16-00)']

    period_stats = {}
    for pk in period_keys:
        sub = exp[exp['time_period'] == pk]
        total = float(sub['amount_pos'].sum())
        count = int(len(sub))
        pct = round(total / A['total_expense'] * 100, 1) if A['total_expense'] > 0 else 0
        top = sub.groupby('category')['amount_pos'].sum().nlargest(5)
        period_stats[pk] = {
            'sum': round(total, 2), 'count': count, 'pct': pct,
            'avg': round(total / count, 2) if count > 0 else 0,
            'top_cats_labels': list(top.index),
            'top_cats_values': [round(float(v), 2) for v in top.values]
        }
    A['period_stats'] = period_stats
    A['period_names'] = period_names

    # ── Ночь (00-06) ──
    night_old = exp[exp['hour'].isin(range(0, 6))]
    A['night_sum']   = round(float(night_old['amount_pos'].sum()), 2)
    A['night_pct']   = round(A['night_sum'] / A['total_expense'] * 100, 1) if A['total_expense'] > 0 else 0
    A['night_count'] = int(len(night_old))

    # ── Категории ──
    cat_stats = (exp.groupby('category')
                    .agg(total=('amount_pos','sum'),
                         count=('amount_pos','count'),
                         avg=('amount_pos','mean'))
                    .sort_values('total', ascending=False))
    A['cat_labels'] = list(cat_stats.index)
    A['cat_totals'] = [round(float(v), 2) for v in cat_stats['total']]
    A['cat_counts'] = [int(v) for v in cat_stats['count']]
    A['cat_avgs']   = [round(float(v), 2) for v in cat_stats['avg']]

    # ── Часовое распределение ──
    hourly = (exp.groupby('hour')
                 .agg(total=('amount_pos','sum'), count=('amount_pos','count'))
                 .reindex(range(24), fill_value=0))
    A['hourly_amounts'] = [round(float(v), 2) for v in hourly['total']]
    A['hourly_counts']  = [int(v) for v in hourly['count']]

    # ── Доставка еды ──
    del_exp = exp[exp['category'] == 'Доставка еды']
    del_m = del_exp.groupby('month')['amount_pos'].sum().reindex(months_order, fill_value=0)
    A['delivery_monthly'] = [round(float(v), 2) for v in del_m]
    A['delivery_total']   = round(float(del_exp['amount_pos'].sum()), 2)
    A['delivery_count']   = int(len(del_exp))
    A['delivery_avg']     = round(float(del_exp['amount_pos'].mean()), 2) if len(del_exp) > 0 else 0
    # Средний чек по дням недели
    DOW_RU = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс']
    del_dow = []
    for d in range(7):
        sub = del_exp[del_exp['weekday'] == d]
        del_dow.append(round(float(sub['amount_pos'].mean()), 2) if len(sub) > 0 else 0)
    A['delivery_dow_avg'] = del_dow
    A['delivery_dow_labels'] = DOW_RU
    # Структура доставки (по описанию)
    if len(del_exp) > 0:
        del_desc_cats = {}
        for _, r in del_exp.iterrows():
            d = str(r['description']).lower()
            if 'dodo' in d:          
                k = 'Dodo Pizza'
            elif 'kfc' in d:         
                k = 'KFC'
            elif 'бургер' in d or 'burger' in d: 
                k = 'Burger King'
            elif 'вкусно' in d:      
                k = 'Вкусно и точка'
            elif 'яндекс еда' in d:  
                k = 'Яндекс Еда'
            elif 'самокат' in d or 'samokat' in d: 
                k = 'Самокат'
            elif 'лавка' in d:       
                k = 'Лавка'
            elif 'пицца' in d or 'pizza' in d: 
                k = 'Пиццерии'
            elif 'суши' in d or 'sushi' in d:  
                k = 'Суши'
            else:                    
                k = 'Другое'
            del_desc_cats[k] = del_desc_cats.get(k, 0) + abs(float(r['amount_pos']))
        sorted_dc = sorted(del_desc_cats.items(), key=lambda x: x[1], reverse=True)
        A['delivery_pie_labels'] = [x[0] for x in sorted_dc]
        A['delivery_pie_values'] = [round(x[1], 2) for x in sorted_dc]
    else:
        A['delivery_pie_labels'] = []
        A['delivery_pie_values'] = []

    # ── Такси vs Метро ──
    taxi_exp  = exp[exp['is_taxi']]
    metro_exp = exp[exp['is_metro']]
    A['taxi_total']  = round(float(taxi_exp['amount_pos'].sum()), 2)
    A['taxi_count']  = int(len(taxi_exp))
    A['metro_total'] = round(float(metro_exp['amount_pos'].sum()), 2)
    A['metro_count'] = int(len(metro_exp))
    taxi_m  = taxi_exp.groupby('month')['amount_pos'].sum().reindex(months_order, fill_value=0)
    metro_m = metro_exp.groupby('month')['amount_pos'].sum().reindex(months_order, fill_value=0)
    A['taxi_monthly']  = [round(float(v), 2) for v in taxi_m]
    A['metro_monthly'] = [round(float(v), 2) for v in metro_m]

    # ── Кофе / Кафе ──
    coffee_df = exp[(exp['amount_pos'] <= 500) &
                    (exp['category'].isin(['Кафе/Рестораны','Фастфуд']))]
    A['coffee_count']   = int(len(coffee_df))
    A['coffee_total']   = round(float(coffee_df['amount_pos'].sum()), 2)
    A['coffee_avg']     = round(float(coffee_df['amount_pos'].mean()), 2) if len(coffee_df) > 0 else 0
    coffee_m = coffee_df.groupby('month')['amount_pos'].sum().reindex(months_order, fill_value=0)
    A['coffee_monthly'] = [round(float(v), 2) for v in coffee_m]

    # ── Штрафы / Комиссии ──
    fines_df = exp[exp['is_fine']]
    A['fines_total'] = round(float(fines_df['amount_pos'].sum()), 2)
    A['fines'] = [
        {'date': r['date_str'], 'desc': str(r['description'])[:80],
         'amount': round(float(r['amount_pos']), 2), 'cat': r['category']}
        for _, r in fines_df.iterrows()
    ]
    # Группировка по подтипу
    fine_by_cat = fines_df.groupby('category')['amount_pos'].sum().sort_values(ascending=False)
    A['fines_by_cat_labels'] = list(fine_by_cat.index)
    A['fines_by_cat_values'] = [round(float(v), 2) for v in fine_by_cat.values]

    # ── Траты после зарплаты ──
    income_dates = inc.sort_values('date_op')['date_op'].tolist()
    b24, b48, b72 = [], [], []
    for dt in income_dates:
        for hrs, lst in [(24, b24), (48, b48), (72, b72)]:
            m = (exp['date_op'] > dt) & (exp['date_op'] <= dt + timedelta(hours=hrs))
            lst.append(float(exp[m]['amount_pos'].sum()))
    avg_daily_exp = A['total_expense'] / max(df['day'].nunique(), 1)
    A['salary_burst'] = {
        '24h': round(float(np.mean(b24)) if b24 else 0, 2),
        '48h': round(float(np.mean(b48)) if b48 else 0, 2),
        '72h': round(float(np.mean(b72)) if b72 else 0, 2),
        'normal': round(avg_daily_exp, 2)
    }

    # ── «Дырявые» категории (waterfall) ──
    cat_monthly = exp.groupby(['month','category'])['amount_pos'].sum().unstack(fill_value=0)
    cat_growth = {}
    for cat in cat_monthly.columns:
        vals = cat_monthly[cat].values
        if len(vals) >= 2 and vals[0] > 0:
            cat_growth[cat] = round(float((vals[-1] - vals[0]) / vals[0] * 100), 1)
        else:
            cat_growth[cat] = 0
    top_leaky = sorted(cat_growth, key=lambda c: cat_growth[c], reverse=True)[:5]
    lc_colors = ['#ef4444','#f97316','#a855f7','#3b82f6','#eab308']
    A['leaky_cats'] = []
    for i, cat in enumerate(top_leaky):
        mv = cat_monthly[cat].reindex(months_order, fill_value=0)
        vals = [round(float(v), 2) for v in mv.values]
        # Waterfall: изменения месяц-к-месяцу
        changes = [0] + [round(vals[j] - vals[j-1], 2) for j in range(1, len(vals))]
        A['leaky_cats'].append({
            'name': cat, 'growth': cat_growth[cat],
            'values': vals, 'changes': changes, 'color': lc_colors[i]
        })

    # ── Подписки ──
    small_exp = exp[(exp['amount_pos'] >= 50) & (exp['amount_pos'] <= 5000)].copy()
    subs_cands = []
    key_col = small_exp['description'].str.slice(0,30).str.lower().str.strip()
    for _, group in small_exp.groupby(key_col):
        if len(group) >= 2 and group['month'].nunique() >= 2:
            subs_cands.append({
                'name':       str(group['description'].iloc[0])[:55],
                'count':      int(len(group)),
                'avg_amount': round(float(group['amount_pos'].mean()), 2),
                'total':      round(float(group['amount_pos'].sum()), 2),
                'months_seen': int(group['month'].nunique()),
                'last_date':  group['date_op'].max().strftime('%d.%m.%Y')
            })
    subs_cands.sort(key=lambda x: x['avg_amount'], reverse=True)
    A['subscriptions'] = subs_cands[:15]

    # ── Sankey ──
    inc_by_cat = inc.groupby('category')['amount'].sum().sort_values(ascending=False)
    exp_by_cat = exp.groupby('category')['amount_pos'].sum().sort_values(ascending=False).head(12)
    inc_cats = [c for c in inc_by_cat.index if inc_by_cat[c] > 0]
    exp_cats = list(exp_by_cat.index)
    s_labels = inc_cats + ['💰 ДОХОДЫ'] + exp_cats
    mid = len(inc_cats)
    s_src, s_tgt, s_val, s_col = [], [], [], []
    for i, cat in enumerate(inc_cats):
        v = round(float(inc_by_cat[cat]), 2)
        if v > 0:
            s_src.append(i) 
            s_tgt.append(mid) 
            s_val.append(v)
            s_col.append('rgba(16,185,129,0.35)')
    for j, cat in enumerate(exp_cats):
        v = round(float(exp_by_cat[cat]), 2)
        if v > 0:
            s_src.append(mid) 
            s_tgt.append(mid + 1 + j) 
            s_val.append(v)
            s_col.append('rgba(239,68,68,0.25)')
    A['sankey'] = {'labels': s_labels, 'src': s_src, 'tgt': s_tgt,
                   'val': s_val, 'colors': s_col}

    # ── Расходные/Доходные pie ──
    # Расходы — top-10 + остальное
    top_exp_cats = exp.groupby('category')['amount_pos'].sum().sort_values(ascending=False)
    if len(top_exp_cats) > 10:
        top10 = top_exp_cats.head(10)
        other = top_exp_cats.iloc[10:].sum()
        A['exp_pie_labels'] = list(top10.index) + ['Остальное']
        A['exp_pie_values'] = [round(float(v), 2) for v in top10.values] + [round(float(other), 2)]
    else:
        A['exp_pie_labels'] = list(top_exp_cats.index)
        A['exp_pie_values'] = [round(float(v), 2) for v in top_exp_cats.values]

    # Доходы
    A['inc_pie_labels'] = list(inc_by_cat.index)
    A['inc_pie_values'] = [round(float(v), 2) for v in inc_by_cat.values]

    # ── Прогноз ──
    last3 = months_order[-min(3, len(months_order)):]
    avg_i = float(monthly_inc[monthly_inc.index.isin(last3)].mean())
    avg_e = float(monthly_exp[monthly_exp.index.isin(last3)].mean())
    nd = (avg_i - avg_e) / 30
    lb = daily_bal[-1] if daily_bal else 0
    fc, fu, fl = [], [], []
    for i in range(31):
        f = lb + nd * i
        fc.append(round(f, 2))
        fu.append(round(f * 1.15 if f >= 0 else f * 0.85, 2))
        fl.append(round(f * 0.85 if f >= 0 else f * 1.15, 2))
    A['forecast']       = fc
    A['forecast_upper'] = fu
    A['forecast_lower'] = fl

    # ── Калькулятор — средние месячные по категориям ──
    calc_cats = {}
    for cat_name in A['cat_labels']:
        sub = exp[exp['category'] == cat_name]
        calc_cats[cat_name] = round(float(sub['amount_pos'].sum()) / n_months, 2)
    A['calc_cats'] = calc_cats

    # ── Топ-10 расходов ──
    top_exp = exp.nlargest(10, 'amount_pos')[['date_str','description','category','amount_pos']]
    A['top_expenses'] = [
        {'date': r['date_str'], 'desc': str(r['description'])[:60],
         'cat': r['category'], 'amount': round(float(r['amount_pos']), 2)}
        for _, r in top_exp.iterrows()
    ]

    # ── Инсайты ──
    A['insights'] = generate_insights(A)
    print(f"  ✓ Аналитика готова. Инсайтов: {len(A['insights'])}")
    return A


# ═══════════════════════════════════════════════════════
#  ИНСАЙТЫ (на русском)
# ════════════════════════════════════��══════════════════
def generate_insights(A: dict) -> list:
    ins = []
    n = A['months_count']

    if A['coverage_ratio'] < 1.2:
        ins.append({'type':'critical','icon':'⚠️',
            'title': f"Коэф. покрытия {A['coverage_ratio']} — ниже безопасного уровня",
            'text':  f"Расходы = {round(100/max(A['coverage_ratio'],0.01),1)}% дохода. Безопасный уровень ≥ 1.2.",
            'saving': f"Сократите расходы на {fmt_r(round(A['avg_monthly_expense']*0.15))} ₽/мес"})

    if A['save_rate'] < 15:
        ins.append({'type':'critical','icon':'💾',
            'title': f"Норма сбережения {A['save_rate']}% — критически низкая",
            'text':  "Финансовые советники рекомендуют ≥ 20%. Используйте правило «заплати себе первым».",
            'saving': f"Цель: {fmt_r(round(A['avg_monthly_income']*0.2))} ₽/мес"})

    if A['impulsive_pct'] > 15:
        ins.append({'type':'warning','icon':'⚡',
            'title': f"Импульсивные траты: {A['impulsive_pct']}% расходов",
            'text':  f"Импульсивные покупки составили {fmt_r(A['impulsive_sum'])} ₽. Используйте правило 24 часов.",
            'saving': f"–30% = экономия {fmt_r(round(A['impulsive_sum']*0.3/n))} ₽/мес"})

    if A['delivery_total'] > 0:
        avg_m = A['delivery_total'] / n
        ins.append({'type':'warning','icon':'🛵',
            'title': f"Доставка еды: {fmt_r(A['delivery_total'])} ₽ за период",
            'text':  f"В среднем {fmt_r(round(avg_m))} ₽/мес, средний заказ {fmt_r(A['delivery_avg'])} ₽.",
            'saving': f"Готовьте 3 раза/нед. дома: экономия ~{fmt_r(round(avg_m*0.4))} ₽/мес"})

    if A['taxi_total'] > 1000:
        ins.append({'type':'info','icon':'🚖',
            'title': f"Такси: {fmt_r(A['taxi_total'])} ₽ — возможна экономия",
            'text':  f"{A['taxi_count']} поездок. Замените 60% на метро (50 ₽/поездка) — экономия {fmt_r(round(A['taxi_total']*0.6))} ₽.",
            'saving': f"Экономия: {fmt_r(round(A['taxi_total']*0.6))} ₽"})

    ps = A['period_stats']
    n_pct = ps.get('night', {}).get('pct', 0)
    if n_pct > 8:
        ins.append({'type':'warning','icon':'🌙',
            'title': f"Ночные расходы (00-08ч): {n_pct}% от всех трат",
            'text':  f"Итого за ночь: {fmt_r(ps['night']['sum'])} ₽. Установите ночной лимит в приложении банка.",
            'saving': f"Экономия до {fmt_r(round(ps['night']['sum']*0.5/n))} ₽/мес"})

    if A['fines_total'] > 0:
        ins.append({'type':'critical','icon':'🚨',
            'title': f"Штрафы и комиссии: {fmt_r(A['fines_total'])} ₽",
            'text':  f"Обнаружено {len(A['fines'])} транзакций. Штрафы и пени — полностью предотвратимые расходы.",
            'saving': f"Избегайте: {fmt_r(A['fines_total'])} ₽ за период"})

    if A['leaky_days'] > 5:
        ins.append({'type':'warning','icon':'🕳',
            'title': f"Обнаружено {A['leaky_days']} «дырявых» дней",
            'text':  "Дни, когда расходы превысили дневной доход на 20%+.",
            'saving': "Контролируйте траты в пиковые дни"})

    return ins[:8]


# ═══════════════════════════════════════════════════════
#  ГЕНЕРАТОР HTML
# ═══════════════════════════════════════════════════════
def generate_html(A: dict) -> str:
    # Читаем шаблон
    template = TEMPLATE_FILE.read_text(encoding='utf-8')

    data_json = json.dumps(A, ensure_ascii=False, default=str)

    net_color  = '#10b981' if A['net_balance'] >= 0 else '#ef4444'
    net_sign   = '+' if A['net_balance'] >= 0 else ''
    save_color = '#10b981' if A['save_rate'] >= 20 else ('#f59e0b' if A['save_rate'] >= 10 else '#ef4444')
    save_status = 'Хорошо' if A['save_rate'] >= 20 else ('Ниже нормы' if A['save_rate'] >= 10 else 'Критично')
    cov_color  = '#10b981' if A['coverage_ratio'] >= 1.2 else '#f97316'
    cov_status = '✓ Безопасно ≥ 1.2' if A['coverage_ratio'] >= 1.2 else '⚠ Опасно < 1.2'
    n = A['months_count']

    replacements = {
        '%%DATA_JSON%%':           data_json,
        '%%DATE_FROM%%':           A['date_from'],
        '%%DATE_TO%%':             A['date_to'],
        '%%MONTHS_COUNT%%':        str(A['months_count']),
        '%%GEN_TIME%%':            datetime.now().strftime('%d.%m.%Y %H:%M'),
        '%%TOTAL_INCOME%%':        fmt_r(A['total_income']),
        '%%TOTAL_EXPENSE%%':       fmt_r(A['total_expense']),
        '%%AVG_MONTHLY_INCOME%%':  fmt_r(A['avg_monthly_income']),
        '%%AVG_MONTHLY_EXPENSE%%': fmt_r(A['avg_monthly_expense']),
        '%%NET_BALANCE%%':         fmt_r(abs(A['net_balance'])),
        '%%NET_SIGN%%':            net_sign,
        '%%NET_COLOR%%':           net_color,
        '%%NET_CARD%%':            'green' if A['net_balance'] >= 0 else 'red',
        '%%SAVE_RATE%%':           str(A['save_rate']),
        '%%SAVE_COLOR%%':          save_color,
        '%%SAVE_STATUS%%':         save_status,
        '%%SAVE_CARD%%':           'yellow' if A['save_rate'] < 20 else 'green',
        '%%COVERAGE_RATIO%%':      str(A['coverage_ratio']),
        '%%COV_COLOR%%':           cov_color,
        '%%COV_STATUS%%':          cov_status,
        '%%COV_CARD%%':            'orange' if A['coverage_ratio'] < 1.2 else 'green',
        '%%IMPULSIVE_PCT%%':       str(A['impulsive_pct']),
        '%%IMPULSIVE_SUM%%':       fmt_r(A['impulsive_sum']),
        '%%LEAKY_DAYS%%':          str(A['leaky_days']),
        '%%TAXI_TOTAL%%':          fmt_r(A['taxi_total']),
        '%%METRO_TOTAL%%':         fmt_r(A['metro_total']),
        '%%COFFEE_COUNT%%':        str(A['coffee_count']),
        '%%COFFEE_TOTAL%%':        fmt_r(A['coffee_total']),
        '%%COFFEE_PER_MONTH%%':    fmt_r(round(A['coffee_total'] / n, 2)),
        '%%COFFEE_AVG%%':          fmt_r(A['coffee_avg']),
        '%%FINES_TOTAL%%':         fmt_r(A['fines_total']),
        '%%FINES_COUNT%%':         str(len(A['fines'])),
        '%%DELIVERY_PER_MONTH%%':  fmt_r(round(A['delivery_total'] / n, 2)),
        '%%TAXI_PER_MONTH%%':      fmt_r(round(A['taxi_total'] / n, 2)),
    }

    html = template
    for ph, val in replacements.items():
        html = html.replace(ph, str(val))
    return html


# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  💰 ФИНАНСОВЫЙ ДАШБОРД — Генератор отчёта")
    print("=" * 60)

    if not EXCEL_FILE.exists():
        print(f"ОШИБКА: файл не найден: {EXCEL_FILE}")
        return
    if not TEMPLATE_FILE.exists():
        print(f"ОШИБКА: шаблон не найден: {TEMPLATE_FILE}")
        return

    try:
        df = load_excel(EXCEL_FILE)
        df = normalize_columns(df)
        df = parse_dates(df)
        df = parse_amounts(df)
        df = filter_internal_transfers(df)
        df = categorize(df)
        analytics = compute_analytics(df)

        print("\n📝 Генерирую HTML...")
        html = generate_html(analytics)
        OUTPUT_FILE.write_text(html, encoding='utf-8')
        kb = OUTPUT_FILE.stat().st_size / 1024
        print(f"✅ Сохранено: {OUTPUT_FILE}  ({kb:.1f} КБ)")
        webbrowser.open(OUTPUT_FILE.as_uri())
        print("🌐 Открываю в браузере...")
    except Exception as e:
        import traceback
        print(f"\n❌ ОШИБКА: {e}")
        traceback.print_exc()

    print("=" * 60)


if __name__ == "__main__":
    main()