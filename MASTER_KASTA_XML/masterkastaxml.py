import requests
import lxml.etree as ET
from datetime import datetime
import re
import time
from html import unescape
from collections import defaultdict, Counter
import openpyxl
import os

# ==============================================================================
# 1. КОНФІГУРАЦІЯ
# ==============================================================================
# (cat_prefix, id_prefix, url)
# cat_prefix  — додається до categoryId з фіду постачальника
# id_prefix   — додається до offer id (через "_"); "" = без префіксу
SOURCES = [
    ("1000", "1000",  "https://dropt.in.ua/index.php?route=export/prom&markup=15"),
    ("2222", "2222",  "https://opt-drop.com/storage/xml/opt-drop-1.xml"),
    ("1100", "1100",  "https://forus.com.ua/vugruzka/forus_opt_prom_stock.xml"),
    ("1200", "1200",  "https://aveon.net.ua/products_feed.xml?hash_tag=7b71fadcc4a12f03cf26a304da032fba&sales_notes=&product_ids=&label_ids=&exclude_fields=&html_description=0&yandex_cpa=&process_presence_sure=&languages=uk&group_ids="),
    ("1300", "1300",  "https://sonechko233.com.ua/products_feed.xml?hash_tag=220ed1761695cce1df21b74fc555efcd&sales_notes=&product_ids=&label_ids=&exclude_fields=&html_description=0&yandex_cpa=&process_presence_sure=&languages=uk%2Cru&extra_fields=&group_ids="),
]

OLD_PRICE_MULT      = 1.25     # old_price = price × 1.25 для всіх
MIN_PRICE_THRESHOLD = 199      # мінімальна ціна в грн
MIN_OFFERS_PER_CATEGORY = 3   # категорії з ≤ N прямих товарів видаляються (якщо не батьківські)
DESC_LIMIT          = 4500     # максимальна довжина опису для Kasta (ліміт Kasta до 5000)
DEFAULT_QTY         = 4        # кількість якщо постачальник не вказав або вказав 0
REQUEST_DELAY       = 6        # затримка між запитами в секундах (щоб не отримати 429)
MAX_FETCH_ATTEMPTS  = 5        # скільки разів пробувати завантажити фід
RETRY_BACKOFF_BASE  = 20       # базова пауза перед повтором (сек); далі росте 20→40→80…
RETRY_BACKOFF_MAX   = 120      # стеля паузи між повторами (сек)

# Наценка по доменах — ОБОВ'ЯЗКОВА для КОЖНОГО постачальника зі SOURCES.
CUSTOM_MARKUP = {
    "dropt.in.ua": {
        "markup_percent": 1.25,
        "markup_fixed":   50,
        "markup_tiers": [
            (500,    1.25,  40),   # до 500 грн
            (1000,   1.22,  35),   # 500–1000 грн
            (2000,   1.20,  30),   # 1000–2000 грн
            (4000,   1.20,  50),   # 2000–4000 грн
            (8000,   1.20,  50),   # 4000–8000 грн
            (999999, 1.20,  50),   # вище 8000 грн
        ],
    },
    "opt-drop.com": {
        "markup_percent": 1.35,
        "markup_fixed":   40,
        "markup_tiers": [
            (500,    1.35,  40),   # до 500 грн
            (1000,   1.35,  35),   # 500–1000 грн
            (2000,   1.32,  30),   # 1000–2000 грн
            (4000,   1.30,  50),   # 2000–4000 грн
            (8000,   1.30,  50),   # 4000–8000 грн
            (999999, 1.30,  50),   # вище 8000 грн
        ],
    },
    "forus.com.ua": {
        "markup_percent": 1.20,
        "markup_fixed":   40,
        "min_price_final": 130,   # мінімум від фінальної ціни (після наценки)
        "markup_tiers": [
            (500,    1.20,  50),   # до 500 грн
            (1000,   1.20,  35),   # 500–1000 грн
            (2000,   1.18,  30),   # 1000–2000 грн
            (4000,   1.18,  20),   # 2000–4000 грн
            (8000,   1.15,  40),   # 4000–8000 грн
            (999999, 1.15,  40),   # вище 8000 грн
        ],
    },
    "aveon.net.ua": {
        "markup_percent": 1.35,
        "markup_fixed":   40,
        "markup_tiers": [
            (500,    1.35,  40),   # до 500 грн
            (1000,   1.35,  35),   # 500–1000 грн
            (2000,   1.32,  30),   # 1000–2000 грн
            (4000,   1.32,  20),   # 2000–4000 грн
            (8000,   1.32,  40),   # 4000–8000 грн
            (999999, 1.32,  40),   # вище 8000 грн
        ],
    },
    "sonechko233.com.ua": {
        "markup_percent": 1.35,
        "markup_fixed":   40,
        "markup_tiers": [
            (500,    1.35,  40),   # до 500 грн
            (1000,   1.35,  35),   # 500–1000 грн
            (2000,   1.30,  30),   # 1000–2000 грн
            (4000,   1.30,  40),   # 2000–4000 грн
            (8000,   1.30,  40),   # 4000–8000 грн
            (999999, 1.30,  40),   # вище 8000 грн
        ],
    },
}

# Захист від підозрілих цін
MAX_PRICE_UAH      = 500_000
SUSPICIOUS_LOW_UAH = 10.0

# Запасні курси валют
FALLBACK_RATES = {
    "UAH": 1.0,
    "USD": 41.5,
    "EUR": 45.0,
    "RUB": 0.45,
    "RUR": 0.45,
    "BYN": 12.5,
    "PLN": 10.5,
    "GBP": 52.0,
}

# Словник для перекладу категорій Sonechko
SONECHKO_CAT_TRANSLATIONS = {
    "Товары для дома и сада": "Товари для дому та саду",
    "Сезонный товар": "Сезонний товар",
    "Красота и здоровье": "Краса та здоров'я",
    "PowerBank, внешние аккумуляторы": "PowerBank, зовнішні акумулятори",
    "Все для кухни": "Все для кухні",
    "Электроника": "Електроніка",
    "Игровые девайсы для ПК": "Ігрові девайси для ПК",
    "Одежда и обувь": "Одяг та взуття",
    "Охота и Рыбалка": "Полювання та риболовля",
    "Автотовары, электроинструмент, ручной инструмент": "Автотовари, електроінструмент, ручний інструмент",
    "Детский мир, детские товары": "Дитячий світ, дитячі товари",
    "Спорт, здоровье, туризм": "Спорт, здоров'я, туризм"
}

# Російсько-український переклад відомих кольорів для нормалізації
COLOR_RU_TO_UA = {
    'черный': 'Чорний',
    'черная': 'Чорний',
    'черное': 'Чорний',
    'белый': 'Білий',
    'белая': 'Білий',
    'белое': 'Білий',
    'красный': 'Червоний',
    'красная': 'Червоний',
    'красное': 'Червоний',
    'синий': 'Синій',
    'синяя': 'Синій',
    'синее': 'Синій',
    'зеленый': 'Зелений',
    'зеленая': 'Зелений',
    'зеленое': 'Зелений',
    'серый': 'Сірий',
    'серая': 'Сірий',
    'серое': 'Сірий',
    'розовый': 'Рожевий',
    'розовая': 'Рожевий',
    'розовое': 'Рожевий',
    'желтый': 'Жовтий',
    'желтая': 'Жовтий',
    'желтое': 'Жовтий',
    'голубой': 'Блакитний',
    'голубая': 'Блакитний',
    'голубое': 'Блакитний',
    'фиолетовый': 'Фіолетовий',
    'фиолетовая': 'Фіолетовий',
    'фиолетовое': 'Фіолетовий',
    'разные цвета': 'Комбінований',
    'разноцветный': 'Комбінований',
    'разноцветная': 'Комбінований',
    'разноцветные': 'Комбінований',
    'разноцветное': 'Комбінований',
    'комбинированный': 'Комбінований',
    'комбинированная': 'Комбінований',
    'комбинированные': 'Комбінований',
}

# ==============================================================================
# 2. ДОПОМІЖНІ ФУНКЦІЇ
# ==============================================================================

def fix_text(text):
    if not text:
        return ""
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', str(text))
    return unescape(unescape(text)).replace("\u2019", "'").strip()


_UA_CHARS = frozenset('\u0457\u0454\u0491\u0407\u0404\u0490')
_RU_CHARS = frozenset('\u044b\u044a\u044d\u042b\u042a\u042d')

def _lang(text):
    t = text or ''
    if any(c in _UA_CHARS for c in t): return 'uk'
    if any(c in _RU_CHARS for c in t): return 'ru'
    return 'other'


_RU_TO_UA = str.maketrans({'ы': 'и', 'Ы': 'И', 'э': 'е', 'Э': 'Е',
                            'ъ': '',  'Ъ': '',  'ё': 'е', 'Ё': 'Е'})

def ru_to_ua(text):
    if not text:
        return text
    return text.translate(_RU_TO_UA)


def clean_description(text, name_ua, vendor):
    fallback = f"<p>{name_ua} від виробника {vendor}.</p>".replace(']]>', ']] >')
    if not text:
        return fallback

    text = unescape(unescape(str(text)))
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)

    # Видаляємо небажані теги
    text = re.sub(r'<(script|style|iframe|video|audio)[^>]*>.*?</\1>', '', text, flags=re.DOTALL)
    text = re.sub(r'<(video|audio|iframe)[^>]*/>', '', text)
    text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)
    text = re.sub(r'<img[^>]*/?>', '', text)

    # Видаляємо URL
    text = re.sub(r'https?://\S+|www\.\S+', '', text)

    # Видаляємо inline стилі
    text = re.sub(r'\s+style="[^"]*"', '', text)
    text = re.sub(r"\s+style='[^']*'", '', text)

    # Видаляємо порожні теги
    text = re.sub(r'<(\w+)[^>]*>\s*</\1>', '', text)

    if len(text) > DESC_LIMIT:
        cut_pos = text.rfind('>', 0, DESC_LIMIT)
        if cut_pos > 0:
            text = text[:cut_pos + 1] + "..."
        else:
            text = text[:DESC_LIMIT] + "..."

    text = text.strip()

    # Перевірка мінімум 30 символів чистого тексту
    plain = re.sub(r'<[^>]+>', '', text).strip()
    if len(plain) < 30:
        return fallback

    text = text.replace(']]>', ']] >')
    return text


def parse_price(raw_text):
    if not raw_text:
        return None

    cleaned = str(raw_text).strip()
    cleaned = cleaned.replace('\xa0', '').replace('\u2009', '').replace('\u202f', '')
    cleaned = cleaned.replace(' ', '').replace('\t', '')

    if ',' in cleaned and '.' in cleaned:
        if cleaned.rfind(',') > cleaned.rfind('.'):
            cleaned = cleaned.replace('.', '').replace(',', '.')
        else:
            cleaned = cleaned.replace(',', '')
    elif ',' in cleaned:
        parts = cleaned.split(',')
        if len(parts) == 2 and len(parts[1]) <= 2:
            cleaned = cleaned.replace(',', '.')
        else:
            cleaned = cleaned.replace(',', '')

    try:
        result = float(cleaned)
        return result if result > 0 else None
    except (ValueError, TypeError):
        return None


def get_currency_rates(root):
    rates = dict(FALLBACK_RATES)
    for cur in root.xpath(".//currencies/currency"):
        cur_id   = (cur.get('id') or '').upper().strip()
        rate_str = cur.get('rate', '1')
        if not cur_id:
            continue
        if rate_str in ('CBR', 'НБУ', 'NBU', 'ECB', 'CB'):
            rates.setdefault(cur_id, FALLBACK_RATES.get(cur_id, 1.0))
        else:
            parsed = parse_price(rate_str)
            if parsed and parsed > 0:
                rates[cur_id] = parsed
    return rates


def convert_to_uah(raw_price, currency_id, rates, domain, offer_id):
    currency_id = (currency_id or 'UAH').upper().strip()
    warning = None

    if currency_id not in rates:
        warning = (f"[НЕВІДОМА ВАЛЮТА] {domain} offer={offer_id} "
                   f"currency={currency_id} — використовуємо UAH")
        currency_id = 'UAH'

    rate      = rates.get(currency_id, 1.0)
    price_uah = raw_price * rate

    if currency_id == 'UAH' and raw_price < SUSPICIOUS_LOW_UAH:
        warning = (f"[ПІДОЗРІЛА ЦІНА] {domain} offer={offer_id} "
                   f"price={raw_price} UAH < {SUSPICIOUS_LOW_UAH} грн — пропускаємо")
        return None, warning

    if currency_id != 'UAH' and raw_price > 500:
        warning = (f"[УВАГА ВАЛЮТА] {domain} offer={offer_id} "
                   f"price={raw_price} {currency_id} — конвертуємо: {price_uah:.2f} UAH")

    if price_uah > MAX_PRICE_UAH:
        warning = (f"[ЦІНА ЗАВИСОКА] {domain} offer={offer_id} "
                   f"raw={raw_price} {currency_id} → {price_uah:.2f} UAH > {MAX_PRICE_UAH} — пропускаємо")
        return None, warning

    if price_uah < SUSPICIOUS_LOW_UAH:
        warning = (f"[ЗАНИЗЬКА ПІСЛЯ КОНВЕРТАЦІЇ] {domain} offer={offer_id} "
                   f"raw={raw_price} {currency_id} → {price_uah:.2f} UAH — пропускаємо")
        return None, warning

    return price_uah, warning


def get_qty(offer):
    qty_nodes = offer.xpath(
        ".//quantity|.//quantity_in_stock|.//stock_quantity|.//amount"
    )
    if qty_nodes:
        node_text = (qty_nodes[0].text or '').strip()
        if node_text:
            try:
                qty = int(re.sub(r'\D', '', node_text))
                if qty > 0:
                    return qty, False
            except (ValueError, TypeError):
                pass

    outlets = offer.xpath(".//outlets")
    if outlets:
        try:
            qty = int(outlets[0].get('count', '0'))
            if qty > 0:
                return qty, False
        except (ValueError, TypeError):
            pass

    return DEFAULT_QTY, True


def get_availability(offer):
    AVAIL_TRUE = {'true', 'yes', '1'}

    avail_raw = offer.get('available', '').lower().strip()
    if avail_raw:
        return avail_raw in AVAIL_TRUE

    avail_tag = offer.findtext('available')
    if avail_tag is not None:
        return avail_tag.lower().strip() in AVAIL_TRUE

    in_stock = offer.get('in_stock', '').lower().strip()
    if in_stock:
        return in_stock in AVAIL_TRUE

    return False


def get_name(offer):
    name = fix_text(offer.findtext('name_ua') or '')
    if not name:
        name = fix_text(offer.findtext('name') or '')
    return name


def get_description(offer):
    desc = offer.findtext('description_ua') or ''
    if not desc or not desc.strip():
        desc = offer.findtext('description') or ''
    return desc


def get_params(offer):
    result = []
    for p in offer.findall('param'):
        val = fix_text(p.text)
        if not val:
            for v in p.findall('value'):
                lang = v.get('lang', '').lower()
                if lang in ('uk', 'ua'):
                    val = fix_text(v.text)
                    break
            if not val and p.findall('value'):
                val = fix_text(p.findall('value')[0].text)

        if not val or val in ('R R', 'r r'):
            continue

        name = (p.get('name') or '').strip()
        if name:
            result.append((name, val))

    return result


def get_article(offer):
    article = fix_text(
        offer.findtext('vendorCode') or
        offer.findtext('article')    or
        offer.findtext('vendor_code') or
        ''
    )
    return article[:255] if article else ''


def fetch_nbu_rates():
    try:
        r = requests.get(
            "https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?json",
            timeout=10
        )
        if not r.ok:
            return dict(FALLBACK_RATES)
        rates = {"UAH": 1.0}
        for item in r.json():
            code = item.get("cc", "").upper()
            rate = item.get("rate")
            if code and rate:
                rates[code] = float(rate)
        if "RUB" in rates:
            rates["RUR"] = rates["RUB"]
        print(f"[НБУ] Курси отримано: USD={rates.get('USD', '?'):.2f}, EUR={rates.get('EUR', '?'):.2f}")
        return rates
    except Exception as e:
        print(f"[НБУ] Помилка отримання курсів: {e} — використовуємо FALLBACK")
        return dict(FALLBACK_RATES)


def load_blacklist():
    try:
        path = "blacklist.txt"
        if not os.path.exists(path) and os.path.exists("../blacklist.txt"):
            path = "../blacklist.txt"
        with open(path, "r", encoding="utf-8") as f:
            ids = {
                line.strip()
                for line in f
                if line.strip() and not line.strip().startswith('#')
            }
        print(f"Blacklist завантажено: {len(ids)} товарів з {path}")
        return ids, len(ids)
    except FileNotFoundError:
        print("blacklist.txt не знайдено — крок пропускається")
        return set(), 0


def load_kasta_colors(filepath="кОЛЬОРИ КАСТА.xlsx"):
    allowed_colors = set()
    if not os.path.exists(filepath) and os.path.exists("../кОЛЬОРИ КАСТА.xlsx"):
        filepath = "../кОЛЬОРИ КАСТА.xlsx"
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True):
            val = row[0]
            if val:
                val_str = str(val).strip()
                if val_str:
                    allowed_colors.add(val_str)
        print(f"[КОЛЬОРИ] Завантажено {len(allowed_colors)} кольорів Kasta з {filepath}")
    except Exception as e:
        print(f"[КОЛЬОРИ] Помилка завантаження {filepath}: {e}. Використовуємо вбудований список.")
        allowed_colors = {
            'Лайм', 'Сірий', 'Синій', 'Бузковий', 'Сливовий', 'Темно-бежевий',
            'Темно-бірюзовий', 'Темно-бордовий', 'Темно-вишневий', 'Темно-блакитний',
            'Темно-зелений', 'Темно-золотистий', 'Темно-коричневий', 'Темно-червоний',
            'Темно-рожевий', 'Темно-сірий', 'Темно-синій', 'Темно-фіолетовий',
            'Теракотовий', 'Фіолетовий', 'Айворі', 'Темно-ліловий', 'Хакі', 'Чорний',
            'Фуксія', 'Кавовий', 'Сіро-голубий', 'Яскраво-червоний', 'Бежевий',
            'Білосніжний', 'Білий', 'Безбарвний', 'Бірюзовий', 'Бордовий', 'Бронзовий',
            'Блакитний', 'Гірчичний', 'Жовтий', 'Перловий', 'Зелений', 'Золотий',
            'Смарагдовий', 'Індиго', 'Кораловий', 'Коричневий', 'Червоний', 'Лавандовий',
            'Лососевий', 'Малиновий', 'Мідний', 'Молочний', 'Морської хвилі', "М'ятний",
            'Оливковий', 'Помаранчевий', 'Персиковий', 'Охра', 'Пурпурний', 'Пісочний',
            'Рожево-ліловий', 'Прозорий', 'Рожево-коричневий', 'Рожевий', 'Салатовий',
            'Світло-бірюзовий', 'Світло-вишневий', 'Світло-бордовий', 'Світло-жовтий',
            'Світло-коричневий', 'Світло-червоний', 'Світло-оранжевий', 'Світло-ліловий',
            'Світло-рожевий', 'Світло-сірий', 'Світло-синій', 'Світло-фіолетовий',
            'Срібний', 'Сіро-синій', 'Комбінований', 'Пудровий', 'Вишневий', 'Ліловий',
            'Світло-пурпурний', 'Темно-пурпурний', 'Фісташковий', 'Сіро-бежевий',
            'Сіро-коричневий', 'Кислотно-жовтий', 'Кислотно-рожевий', 'Кислотно-оранжевий',
            'Кислотно-зелений', 'Сіро-зелений', 'Чорно-білий', 'Сіро-червоний',
            'Синьо-жовтий', 'Метал', 'Графітовий', 'Нержавіюча сталь', 'Золотистий',
            'Койот', 'Світло-бежевий', 'Світло-зелений', 'Світло-блакитний', 'Кремовий',
            'Пляшковий зелений', 'Волошковий', 'Рудий', 'Бурштиновий', 'Блідо-рожевий',
            'Цегляний', 'Помаранчево-червоний', 'Сріблястий'
        }
    return allowed_colors


def normalize_color(color_val, allowed_colors):
    if not color_val:
        return "Комбінований"
    
    val = color_val.strip().lower()
    
    if val in COLOR_RU_TO_UA:
        return COLOR_RU_TO_UA[val]
    
    for allowed in allowed_colors:
        if allowed.lower() == val:
            return allowed
            
    translated = ru_to_ua(color_val).strip()
    trans_low = translated.lower()
    for allowed in allowed_colors:
        if allowed.lower() == trans_low:
            return allowed
            
    return "Комбінований"


def normalize_feed_tags(root):
    if root.xpath(".//offer") or not root.xpath(".//item"):
        return False

    for it in root.xpath(".//item"):
        it.tag = "offer"
    for im in root.xpath(".//image"):
        im.tag = "picture"
    for cat in root.xpath(".//category[@parentID]"):
        cat.set("parentId", cat.get("parentID"))
        del cat.attrib["parentID"]

    return True


def validate_markup_config():
    missing = []
    for _cat_prefix, _id_prefix, url in SOURCES:
        domain = url.split('/')[2]
        cfg = CUSTOM_MARKUP.get(domain)
        if not cfg or 'markup_percent' not in cfg or 'markup_fixed' not in cfg:
            missing.append(domain)
            continue
        tiers = cfg.get('markup_tiers')
        if tiers:
            for i, t in enumerate(tiers):
                if len(t) != 3:
                    raise SystemExit(
                        f"[КОНФІГ] markup_tiers[{i}] для '{domain}' має неправильний формат. "
                        f"Очікується кортеж (max_price, percent, fixed), отримано: {t}"
                    )
            prices = [t[0] for t in tiers]
            if prices != sorted(prices):
                raise SystemExit(
                    f"[КОНФІГ] markup_tiers для '{domain}' не відсортовані за зростанням ціни! "
                    f"Поточний порядок: {prices}"
                )
    if missing:
        raise SystemExit(
            "[КОНФІГ] Немає наценки в CUSTOM_MARKUP для: " + ", ".join(missing) +
            ". Додай markup_percent і markup_fixed для кожного з них."
        )
    print(f"[КОНФІГ] Наценку перевірено: усі {len(SOURCES)} постачальників мають явні значення")


def get_markup(price_uah, cfg):
    tiers = cfg.get('markup_tiers')
    if tiers:
        for max_price, pct, fixed in tiers:
            if price_uah <= max_price:
                return pct, fixed
        return tiers[-1][1], tiers[-1][2]
    return cfg['markup_percent'], cfg['markup_fixed']


# ==============================================================================
# 3. ГОЛОВНА ФУНКЦІЯ
# ==============================================================================

def process():
    final_categories = {}
    category_id_map  = {}
    price_warnings   = []
    source_results   = []

    report_stats     = {}
    cross_duplicates = []
    inner_duplicates = []
    blacklist_hits   = defaultdict(int)
    category_errors  = []

    print("--- СТАРТ ОБРОБКИ KASTA FEED ---")

    validate_markup_config()

    live_rates = fetch_nbu_rates()
    FALLBACK_RATES.update(live_rates)

    blacklisted_ids, blacklist_count = load_blacklist()
    allowed_colors = load_kasta_colors()

    feeds = []

    for i, (prefix, id_prefix, url) in enumerate(SOURCES):
        domain = url.split('/')[2]

        if i > 0:
            time.sleep(REQUEST_DELAY)

        last_error = None
        for attempt in range(1, MAX_FETCH_ATTEMPTS + 1):
            backoff = min(RETRY_BACKOFF_BASE * (2 ** (attempt - 1)), RETRY_BACKOFF_MAX)
            try:
                r = requests.get(url, timeout=120, headers={
                    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                                   'AppleWebKit/537.36 (KHTML, like Gecko) '
                                   'Chrome/124.0 Safari/537.36'),
                    'Accept': 'application/xml,text/xml,*/*;q=0.9',
                    'Accept-Language': 'uk,ru;q=0.9,en;q=0.8',
                })
                if not r.ok:
                    last_error = f"HTTP {r.status_code}"
                    if attempt < MAX_FETCH_ATTEMPTS:
                        wait = backoff
                        if r.status_code in (429, 503):
                            ra = r.headers.get('Retry-After', '').strip()
                            if ra.isdigit():
                                wait = max(int(ra), backoff)
                        print(f"[RETRY {attempt}/{MAX_FETCH_ATTEMPTS}] {domain}: {last_error} — повтор через {wait}с")
                        time.sleep(wait)
                        continue
                    print(f"[HTTP ERROR] {domain}: {r.status_code}")
                    report_stats[domain] = {"http_error": r.status_code}
                    break

                ct = r.headers.get('Content-Type', '')
                if 'html' in ct and 'xml' not in ct and len(r.content) < 50_000:
                    last_error = f"Content-Type={ct} (можливо HTML замість XML)"
                    if attempt < MAX_FETCH_ATTEMPTS:
                        print(f"[RETRY {attempt}/{MAX_FETCH_ATTEMPTS}] {domain}: {last_error} — повтор через {backoff}с")
                        time.sleep(backoff)
                        continue

                root = ET.fromstring(r.content, parser=ET.XMLParser(recover=True))
                if normalize_feed_tags(root):
                    print(f"[{domain}] Нестандартні теги (item/image/parentID) нормалізовано")
                currency_rates = get_currency_rates(root)
                visible_rates  = {k: v for k, v in currency_rates.items() if k in ('UAH', 'USD', 'EUR')}
                print(f"[{domain}] Завантажено (спроба {attempt}). Курси: {visible_rates}")
                feeds.append((prefix, id_prefix, url, domain, root, currency_rates))
                break

            except Exception as e:
                last_error = str(e)
                if attempt < MAX_FETCH_ATTEMPTS:
                    print(f"[RETRY {attempt}/{MAX_FETCH_ATTEMPTS}] {domain}: {e} — повтор через {backoff}с")
                    time.sleep(backoff)
                else:
                    print(f"[ПОМИЛКА ФІДУ] {domain}: {e}")
                    report_stats[domain] = {"feed_error": last_error}

    id_registry = defaultdict(list)

    for prefix, id_prefix, url, domain, root, currency_rates in feeds:
        for offer in root.xpath(".//offer"):
            raw_id   = offer.get('id', '').strip().upper()
            if not raw_id:
                continue
            offer_id  = f"{id_prefix}_{raw_id}" if id_prefix else raw_id
            price_nodes = offer.xpath('./price')
            price_text  = price_nodes[0].text if price_nodes else ''
            id_registry[offer_id].append((domain, price_text or ''))

    conflict_ids = set()
    for offer_id, entries in id_registry.items():
        if len(entries) > 1:
            domains = [e[0] for e in entries]
            if len(set(domains)) == 1:
                inner_duplicates.append({
                    "offer_id": offer_id,
                    "domain":   domains[0],
                    "count":    len(entries)
                })
            else:
                cross_duplicates.append({
                    "offer_id": offer_id,
                    "entries":  entries
                })
            conflict_ids.add(offer_id)

    print(f"\nДублікати між постачальниками: {len(cross_duplicates)}")
    print(f"Дублікати всередині постачальника: {len(inner_duplicates)}")

    for prefix, id_prefix, url, domain, root, currency_rates in feeds:
        for cat in root.xpath(".//category"):
            orig_id = cat.get('id')
            if not orig_id:
                continue
            new_id = f"{prefix}{orig_id}" if prefix else orig_id

            while new_id in category_id_map and category_id_map[new_id] != domain:
                new_id = f"{new_id}9"

            category_id_map[new_id] = domain
            cat.set('id', new_id)

            parent = cat.get('parentId') or cat.get('parent_id') or cat.get('parentID')
            if parent:
                cat.set('parentId', f"{prefix}{parent}" if prefix else parent)
                if 'parent_id' in cat.attrib:
                    del cat.attrib['parent_id']
                if 'parentID' in cat.attrib:
                    del cat.attrib['parentID']

            if cat.text:
                cat_text_clean = " ".join(fix_text(cat.text).split())
                if domain == "sonechko233.com.ua" and cat_text_clean in SONECHKO_CAT_TRANSLATIONS:
                    cat.text = SONECHKO_CAT_TRANSLATIONS[cat_text_clean]
                else:
                    cat.text = cat_text_clean

            final_categories[new_id] = cat

    processed_offers = []

    for prefix, id_prefix, url, domain, root, currency_rates in feeds:
        count_ok          = 0
        count_low         = 0
        count_no          = 0
        count_price_err   = 0
        count_duplicate   = 0
        count_blacklist   = 0
        count_default_qty = 0
        count_name_ua     = 0
        count_name_ru     = 0
        count_desc_ua     = 0
        count_desc_ru     = 0
        count_desc_none   = 0
        count_multi_pic   = 0
        count_no_params   = 0
        count_no_article  = 0
        price_min         = float('inf')
        price_max         = 0.0

        for offer in root.xpath(".//offer"):
            raw_id   = offer.get('id', '').strip().upper()
            if not raw_id:
                continue
            offer_id  = f"{id_prefix}_{raw_id}" if id_prefix else raw_id

            if offer_id in blacklisted_ids:
                blacklist_hits[domain] += 1
                count_blacklist += 1
                continue

            if offer_id in conflict_ids:
                count_duplicate += 1
                continue

            if not get_availability(offer):
                count_no += 1
                continue

            qty, used_default = get_qty(offer)
            if used_default:
                count_default_qty += 1

            price_nodes = offer.xpath('./price')
            if not price_nodes or not (price_nodes[0].text or '').strip():
                count_price_err += 1
                continue
            p_node = price_nodes[0]

            if p_node.get('from', 'false').lower() == 'true':
                price_warnings.append(
                    f"[ЦІНА З ДІАПАЗОНУ] {domain} offer={offer_id} "
                    f"price='{p_node.text}' — мінімальна ціна з діапазону"
                )

            try:
                raw_p = parse_price(p_node.text)
                if raw_p is None:
                    price_warnings.append(
                        f"[НЕМОЖЛИВО РОЗПАРСИТИ] {domain} offer={offer_id} "
                        f"raw='{p_node.text}'"
                    )
                    count_price_err += 1
                    continue

                currency_id     = (offer.findtext('currencyId') or 'UAH').strip().upper()
                price_uah, warn = convert_to_uah(raw_p, currency_id, currency_rates, domain, offer_id)

                if warn:
                    price_warnings.append(warn)
                if price_uah is None:
                    count_price_err += 1
                    continue

                cfg                = CUSTOM_MARKUP[domain]
                m_percent, m_fixed = get_markup(price_uah, cfg)

                price     = round(price_uah * m_percent + m_fixed)
                old_price = round(price * OLD_PRICE_MULT)

                min_raw   = cfg.get("min_price_raw")
                min_final = cfg.get("min_price_final")
                if min_raw is not None:
                    if price_uah < min_raw:
                        count_low += 1
                        continue
                elif min_final is not None:
                    if price < min_final:
                        count_low += 1
                        continue
                elif price < MIN_PRICE_THRESHOLD:
                    count_low += 1
                    continue

                if price < price_uah - 1:
                    price_warnings.append(
                        f"[ЦІНА НИЖЧА ЗА ОРИГІНАЛ] {domain} offer={offer_id} "
                        f"original={price_uah:.0f} UAH our_price={price} UAH — видаляємо"
                    )
                    count_price_err += 1
                    continue

                vendor  = fix_text(offer.findtext('vendor') or '') or 'NoBrand'
                name_ua = ru_to_ua(get_name(offer))

                if not name_ua or len(name_ua.strip()) < 3:
                    count_price_err += 1
                    continue

                if vendor != 'NoBrand' and vendor.lower() not in name_ua.lower():
                    name_ua = f"{name_ua} {vendor}"

                desc_raw = get_description(offer)
                desc     = ru_to_ua(clean_description(desc_raw, name_ua, vendor))

                orig_cat = offer.findtext('categoryId') or ''
                cat_id   = f"{prefix}{orig_cat}" if prefix else orig_cat

                new_off = ET.Element("offer", id=offer_id, available="true")

                ET.SubElement(new_off, "price").text          = str(price)
                ET.SubElement(new_off, "price_old").text      = str(old_price)
                ET.SubElement(new_off, "stock_quantity").text = str(min(qty, 9999))
                ET.SubElement(new_off, "currencyId").text     = "UAH"
                ET.SubElement(new_off, "categoryId").text     = cat_id

                pic_count = 0
                seen_pics = set()
                for pic in offer.findall('picture'):
                    if pic_count >= 15:
                        break
                    url_val = (pic.text or '').strip()
                    if url_val and url_val.startswith(('http://', 'https://')) and url_val not in seen_pics:
                        ET.SubElement(new_off, "picture").text = url_val
                        seen_pics.add(url_val)
                        pic_count += 1
                if pic_count == 0:
                    price_warnings.append(
                        f"[БЕЗ ФОТО] {domain} offer={offer_id} — пропускаємо"
                    )
                    count_price_err += 1
                    continue

                ET.SubElement(new_off, "vendor").text         = vendor

                article = get_article(offer)
                if article:
                    ET.SubElement(new_off, "article").text    = article

                ET.SubElement(new_off, "name_ua").text        = name_ua[:250]
                ET.SubElement(new_off, "description_ua").text = ET.CDATA(desc)

                params = get_params(offer)
                normalized_params = []
                has_color = False
                has_size = False

                for p_name, p_val in params:
                    p_name_low = p_name.lower().strip()
                    if p_name_low in ("колір", "цвет", "color"):
                        if not has_color:
                            normalized_color = normalize_color(p_val, allowed_colors)
                            normalized_params.append(("Колір", normalized_color))
                            has_color = True
                    elif any(w in p_name_low for w in ('розмір', 'размер', 'size', 'габарит')):
                        normalized_params.append((p_name, p_val))
                        has_size = True
                    else:
                        normalized_params.append((p_name, p_val))

                if not has_color:
                    normalized_params.append(("Колір", "Комбінований"))

                if not has_size:
                    normalized_params.append(("Розмір Size", "-"))

                for p_name, p_val in normalized_params:
                    ET.SubElement(new_off, "param", name=p_name[:100]).text = p_val[:255]

                _n_ua = (offer.findtext('name_ua') or '').strip()
                _n    = (offer.findtext('name')    or '').strip()
                if _n_ua or _lang(_n_ua or _n) == 'uk':
                    count_name_ua += 1
                elif _lang(_n) == 'ru':
                    count_name_ru += 1

                _d_ua = (offer.findtext('description_ua') or '').strip()
                _d    = (offer.findtext('description')    or '').strip()
                if not _d_ua and not _d:
                    count_desc_none += 1
                elif _d_ua or _lang(_d_ua or _d) == 'uk':
                    count_desc_ua += 1
                elif _lang(_d) == 'ru':
                    count_desc_ru += 1

                if pic_count >= 2:  count_multi_pic  += 1
                if not params:      count_no_params  += 1
                if not article:     count_no_article += 1
                if price < price_min: price_min = price
                if price > price_max: price_max = price

                processed_offers.append(new_off)
                count_ok += 1

            except Exception as e:
                price_warnings.append(
                    f"[ВИНЯТОК] {domain} offer={offer_id} "
                    f"price='{p_node.text if p_node is not None else 'N/A'}' err={e}"
                )
                count_price_err += 1
                continue

        report_stats[domain] = {
            "ok":           count_ok,
            "low":          count_low,
            "not_avail":    count_no,
            "price_err":    count_price_err,
            "duplicate":    count_duplicate,
            "blacklist":    count_blacklist,
            "default_qty":  count_default_qty,
            "name_ua":      count_name_ua,
            "name_ru":      count_name_ru,
            "desc_ua":      count_desc_ua,
            "desc_ru":      count_desc_ru,
            "desc_none":    count_desc_none,
            "multi_pic":    count_multi_pic,
            "no_params":    count_no_params,
            "no_article":   count_no_article,
            "price_min":    int(price_min) if price_min != float('inf') else 0,
            "price_max":    int(price_max),
        }
        source_results.append(
            f"{domain}: OK={count_ok} | LOW={count_low} | NOT_AVAIL={count_no} | "
            f"PRICE_ERR={count_price_err} | DUPL={count_duplicate} | "
            f"BLACKLIST={count_blacklist} | DEFAULT_QTY={count_default_qty}"
        )

    valid_offers = []
    for offer in processed_offers:
        cat_id   = offer.findtext('categoryId')
        offer_id = offer.get('id', 'unknown')
        if cat_id and cat_id in final_categories:
            valid_offers.append(offer)
        else:
            category_errors.append(
                f"offer={offer_id} categoryId={cat_id} — категорія не знайдена, товар видалено"
            )

    if category_errors:
        print(f"\n[УВАГА] Видалено товарів через відсутню категорію: {len(category_errors)}")

    parent_ids = {
        cat_el.get('parentId')
        for cat_el in final_categories.values()
        if cat_el.get('parentId') in final_categories
    }

    cat_direct_counts = Counter(
        offer.findtext('categoryId') for offer in valid_offers
    )

    thin_cats = {
        cat_id
        for cat_id in final_categories
        if cat_direct_counts.get(cat_id, 0) <= MIN_OFFERS_PER_CATEGORY
        and cat_id not in parent_ids
    }

    if thin_cats:
        new_valid = []
        thin_offers_removed = 0
        for offer in valid_offers:
            if offer.findtext('categoryId') in thin_cats:
                thin_offers_removed += 1
            else:
                new_valid.append(offer)
        valid_offers = new_valid

        for cat_id in thin_cats:
            del final_categories[cat_id]

        empty_removed = sum(1 for c in thin_cats if cat_direct_counts.get(c, 0) == 0)
        thin_removed = len(thin_cats) - empty_removed
        print(
            f"\n[КАТЕГОРІЇ] Видалено {len(thin_cats)} категорій "
            f"({empty_removed} порожніх + {thin_removed} з <= {MIN_OFFERS_PER_CATEGORY} товарів), "
            f"{thin_offers_removed} товарів видалено"
        )
    else:
        print(f"\n[КАТЕГОРІЇ] Тонких/порожніх категорій не знайдено")

    yml  = ET.Element("yml_catalog", date=datetime.now().strftime("%Y-%m-%d %H:%M"))
    shop = ET.SubElement(yml, "shop")
    ET.SubElement(shop, "name").text = "AVI KASTA"
    ET.SubElement(shop, "url").text  = "https://avi.in.ua"

    currencies = ET.SubElement(shop, "currencies")
    ET.SubElement(currencies, "currency", id="UAH", rate="1")

    cats_n = ET.SubElement(shop, "categories")
    for c in final_categories.values():
        cats_n.append(c)

    offers_n = ET.SubElement(shop, "offers")
    for o in valid_offers:
        offers_n.append(o)

    output_filename = "Masterkastanew.xml"
    with open(output_filename, "wb") as f:
        xml_bytes = ET.tostring(yml, encoding='UTF-8', xml_declaration=True, pretty_print=True)
        xml_bytes = xml_bytes.replace(
            b"<?xml version='1.0' encoding='UTF-8'?>\n",
            b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE yml_catalog SYSTEM "shops.dtd">\n',
            1
        )
        xml_bytes = xml_bytes.replace(
            b'<?xml version="1.0" encoding="UTF-8"?>\n<?xml',
            b'<?xml'
        )
        if b'<!DOCTYPE' not in xml_bytes:
            xml_bytes = xml_bytes.replace(
                b'<?xml version="1.0" encoding="UTF-8"?>\n',
                b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE yml_catalog SYSTEM "shops.dtd">\n',
                1
            )
        f.write(xml_bytes)
    print(f"[XML] Сгенеровано файл {output_filename}")

    with open("kasta_price_warnings.log", "w", encoding="utf-8") as f:
        f.write('\n'.join(price_warnings))

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    has_critical = any("http_error" in v or "feed_error" in v for v in report_stats.values())
    has_warnings = (len(price_warnings) > 0 or
                    len(cross_duplicates) > 0 or
                    len(inner_duplicates) > 0 or
                    len(category_errors) > 0)

    if has_critical:
        status = "🔴 КРИТИЧНО"
    elif has_warnings:
        status = "🟡 УВАГА"
    else:
        status = "🟢 ОК"

    feeds_ok    = sum(1 for v in report_stats.values() if "ok" in v)
    feeds_total = len(SOURCES)

    md = []
    md.append("# MASTERKASTANEW — Звіт запуску")
    md.append(f"**Дата:** {now_str}  ")
    md.append(f"**Статус:** {status}\n")

    md.append("## Загальний підсумок")
    md.append("| Показник | Значення |")
    md.append("|---|---|")
    md.append(f"| Всього товарів у прайсі | {len(valid_offers):,} |")
    md.append(f"| Постачальників оброблено | {feeds_ok}/{feeds_total} |")
    md.append(f"| Видалено через дублі (між постачальниками) | {len(cross_duplicates)} |")
    md.append(f"| Видалено через дублі (всередині постачальника) | {len(inner_duplicates)} |")
    md.append(f"| Видалено через blacklist | {sum(blacklist_hits.values())} |")
    md.append(f"| Видалено через відсутню категорію | {len(category_errors)} |")
    md.append(f"| Попереджень по цінах | {len(price_warnings)} |\n")

    md.append("## По постачальниках")
    md.append("| Постачальник | ✅ OK | 💰 Низька ціна | 🚫 Недоступні | ⚠️ Помилки ціни | 📦 Сток за замовч. | 🔁 Дублі | 🚷 Blacklist |")
    md.append("|---|---|---|---|---|---|---|---|")
    for prefix, id_prefix, url in SOURCES:
        domain = url.split('/')[2]
        v = report_stats.get(domain, {})
        if "http_error" in v:
            md.append(f"| {domain} | 🔴 HTTP {v['http_error']} | — | — | — | — | — | — |")
        elif "feed_error" in v:
            md.append(f"| {domain} | 🔴 ПОМИЛКА ЗАВАНТАЖЕННЯ | — | — | — | — | — | — |")
        else:
            md.append(
                f"| {domain} "
                f"| {v.get('ok', 0)} "
                f"| {v.get('low', 0)} "
                f"| {v.get('not_avail', 0)} "
                f"| {v.get('price_err', 0)} "
                f"| {v.get('default_qty', 0)} "
                f"| {v.get('duplicate', 0)} "
                f"| {v.get('blacklist', 0)} |"
            )

    md.append("\n## Якість даних по постачальниках")
    md.append("| Постачальник | 🇺🇦 Назва UA | 🇷🇺 Назва RU | 🇺🇦 Опис UA | 🇷🇺 Опис RU | ❌ Без опису | 📸 2+ фото | ⚙️ Без парамів | 🏷️ Без артикула | 💰 Ціна min–max |")
    md.append("|---|---|---|---|---|---|---|---|---|---|")
    for prefix, id_prefix, url in SOURCES:
        domain = url.split('/')[2]
        v = report_stats.get(domain, {})
        if "http_error" in v or "feed_error" in v:
            md.append(f"| {domain} | — | — | — | — | — | — | — | — | — |")
        else:
            p_min = v.get('price_min', 0)
            p_max = v.get('price_max', 0)
            price_range = f"{p_min}–{p_max} грн" if p_max > 0 else "—"
            md.append(
                f"| {domain} "
                f"| {v.get('name_ua', 0)} "
                f"| {v.get('name_ru', 0)} "
                f"| {v.get('desc_ua', 0)} "
                f"| {v.get('desc_ru', 0)} "
                f"| {v.get('desc_none', 0)} "
                f"| {v.get('multi_pic', 0)} "
                f"| {v.get('no_params', 0)} "
                f"| {v.get('no_article', 0)} "
                f"| {price_range} |"
            )

    if cross_duplicates:
        md.append("\n## ⚠️ Дублікати між постачальниками (видалено з прайсу)")
        md.append("| offer id | Постачальник 1 | Ціна 1 | Постачальник 2 | Ціна 2 |")
        md.append("|---|---|---|---|---|")
        for d in cross_duplicates[:50]:
            entries = d["entries"]
            e1 = entries[0] if len(entries) > 0 else ("—", "—")
            e2 = entries[1] if len(entries) > 1 else ("—", "—")
            md.append(f"| {d['offer_id']} | {e1[0]} | {e1[1]} | {e2[0]} | {e2[1]} |")
        if len(cross_duplicates) > 50:
            md.append(f"\n*... і ще {len(cross_duplicates) - 50} дублікатів*")

    if inner_duplicates:
        md.append("\n## ⚠️ Дублікати всередині постачальника (видалено з прайсу)")
        md.append("| offer id | Постачальник | Кількість входжень |")
        md.append("|---|---|---|")
        for d in inner_duplicates[:50]:
            md.append(f"| {d['offer_id']} | {d['domain']} | {d['count']} |")
        if len(inner_duplicates) > 50:
            md.append(f"\n*... і ще {len(inner_duplicates) - 50} дублікатів*")

    if price_warnings:
        md.append("\n## ⚠️ Попередження по цінах")
        md.append("```")
        for w in price_warnings[:50]:
            md.append(w)
        if len(price_warnings) > 50:
            md.append(f"... і ще {len(price_warnings) - 50} попереджень (див. kasta_price_warnings.log)")
        md.append("```")

    if category_errors:
        md.append("\n## ⚠️ Товари видалені через відсутню категорію")
        md.append("```")
        for e in category_errors[:20]:
            md.append(e)
        if len(category_errors) > 20:
            md.append(f"... і ще {len(category_errors) - 20}")
        md.append("```")

    md.append("\n## 🚷 Blacklist")
    if blacklist_count == 0:
        md.append("blacklist.txt не знайдено або порожній — крок пропущено")
    else:
        md.append(f"Файл blacklist.txt завантажено: **{blacklist_count}** id у списку  ")
        md.append(f"Видалено товарів з прайсу: **{sum(blacklist_hits.values())}**")

    md.append(f"\n---")
    md.append(f"*Звіт сформовано автоматично: {now_str}*")

    with open("KASTA_REPORT.md", "w", encoding="utf-8") as f:
        f.write('\n'.join(md))

    print("\n=== ПІДСУМОК ПО ДЖЕРЕЛАХ ===")
    for s in source_results:
        print(f"  {s}")

    try:
        print(f"\n  Статус: {status}")
    except UnicodeEncodeError:
        status_safe = status.replace("🔴", "КРИТИЧНО").replace("🟡", "УВАГА").replace("🟢", "ОК")
        print(f"\n  Статус: {status_safe}")
    print(f"  Всього товарів у прайсі: {len(valid_offers):,}")
    print(f"  Звіт збережено у KASTA_REPORT.md")
    print("--- ГОТОВО ---")


if __name__ == "__main__":
    process()
