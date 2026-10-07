"""myshops.az məhsul adlarının yoxlanması.

İstifadə:
    python myshops_name_check.py                     # https://myshops.az/products.xml yüklənir
    python myshops_name_check.py products.xml        # lokal fayl
    python myshops_name_check.py products.xml -o netice.xlsx

Nəticə: Excel faylı, iki vərəq:
    "Səhv adlar"      - problemi olan məhsullar və təklif olunan düzgün ad
    "Bütün məhsullar" - feed-dəki bütün məhsullar
"""

import argparse
import re
import sys
import urllib.request

from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

FEED_URL = "https://myshops.az/products.xml"

# Xarici (ingilis/rus) növ sözü -> Azərbaycan dilində qarşılığı.
# Çoxsözlü ifadələr birsözlülərdən əvvəl yoxlanır (uzunluğa görə sıralanır).
FOREIGN_TO_AZ = {
    # Şəxsi qulluq
    "trimmer": "Üzqırxan",
    "shaver": "Üzqırxan",
    "electric shaver": "Üzqırxan",
    "razor": "Üzqırxan",
    "hair trimmer": "Saç maşını",
    "hair clipper": "Saç maşını",
    "clipper": "Saç maşını",
    "epilator": "Epilyator",
    "hair dryer": "Fen",
    "hairdryer": "Fen",
    "dryer": "Fen",
    "hair straightener": "Saç düzləndirici",
    "straightener": "Saç düzləndirici",
    "curling iron": "Saç burğusu",
    "curler": "Saç burğusu",
    "hair styler": "Saç stayleri",
    "styler": "Saç stayleri",
    "toothbrush": "Diş fırçası",
    "electric toothbrush": "Elektrik diş fırçası",
    "scale": "Tərəzi",
    "bathroom scale": "Döşəmə tərəzisi",
    "kitchen scale": "Mətbəx tərəzisi",
    # Mətbəx
    "blender": "Blender",  # AZ-da da "Blender" işlənir - aşağıda AZ siyahısında da var
    "mixer": "Mikser",
    "hand mixer": "Mikser",
    "kettle": "Elektrik çaydanı",
    "electric kettle": "Elektrik çaydanı",
    "toaster": "Toster",
    "coffee machine": "Qəhvə maşını",
    "coffee maker": "Qəhvə maşını",
    "coffee grinder": "Qəhvə üyüdən",
    "grinder": "Üyüdən",
    "juicer": "Şirəçəkən",
    "meat grinder": "Ət maşını",
    "food processor": "Mətbəx kombaynı",
    "multicooker": "Multibişirici",
    "air fryer": "Hava fritözü",
    "airfryer": "Hava fritözü",
    "fryer": "Fritöz",
    "microwave": "Mikrodalğalı soba",
    "microwave oven": "Mikrodalğalı soba",
    "oven": "Soba",
    "cooker": "Plita",
    "hob": "Bişirmə paneli",
    "refrigerator": "Soyuducu",
    "fridge": "Soyuducu",
    "freezer": "Dondurucu",
    "dishwasher": "Qabyuyan maşın",
    "hood": "Aspirator",
    "cooker hood": "Aspirator",
    "sandwich maker": "Sendviç bişirən",
    "waffle maker": "Vafli bişirən",
    "grill": "Qril",
    "steamer": "Buxarda bişirən",
    # Ev texnikası
    "washing machine": "Paltaryuyan maşın",
    "dryer machine": "Paltarqurudan maşın",
    "vacuum cleaner": "Tozsoran",
    "vacuum": "Tozsoran",
    "robot vacuum": "Robot tozsoran",
    "iron": "Ütü",
    "steam iron": "Ütü",
    "garment steamer": "Buxar ütüsü",
    "air conditioner": "Kondisioner",
    "conditioner": "Kondisioner",
    "heater": "Qızdırıcı",
    "fan": "Ventilyator",
    "humidifier": "Nəmləndirici",
    "air purifier": "Hava təmizləyici",
    "water heater": "Su qızdırıcısı",
    # Elektronika
    "smartphone": "Smartfon",
    "mobile phone": "Mobil telefon",
    "phone": "Telefon",
    "tablet": "Planşet",
    "laptop": "Noutbuk",
    "notebook": "Noutbuk",
    "monitor": "Monitor",
    "tv": "Televizor",
    "television": "Televizor",
    "smart tv": "Televizor",
    "headphones": "Qulaqlıq",
    "headphone": "Qulaqlıq",
    "headset": "Qulaqlıq",
    "earbuds": "Qulaqlıq",
    "earphones": "Qulaqlıq",
    "speaker": "Səsgücləndirici",
    "portable speaker": "Portativ dinamik",
    "smart watch": "Smart saat",
    "smartwatch": "Smart saat",
    "watch": "Saat",
    "fitness tracker": "Fitnes qolbaq",
    "smart band": "Fitnes qolbaq",
    "power bank": "Powerbank",
    "charger": "Adapter",
    "cable": "Kabel",
    "keyboard": "Klaviatura",
    "mouse": "Siçan",
    "printer": "Printer",
    "router": "Router",
    "camera": "Kamera",
    "game console": "Oyun konsolu",
    "console": "Oyun konsolu",
    "case": "Örtük",
    "cover": "Örtük",
    "screen protector": "Qoruyucu şüşə",
    "protective glass": "Qoruyucu şüşə",
    # Rus dilində ən çox rast gəlinənlər
    "триммер": "Üzqırxan",
    "бритва": "Üzqırxan",
    "фен": "Fen",
    "утюг": "Ütü",
    "пылесос": "Tozsoran",
    "чайник": "Elektrik çaydanı",
    "холодильник": "Soyuducu",
    "телефон": "Telefon",
    "смартфон": "Smartfon",
    "ноутбук": "Noutbuk",
    "наушники": "Qulaqlıq",
    "телевизор": "Televizor",
}

# Azərbaycan dilində düzgün növ sözləri (adın əvvəlində olmalıdır).
AZ_TYPE_WORDS = {
    "üzqırxan", "ülgüc", "saç", "saqqal", "epilyator", "fen", "diş", "tərəzi",
    "döşəmə", "mətbəx", "blender", "mikser", "çaydan", "elektrik", "toster",
    "qəhvə", "qəhvəbişirən", "üyüdən", "şirəçəkən", "ət", "multibişirici",
    "multivarka", "hava", "fritöz", "mikrodalğalı", "soba", "plita", "bişirmə",
    "soyuducu", "dondurucu", "qabyuyan", "aspirator", "sendviç", "vafli", "qril",
    "buxarda", "paltaryuyan", "paltarqurudan", "tozsoran", "robot", "ütü",
    "buxar", "kondisioner", "qızdırıcı", "radiator", "ventilyator",
    "nəmləndirici", "su", "smartfon", "mobil", "telefon", "planşet", "noutbuk",
    "monitor", "televizor", "qulaqlıq", "səsgücləndirici", "dinamik",
    "akustik", "smart", "saat", "fitnes", "qolbaq", "powerbank", "adapter",
    "şarj", "kabel", "klaviatura", "siçan", "printer", "router", "kamera",
    "oyun", "örtük", "qoruyucu", "çanta", "kompüter", "proyektor", "tikiş",
    "massaj", "masajor", "manikür", "pedikür", "termos", "qazan", "tava",
    "dəst", "bıçaq", "mikrofon", "naviqator", "videoreqistrator", "batareya",
    "lampa", "işıq", "fənər", "stol", "stul", "kreslo", "yaddaş", "fləş",
    "sərt", "ssd", "operativ", "videokart", "prosessor", "qurğu", "pult",
    "ayaqqabı", "köynək", "şalvar", "uşaq", "oyuncaq", "ətir", "krem",
    "şampun", "sabun", "gel", "losyon", "parfüm", "eynək", "pul", "kisə",
    "qab", "fincan", "stəkan", "boşqab", "çəngəl", "qaşıq", "nimçə",
    "elektrikli", "portativ", "simsiz", "bluetooth", "multimedia",
}

# Bəzi geniş yayılmış brendlər - ad birbaşa brendlə başlayırsa, növ sözü yoxdur.
KNOWN_BRANDS = {
    "philips", "braun", "remington", "panasonic", "rowenta", "tefal", "moulinex",
    "bosch", "siemens", "samsung", "lg", "apple", "iphone", "ipad", "macbook",
    "xiaomi", "redmi", "poco", "huawei", "honor", "oppo", "vivo", "realme",
    "oneplus", "nokia", "motorola", "sony", "jbl", "marshall", "beats", "bose",
    "lenovo", "hp", "dell", "asus", "acer", "msi", "dyson", "kenwood",
    "delonghi", "de'longhi", "beko", "arcelik", "indesit", "ariston",
    "electrolux", "zanussi", "gorenje", "hisense", "tcl", "sharp", "toshiba",
    "hitachi", "midea", "haier", "vitek", "polaris", "scarlett", "redmond",
    "kitfort", "babyliss", "valera", "gamma", "ga.ma", "oral-b", "garmin",
    "amazfit", "fitbit", "anker", "baseus", "ugreen", "logitech", "razer",
    "canon", "nikon", "epson", "tp-link", "nintendo", "playstation", "xbox",
    "tecno", "infinix", "itel", "zte", "blackview", "ulefone", "nothing",
    "google", "pixel", "karcher", "kärcher", "rozenberg", "fakir", "arzum",
    "sinbo", "vestel", "grundig", "smeg", "nespresso", "krups", "russell",
}

NAME_TAGS = ("name", "title", "model", "product_name", "productname")
FIELD_TAGS = {
    "id": ("id", "sku", "code", "product_id", "item_group_id"),
    "url": ("url", "link", "product_url"),
    "price": ("price", "sale_price"),
    "brand": ("brand", "vendor", "manufacturer"),
    "category": ("category", "category_name", "product_type", "google_product_category", "categoryid"),
}

NOTE_EMPTY = "Ad yoxdur (boşdur)"
NOTE_FOREIGN = "Ad xarici dildə növ sözü ilə başlayır"
NOTE_NO_TYPE = "Adda növ sözü yoxdur (birbaşa brendlə başlayır)"
NOTE_UNKNOWN = "Növ sözü tanınmadı - əl ilə yoxlanmalı"


def local(tag):
    return etree.QName(tag).localname.lower() if isinstance(tag, str) else ""


def child_text(el, names):
    for ch in el:
        if local(ch.tag) in names:
            txt = (ch.text or "").strip()
            if txt:
                return txt
    for attr in names:
        if el.get(attr):
            return el.get(attr).strip()
    return ""


def find_product_elements(root):
    """Ən çox təkrarlanan, ad/qiymət/link sahəsi olan elementi məhsul kimi götürür."""
    counts = {}
    for el in root.iter():
        tag = local(el.tag)
        if tag in ("offer", "item", "product", "entry", "goods"):
            counts[tag] = counts.get(tag, 0) + 1
    if counts:
        tag = max(counts, key=counts.get)
        return [el for el in root.iter() if local(el.tag) == tag]
    # Naməlum struktur: uşaqlarında name/title olan elementlər
    return [el for el in root.iter() if any(local(c.tag) in NAME_TAGS for c in el)]


def category_map(root):
    """YML formatında <categories><category id=".."> xəritəsi."""
    cats = {}
    for el in root.iter():
        if local(el.tag) == "category" and el.get("id"):
            cats[el.get("id")] = (el.text or "").strip()
    return cats


def match_foreign(name_lower):
    for phrase in sorted(FOREIGN_TO_AZ, key=len, reverse=True):
        if re.match(rf"{re.escape(phrase)}(?![\w-])", name_lower):
            return phrase
    return None


def check_name(name, brand):
    """(problem, təklif) qaytarır; problem yoxdursa ("", "")."""
    name = (name or "").strip()
    if not name:
        return NOTE_EMPTY, ""
    lower = name.lower()
    first = re.split(r"[\s,/]+", lower)[0]

    foreign = match_foreign(lower)
    if foreign:
        az = FOREIGN_TO_AZ[foreign]
        # "Blender" kimi hər iki dildə eyni olan sözlər səhv sayılmır
        if az.lower() == foreign:
            return "", ""
        return NOTE_FOREIGN, az + name[len(foreign):]

    if first in AZ_TYPE_WORDS:
        return "", ""

    brand_l = (brand or "").lower()
    if first in KNOWN_BRANDS or (brand_l and lower.startswith(brand_l)):
        return NOTE_NO_TYPE, ""

    return NOTE_UNKNOWN, ""


def suggest_from_category(category):
    """Növ sözü olmayan ad üçün kateqoriya adından təklif (məs. 'Üzqırxanlar' -> 'Üzqırxan')."""
    if not category:
        return ""
    last = category.split(">")[-1].split("/")[-1].strip()
    return re.sub(r"(lar|lər)$", "", last, flags=re.IGNORECASE)


def load(source):
    if re.match(r"https?://", source):
        req = urllib.request.Request(source, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=120) as r:
            data = r.read()
    else:
        with open(source, "rb") as f:
            data = f.read()
    parser = etree.XMLParser(recover=True, huge_tree=True, resolve_entities=False, no_network=True)
    return etree.fromstring(data, parser)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", nargs="?", default=FEED_URL)
    ap.add_argument("-o", "--output", default="myshops_sehv_adlar.xlsx")
    args = ap.parse_args()

    root = load(args.source)
    cats = category_map(root)
    products = find_product_elements(root)
    if not products:
        sys.exit("XML-də məhsul tapılmadı.")

    rows = []
    for el in products:
        rec = {k: child_text(el, v) for k, v in FIELD_TAGS.items()}
        rec["id"] = rec["id"] or el.get("id", "")
        if rec["category"] in cats:
            rec["category"] = cats[rec["category"]]
        rec["name"] = child_text(el, NAME_TAGS)
        problem, fix = check_name(rec["name"], rec["brand"])
        if problem == NOTE_NO_TYPE and not fix:
            type_word = suggest_from_category(rec["category"])
            fix = f"{type_word} {rec['name']}" if type_word else ""
        rec["problem"], rec["fix"] = problem, fix
        rows.append(rec)

    bad = [r for r in rows if r["problem"]]
    cols = [("id", "ID"), ("name", "Hazırkı ad"), ("problem", "Problem"),
            ("fix", "Təklif olunan ad"), ("brand", "Brend"),
            ("category", "Kateqoriya"), ("price", "Qiymət"), ("url", "Link")]

    wb = Workbook()
    for title, data in (("Səhv adlar", bad), ("Bütün məhsullar", rows)):
        ws = wb.active if title == "Səhv adlar" else wb.create_sheet(title)
        ws.title = title
        ws.append([c[1] for c in cols])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="305496")
        for r in data:
            ws.append([r.get(k, "") for k, _ in cols])
        widths = [12, 50, 40, 50, 15, 30, 10, 50]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(args.output)

    print(f"Cəmi məhsul: {len(rows)}")
    for note in (NOTE_EMPTY, NOTE_FOREIGN, NOTE_NO_TYPE, NOTE_UNKNOWN):
        print(f"  {note}: {sum(r['problem'] == note for r in rows)}")
    print(f"Nəticə: {args.output}")


if __name__ == "__main__":
    main()
