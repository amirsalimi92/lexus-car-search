"""
Lexus RX/NX scraper — AutoScout24 (Europe-wide)
Filters: 2016+, max 30.000 EUR, Automatic, No accident
"""

import requests, json, time, os
from bs4 import BeautifulSoup
from datetime import datetime

# Use script directory for all output files
DATA_DIR = os.path.dirname(os.path.abspath(__file__))

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.8",
    "Sec-Ch-Ua": '"Chromium";v="124", "Google Chrome";v="124"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Upgrade-Insecure-Requests": "1",
}

BASE = "https://www.autoscout24.de"
MAX_PRICE = 30000

SEARCHES = [
    ("rx-serie-(alle)", "RX"),
    ("nx-serie-(alle)", "NX"),
]

COLOR_MAP = {
    "grau": "Grau", "schwarz": "Schwarz", "weiss": "Weiß",
    "silber": "Silber", "blau": "Blau", "rot": "Rot", "gruen": "Grün",
    "beige": "Beige", "braun": "Braun", "gelb": "Gelb",
    "orange": "Orange", "violett": "Violett", "gold": "Gold", "weib": "Weiß",
}

SUNROOF_KW = ["pano", "panorama", "schiebedach", "glasdach", "sunroof"]


def detect_color_from_url(url_path):
    url_lower = url_path.lower()
    for key, val in COLOR_MAP.items():
        if f"-{key}-" in url_lower or url_lower.endswith(f"-{key}"):
            return val
    return ""


def detect_sunroof(text):
    return any(kw in (text or "").lower() for kw in SUNROOF_KW)


def fetch_page(slug, page=1):
    url = (
        f"{BASE}/lst/lexus/{slug}"
        f"?atype=C&damaged_listing=exclude"
        f"&fregfrom=2016&priceto={MAX_PRICE}"
        f"&sort=price&desc=0&page={page}"
    )
    resp = requests.get(url, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    script = soup.find("script", id="__NEXT_DATA__")
    if not script:
        return [], 0
    data = json.loads(script.string)
    props = data["props"]["pageProps"]
    return props.get("listings", []), int(props.get("numberOfPages", 1) or 1)


def fetch_listing_details(listing_url):
    """Fetch individual listing page to get noOfPreviousOwners + accurate bodyColor."""
    try:
        resp = requests.get(BASE + listing_url, headers=HEADERS, timeout=12)
        if resp.status_code != 200:
            return None, None
        soup = BeautifulSoup(resp.text, "html.parser")
        script = soup.find("script", id="__NEXT_DATA__")
        if not script:
            return None, None
        data = json.loads(script.string)
        v = data["props"]["pageProps"].get("listingDetails", {}).get("vehicle", {})
        owners = v.get("noOfPreviousOwners")
        color = v.get("bodyColor") or v.get("bodyColorRaw") or None
        return owners, color
    except Exception:
        return None, None


def parse(raw, target_group):
    v = raw.get("vehicle", {})
    if v.get("modelGroup", "").upper() != target_group.upper():
        return None

    tracking = raw.get("tracking", {})
    reg_raw = tracking.get("firstRegistration", "")
    if reg_raw:
        parts = reg_raw.split("-")
        first_reg = f"{parts[0]}/{parts[1]}" if len(parts) == 2 else reg_raw
    else:
        first_reg = ""

    details_map = {d.get("iconName", ""): d.get("data", "") for d in raw.get("vehicleDetails", [])}
    transmission = details_map.get("gearbox", v.get("transmission", ""))
    if transmission and "automatik" not in transmission.lower() and "automatic" not in transmission.lower():
        return None  # skip manual

    price_data = raw.get("price", {})
    loc = raw.get("location", {})
    url_path = raw.get("url", "")
    full_text = f"{(v.get('modelVersionInput') or '')} {(v.get('subtitle') or '')}"

    label_map = {1: "Top Preis", 2: "Guter Preis", 3: "Fairer Preis", 4: "Erhöhter Preis", 5: "Hoher Preis"}

    return {
        "source": "AutoScout24",
        "listing_id": raw.get("id", ""),
        "_url_path": url_path,
        "model": f"Lexus {v.get('modelGroup', '')}",
        "variant": v.get("model", ""),
        "title": (v.get("modelVersionInput") or "").strip(),
        "description_de": (v.get("subtitle") or "").strip(),
        "price_eur": int(tracking.get("price", 0) or 0),
        "price_label": label_map.get(price_data.get("priceEvaluation"), ""),
        "prev_price_eur": None,
        "first_registration": first_reg,
        "year": first_reg[-4:] if len(first_reg) >= 4 else "",
        "mileage": details_map.get("mileage_odometer", v.get("mileageInKm", "")),
        "fuel": v.get("fuel", ""),
        "transmission": transmission,
        "color": detect_color_from_url(url_path),
        "sunroof": "Ja" if detect_sunroof(full_text) else "Nein",
        "no_accident": "Ja",
        "previous_owners": None,
        "city": loc.get("city", ""),
        "country": loc.get("countryCode", ""),
        "link": BASE + url_path if url_path else "",
        "description_fa": "",
    }


def scrape_as24(slug, group):
    results = []
    try:
        raw_list, total = fetch_page(slug, 1)
    except Exception as e:
        print(f"  [{group}] ERROR page 1: {e}")
        return []

    for r in raw_list:
        p = parse(r, group)
        if p:
            results.append(p)

    for page in range(2, min(total + 1, 8)):
        time.sleep(1)
        try:
            raw_list, _ = fetch_page(slug, page)
            for r in raw_list:
                p = parse(r, group)
                if p:
                    results.append(p)
        except Exception as e:
            print(f"  [{group}] ERROR page {page}: {e}")
            break

    print(f"  [{group}] {len(results)} listings — fetching details...")

    for i, car in enumerate(results):
        url_path = car.pop("_url_path", "")
        if url_path and "/angebote/" in url_path:
            owners, color = fetch_listing_details(url_path)
            if owners is not None:
                car["previous_owners"] = owners
            if color:
                car["color"] = color
            if (i + 1) % 10 == 0:
                print(f"    {i+1}/{len(results)} details fetched")
            time.sleep(0.3)

    return results


def run():
    print(f"\n{'='*55}\nLexus RX & NX — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"Filter: 2016+, max {MAX_PRICE:,} EUR, Automatic, No accident\n{'='*55}\n")

    all_results = []
    for slug, group in SEARCHES:
        all_results += scrape_as24(slug, group)
        time.sleep(2)

    # Deduplicate by link
    seen, unique = set(), []
    for car in all_results:
        car.pop("_url_path", None)
        key = car["link"] or (car["title"] + str(car["price_eur"]))
        if key not in seen:
            seen.add(key)
            unique.append(car)

    unique.sort(key=lambda x: x.get("price_eur") or 999999)

    output = {
        "last_updated": datetime.now().isoformat(),
        "source": "AutoScout24",
        "total": len(unique),
        "listings": unique,
    }

    # Save as both as24 and combined (GitHub Actions uses AS24 only)
    for fname in ["lexus_as24.json", "lexus_combined.json"]:
        path = os.path.join(DATA_DIR, fname)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\nSaved {len(unique)} listings")
    return output


if __name__ == "__main__":
    run()
