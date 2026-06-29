"""
generate_dummy_leads.py
Generates a sample Google Sheets "contacts" tab export as a local CSV
for local dev and testing without needing live Sheet access.

Usage:
    python generate_dummy_leads.py           # writes inputs/sample_leads.csv
"""

import csv
import random
from pathlib import Path

random.seed(42)

ROOT   = Path(__file__).parent
INPUTS = ROOT / "inputs"
INPUTS.mkdir(exist_ok=True)

BUSINESSES = [
    ("The Ivy", "London", "Modern British", "Central London", "Moderate"),
    ("Dishoom", "London", "Indian", "Shoreditch", "Moderate"),
    ("Café Rouge", "Birmingham", "French", "City Centre", "Moderate"),
    ("Gaucho", "Manchester", "Argentinian", "Spinningfields", "Expensive"),
    ("Bill's", "Bristol", "All-day dining", "Clifton", "Moderate"),
    ("Côte Brasserie", "Leeds", "French", "City Centre", "Moderate"),
    ("Miller & Carter", "Liverpool", "Steakhouse", "Albert Dock", "Expensive"),
    ("Nando's", "Newcastle", "Portuguese", "Eldon Square", "Inexpensive"),
    ("Turtle Bay", "Sheffield", "Caribbean", "Division Street", "Moderate"),
    ("Zizzi", "Nottingham", "Italian", "Market Square", "Moderate"),
    ("Pizza Express", "Leicester", "Italian", "Highcross", "Moderate"),
    ("Franco Manca", "Oxford", "Pizza", "Covered Market", "Inexpensive"),
    ("Hawksmoor", "Cambridge", "Steakhouse", "City Centre", "Very Expensive"),
    ("Wagamama", "Coventry", "Japanese", "Broadgate", "Moderate"),
    ("Five Guys", "Reading", "Burgers", "The Oracle", "Moderate"),
]

DOMAINS = [
    "theivyrestaurant.com", "dishoom.com", "caferouge.co.uk", "gaucho.co.uk",
    "bills-website.co.uk", "cote.co.uk", "millerandcarter.co.uk", "nandos.co.uk",
    "turtlebay.co.uk", "zizzi.co.uk", "pizzaexpress.com", "francomanca.co.uk",
    "thehawksmoor.com", "wagamama.com", "fiveguys.co.uk",
]

COLUMNS = [
    "city", "location_count", "lead_tier", "location", "cuisine",
    "price_range", "scraped_at", "restaurant_name", "email", "phone",
    "address", "website_url", "instagram_url", "opentable_url",
    "rating", "review_count", "email_sent", "sent_at", "replied",
    "drive_audio_url", "audio_verified", "drive_video_url", "drive_thumbnail_url",
]


def fake_phone():
    return f"0{random.randint(1000,9999)} {random.randint(100000,999999)}"


def fake_postcode(city):
    prefix = city[:2].upper()
    return f"{prefix}{random.randint(1,9)} {random.randint(1,9)}{''.join(random.choices('ABCDEFGHKLMNPRSTUVWXY', k=2))}"


def fake_rating():
    return round(random.uniform(3.5, 5.0), 1)


def fake_reviews():
    return random.randint(50, 8000)


def fake_location_count():
    n = random.choices([1, random.randint(2, 5), random.randint(6, 30)], weights=[50, 30, 20])[0]
    return n


rows = []
for i, (name, city, cuisine, location, price) in enumerate(BUSINESSES):
    domain = DOMAINS[i]
    loc_count = fake_location_count()
    tier = "Independent" if loc_count == 1 else ("Small chain" if loc_count <= 5 else "Large chain")
    postcode = fake_postcode(city)
    address = f"{random.randint(1,200)} {location}, {city}, {postcode}"

    row = {
        "city":            city,
        "location_count":  loc_count,
        "lead_tier":       tier,
        "location":        location,
        "cuisine":         cuisine,
        "price_range":     price,
        "scraped_at":      "2026-06-01",
        "restaurant_name": name,
        "email":           f"hello@{domain}",
        "phone":           fake_phone(),
        "address":         address,
        "website_url":     f"https://www.{domain}/",
        "instagram_url":   f"https://www.instagram.com/{name.lower().replace(' ','').replace('&','and')}/",
        "opentable_url":   f"https://www.opentable.co.uk/r/{name.lower().replace(' ','-').replace(\"'\",'')}",
        "rating":          fake_rating(),
        "review_count":    fake_reviews(),
        "email_sent":      "",
        "sent_at":         "",
        "replied":         "",
        "drive_audio_url": "",
        "audio_verified":  "",
        "drive_video_url": "",
        "drive_thumbnail_url": "",
    }
    rows.append(row)

out = INPUTS / "sample_leads.csv"
with open(out, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=COLUMNS)
    writer.writeheader()
    writer.writerows(rows)

print(f"✓ {len(rows)} sample leads → {out}")
print("  Import this CSV into Google Sheets to test the pipeline without live scraping.")
