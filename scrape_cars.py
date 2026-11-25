import re
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

start_url = "https://webscraper.io/test-sites/pagination/Ford"


def get_soup(url, page=None):
    params = {"page": page} if page else None
    response = requests.get(url, params=params)
    return BeautifulSoup(response.text, "html.parser")


def text_of(tag):
    return tag.get_text(strip=True) if tag else None


def digits(text):
    found = re.sub(r"\D", "", text or "")
    return int(found) if found else None


# 1. Find every make in the sidebar
first = get_soup(start_url)
makes = {}
for a in first.find_all("a", class_="category-link"):
    makes[urljoin(start_url, a["href"])] = a.get_text(strip=True)
if not makes:
    makes = {start_url: "Ford"}
print("Makes found:", list(makes.values()))

# 2. Go through every page of every make
cars = {}
for make_url, make in makes.items():
    page = 1
    while page <= 20:
        cards = get_soup(make_url, page).find_all("div", class_="sitemap-card")
        new = 0

        for card in cards:
            link = card.find("a", class_="card-head-url")["href"]
            if link in cars:
                continue
            new += 1

            details = {}
            for p in card.find_all("p", class_="card-text"):
                b = p.find("b")
                label = b.get_text(strip=True).rstrip(":")
                b.extract()
                details[label] = p.get_text(" ", strip=True)

            price_text = text_of(card.find("span", itemprop="price")) or ""
            currency = re.match(r"[A-Za-z]+", price_text)

            cars[link] = {
                "Make": make,
                "Car": text_of(card.find("h3", class_="card-title")),
                "Description": text_of(card.find("p", class_="description")),
                "Status": text_of(card.find("div", class_="badge")),
                "Year": digits(details.get("Year")),
                "Country": details.get("Country of origin"),
                "Mileage (km)": digits(details.get("Mileage")),
                "Rating": int(card.find(class_="rarity-rating")["data-rating"]),
                "Currency": currency.group() if currency else None,
                "Price": digits(price_text),
                "Link": link,
            }

        if len(cards) == 0 or new == 0:
            break
        page += 1
        time.sleep(1)

    count = sum(1 for c in cars.values() if c["Make"] == make)
    print(make, "done -", count, "cars")

df = pd.DataFrame(cars.values())
print("Currencies:", sorted(df["Currency"].dropna().unique()))

# 3. Summary sheets
summary = df.groupby("Make").agg(**{
    "Cars": ("Car", "count"),
    "Sold": ("Status", lambda s: int((s == "Sold").sum())),
    "Average Price": ("Price", "mean"),
    "Cheapest": ("Price", "min"),
    "Most Expensive": ("Price", "max"),
}).reset_index()
summary["Average Price"] = summary["Average Price"].round(0)

top = (df.sort_values(["Rating", "Price"], ascending=False)
         .head(15)[["Make", "Car", "Year", "Rating", "Price", "Status"]])

# 4. Save and format the Excel file
file = "cars_report.xlsx"
with pd.ExcelWriter(file) as writer:
    df.to_excel(writer, sheet_name="Cars", index=False)
    summary.to_excel(writer, sheet_name="Make Summary", index=False)
    top.to_excel(writer, sheet_name="Top Rated", index=False)

wb = load_workbook(file)
for ws in wb.worksheets:
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A2"
    for col in ws.columns:
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(longest + 3, 60)
wb.save(file)

print("Saved", len(df), "cars from", len(makes), "makes to", file)