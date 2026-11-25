# Car Listings Scraper



Python scraper that collects listings from a paginated demo site with several categories, cleans the messy text fields, and delivers a formatted Excel report.



## What it does



\- \*\*Categories:\*\* reads the list of car makes from the site's sidebar, so none are typed in by hand.

\- \*\*Pagination:\*\* goes through every page of every make (100 cars from 8 makes) and skips duplicates.

\- \*\*Cleaning:\*\* turns text like `USD 235 600` and `212 581 km` into real numbers, and reads each rating from the page data.

\- \*\*Report:\*\* one Excel file with 3 sheets: Cars (all data), Make Summary (count, sold, average, cheapest and most expensive price per make) and Top Rated.



## Tools used



Python, requests, BeautifulSoup, pandas, openpyxl



## How to run



&#x20;   pip install -r requirements.txt

&#x20;   python scrape\_cars.py



Output: `cars\_report.xlsx`



## Data source



\[webscraper.io test sites](https://webscraper.io/test-sites), a demo site built for practising web scraping. All listings are fictional.



## Available for freelance work



I build scrapers that collect data from websites and deliver it as clean Excel or CSV files.

