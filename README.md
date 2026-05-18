# Flight Search

Searches Google Flights via SerpAPI across multiple airport codes, stop counts, and return dates in a single run. Exports results to a formatted Excel workbook.

## What It Does

- Queries multiple destination airports (e.g., JFK + EWR + LGA)
- Varies stop count (1-stop max, 2-stop max) and return dates
- Exports all results to an Excel workbook with airline, flight numbers, layover breakdowns, duration, and price
- Deduplicates across search combinations

## Setup

```bash
pip install requests openpyxl
```

Get a free API key from [SerpAPI](https://serpapi.com) (100 searches/month on the free tier).

```bash
cp .env.example .env
# add your SERPAPI_KEY to .env
```

## Configuration

Edit the top of `flight_search.py`:

```python
DESTINATIONS = [("JFK", "New York JFK"), ("EWR", "Newark")]
DEPART_DATE  = "2026-07-12"
RETURN_DATES = ["2026-07-20", "2026-07-21"]
STOP_OPTIONS = [1, 2]
```

## Usage

```bash
python3 flight_search.py
```

Output: `flights_results.xlsx` in the current directory.

## Stack

Python 3, `requests`, `openpyxl`. Data from [SerpAPI Google Flights](https://serpapi.com/google-flights-api).
