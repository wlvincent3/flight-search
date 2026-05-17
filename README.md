# Flight Search Tool

A Python script that searches Google Flights via SerpAPI for multi-airport, multi-stop-count flight combinations and exports results to a formatted Excel workbook for easy comparison.

## The Problem

Searching flights manually means running the same query multiple times — once per destination airport, once per stop count, once per return date. For a trip with three possible destination airports, two return dates, and two stop limits, that's 12 separate searches. This does all of them in one run and puts the results in a spreadsheet.

## What This Does

- Searches multiple destination airport codes in a single run (e.g., primary airport + nearby alternatives)
- Tries multiple stop configurations (1-stop max, 2-stop max)
- Tries multiple return date options
- Exports all results to a formatted Excel workbook with:
  - Flight details (airline, flight numbers, departure/arrival times)
  - Layover breakdowns by airport and duration
  - Total trip duration and price
  - Sorted by price within each search configuration

## Setup

```bash
pip install requests openpyxl
```

Get a free API key from [SerpAPI](https://serpapi.com) (100 free searches/month).

Create a `.env` file:
```
SERPAPI_KEY=your_key_here
```

Or pass it directly:
```bash
SERPAPI_KEY=your_key python3 flight_search.py
```

## Configuration

Edit the top of `flight_search.py` to set your search parameters:

```python
DESTINATIONS = [("JFK", "New York"), ("EWR", "Newark")]  # airport code + label
DEPART_DATE  = "2026-07-01"
RETURN_DATES = ["2026-07-14", "2026-07-15"]
STOP_OPTIONS = [1, 2]     # max stops per search
# travel_class: 1=economy, 2=premium economy, 3=business, 4=first
```

Output: `flights_YYYY-MM-DD.xlsx` in the current directory.

## Built With

Python 3, `requests`, `openpyxl`. Data from [SerpAPI Google Flights](https://serpapi.com/google-flights-api). Built with AI assistance (Claude).
