
# 🇮🇹 LinkedIn Job Scraper - Italy GIS & Geospatial Jobs

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Apify](https://img.shields.io/badge/Powered%20by-Apify-orange.svg)](https://apify.com)

An automated job scraper that finds GIS (Geographic Information Systems), Geospatial, and Surveying jobs in Italy from LinkedIn. Perfect for geospatial professionals, GIS analysts, and surveyors looking for opportunities in the Italian job market.

## ✨ Features

- 🔍 **30+ Smart Keywords** - Searches in both English and Italian (GIS, Remote Sensing, Geomatica, Telerilevamento, etc.)
- 🎯 **Precise Filters** - Last 24 hours, Entry level/Internship, Full-time/Part-time/Internship
- 🗺️ **Italy-Focused** - Specifically configured for the Italian job market
- 🚫 **Smart Exclusions** - Automatically filters out unwanted titles (Stage, Tirocinio)
- 🔄 **Automatic Deduplication** - Removes duplicate jobs found across multiple keywords
- 📊 **Excel Export** - Clean, formatted spreadsheet with clickable job links
- ⚡ **Fast** - Parallel processing scrapes all keywords in ~2 minutes
- 🔐 **Secure** - API tokens stored locally, never committed to GitHub

## 📊 Sample Results

On first run, the scraper found **219 unique GIS/Geospatial jobs** in Italy across 30 keywords.

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- [Apify](https://apify.com) account (free tier available)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/zafariabbas68/LinkedIn-Job-Scraper.git
   cd LinkedIn-Job-Scraper
   ```

2. **Create and activate virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Set up Apify API token**
   - Sign up at [apify.com](https://apify.com)
   - Go to Settings → API & Integrations
   - Copy your personal API token
   - Create `.env` file:
   ```bash
   echo 'APIFY_API_TOKEN=your_token_here' > .env
   ```

5. **Run the scraper**
   ```bash
   python linkedin_job_scraper.py
   ```

## 📁 Output

The script generates `jobs_italy.xlsx` with the following columns:

| Column | Description |
|--------|-------------|
| Application Status | Track your applications (manual entry) |
| Source | Platform (LinkedIn) |
| Job Title | Position name |
| Company | Employer name |
| Location | City/region in Italy |
| Job Type | Full-time, Part-time, Internship |
| Posted | Posting date |
| Applicants | Number of applicants (if visible) |
| Keywords Matched | Which search terms found this job |
| Job URL | Clickable link to apply |

## 🔧 Configuration

Easily customize the scraper by editing variables in `linkedin_job_scraper.py`:

```python
# Location settings
LOCATION = "Italy"                    # Change to specific city
GEO_ID = "103644278"                  # LinkedIn geo ID for Italy

# Search parameters
MAX_RESULTS_PER_SEARCH = 500          # Jobs per keyword
SEARCH_CONCURRENCY = 10               # Parallel searches

# Filters
PUBLISHED_AT = "r86400"               # Last 24 hours
CONTRACT_TYPES = ["F", "P", "I"]     # Full-time, Part-time, Internship
EXCLUDED_TITLE_TERMS = ["Stage", "Tirocinio"]

# Keywords (add or remove as needed)
KEYWORDS = [
    "GIS",
    "Remote Sensing",
    "Geospatial",
    "Sistemi Informativi Geografici",
    # ... 30+ keywords total
]
```

## 📈 Keywords Scraped

### English Keywords
GIS, GIS Analyst, GIS Developer, GIS Specialist, Remote Sensing, Earth Observation, Cartography, Geomatics, Geospatial, Spatial Data, 3D Mapping, Geospatial Analyst, GeoAI, Geodata

### Italian Keywords
Sistemi Informativi Geografici, Analista GIS, Sviluppatore GIS, Geomatica, Telerilevamento, Osservazione della Terra, Cartografia, Dati Geospaziali, Geodati, Fotogrammetria, Topografia, Rilievo, Geodesia, Urbanistica GIS, Ambiente GIS

## ⏰ Automating with Cron (macOS/Linux)

Run the scraper daily at 8 AM:

```bash
crontab -e
```

Add this line:
```cron
0 8 * * * cd /path/to/LinkedIn-Job-Scraper && /path/to/venv/bin/python linkedin_job_scraper.py >> scraper.log 2>&1
```

## 🛠️ Tech Stack

- **Python 3.8+** - Core language
- **Apify API** - LinkedIn job scraping
- **Requests** - API communication
- **OpenPyXL** - Excel file generation
- **ThreadPoolExecutor** - Parallel processing

## 📝 Notes

- Each run creates a new sheet in `jobs_italy.xlsx` with timestamp
- The script respects LinkedIn's rate limits
- Apify charges per result (starting from ~$0.30/1000 results)
- Free tier includes $5 credit for new accounts

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Add more Italian GIS keywords
- Improve filtering logic
- Add Indeed.com support
- Add Google Sheets export

## 📄 License

MIT License - feel free to use and modify for your own job search!

## ⚠️ Disclaimer

This tool is for personal job search use only. Please respect LinkedIn's terms of service and Apify's usage policies.

## 🌟 Star This Project

If this helped you find a job in Italy, please star the repository! ⭐

