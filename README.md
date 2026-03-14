# 📞 Agent Performance Analytics — BigQuery Pipeline

> **End-to-end automated pipeline** that ingests call data from 4 dialers (Ozonetel, TATA Smartflo, Exotel, Avyukta) into Google BigQuery, runs SQL-based performance scoring, and refreshes a Looker Studio dashboard every few hours — replacing 3 hours of daily manual Excel reporting for 60 sales agents.

---

## 🔍 Problem Statement

QHT Clinic's sales team runs across **4 parallel dialers**. Before this pipeline:
- A manual analyst pulled CSVs from each dialer every morning
- Data was stitched together in Excel — error-prone and took ~3 hours/day
- Leadership had **no intraday visibility** into agent performance
- Underperforming agents and connectivity issues went undetected until end-of-day

---

## 🏗️ Architecture

```
Ozonetel API ──┐
TATA Smartflo API ──┤
Exotel API ──┤  →  Google Apps Script  →  BigQuery  →  Looker Studio Dashboard
Avyukta API ──┘         (ingestion)       (transform)     (CEO / CBO / Sales Head)
```

**Refresh cadence:** Every 3 hours via BigQuery Scheduled Queries + Apps Script time-based triggers

---

## 📊 Dashboard Metrics

| Metric | Description |
|---|---|
| `total_calls` | Outbound calls made per agent per day |
| `talk_time_minutes` | Total connected talk time |
| `break_duration_minutes` | Total break time taken |
| `connectivity_pct` | % of calls that connected (answered / dialled) |
| `hourly_calls` | Calls-per-hour bucketed across the shift |
| `efficiency_index` | Composite score: calls × connectivity % ÷ break ratio |
| `dialer_connectivity_rank` | Which dialer has best live connectivity |

---

## 🗂️ Repository Structure

```
agent-performance-pipeline/
│
├── sql/
│   ├── staging/
│   │   ├── stg_ozonetel.sql          # Normalize Ozonetel raw API response
│   │   ├── stg_smartflo.sql          # Normalize TATA Smartflo data
│   │   ├── stg_exotel.sql            # Normalize Exotel data
│   │   └── stg_avyukta.sql           # Normalize Avyukta data
│   │
│   ├── analysis/
│   │   ├── agent_daily_scorecard.sql # Per-agent KPIs with RANK + LAG
│   │   ├── dialer_comparison.sql     # Cross-dialer connectivity analysis
│   │   ├── hourly_call_volume.sql    # Intraday call distribution
│   │   └── low_efficiency_flags.sql  # Agents flagged for low performance
│   │
│   └── scheduling/
│       └── scheduled_query_config.md # BigQuery scheduled query setup guide
│
├── python/
│   ├── ingest_dialers.py             # Calls all 4 dialer APIs → dumps to BigQuery
│   ├── apps_script_trigger.js        # Google Apps Script for time-based refresh
│   └── requirements.txt
│
├── dashboard/
│   ├── screenshots/                  # Dashboard screenshots
│   └── looker_data_sources.md        # How to connect Looker Studio to BigQuery
│
├── data/
│   └── sample_agent_data.csv         # Anonymised sample data for testing
│
└── docs/
    ├── setup_guide.md                # How to replicate this pipeline
    └── schema.md                     # BigQuery table schemas
```

---

## 💡 Key SQL Techniques Used

- **`RANK()`** — ranks agents by efficiency index within each shift
- **`LAG()`** — compares current day's talk time vs previous day (WoW trend)
- **`ROW_NUMBER()`** — deduplicates API responses where a call appears in 2 dialers
- **`CASE WHEN`** — flags agents with connectivity < 40% or calls < 20/day
- **`DATE_TRUNC` + `HOUR`** — hourly bucketing for intraday heatmaps

---

## 📈 Impact

| Before | After |
|---|---|
| 3 hrs/day manual Excel work | Fully automated, refreshes every 3 hours |
| End-of-day visibility only | Real-time intraday dashboard |
| No cross-dialer benchmarking | Live dialer connectivity rankings |
| Issues caught next day | Low-efficiency agents flagged same hour |
| Data for ~1 dialer at a time | Unified view across all 60 agents × 4 dialers |

---

## 🛠️ Tech Stack

| Layer | Tool |
|---|---|
| Data ingestion | Dialer REST APIs + Google Apps Script |
| Storage & transform | Google BigQuery |
| Scheduling | BigQuery Scheduled Queries (every 3 hrs) |
| Visualisation | Looker Studio |
| Scripting | Python 3.x (pandas, requests, google-cloud-bigquery) |

---

## 🚀 How to Replicate

See [`docs/setup_guide.md`](docs/setup_guide.md) for full instructions. At a high level:

1. Set up BigQuery project and create dataset `agent_analytics`
2. Configure dialer API credentials in `python/ingest_dialers.py`
3. Run staging SQL files to create normalised tables
4. Deploy Apps Script trigger for scheduled ingestion
5. Connect Looker Studio to BigQuery views
6. Set BigQuery Scheduled Query to refresh every 3 hours

---

## 👤 Author

**Rahul Saini** — Data Analyst  
[LinkedIn](https://linkedin.com/in/rahulsaini) · [Email](mailto:sainirahul430@gmail.com)
