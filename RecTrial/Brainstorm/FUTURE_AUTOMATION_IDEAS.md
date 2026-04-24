# Future Automation Ideas — iPipeline Finance

Living document of automation + workflow improvement ideas worth exploring after the 4-video demo project ships. Kept in two parts: (1) things we can build ourselves using code + tools we already have, (2) external third-party software / tools / AI platforms that could slot in.

**Started:** 2026-04-22
**Owner:** Connor — Finance & Accounting, iPipeline
**Update style:** add as you think of them; strike through or move to "Done" section as they're delivered. This is a scratchpad, not a plan.

---

## Part 1 — Internal builds: AI + Excel/Python/VBA/SQL combos

Ideas we can build using what iPipeline already licenses or with minimal new spend. These lean on Microsoft 365 Copilot, Python, VBA, SQL, and whatever we have API access to.

### 1.1 AI-powered Finance narratives

- **Python generates data → M365 Copilot writes narrative.** Python exports a variance CSV. User pastes into Excel. Copilot sidebar turns it into a 3-bullet CFO summary. Demo-friendly, zero new tooling. (Pitched as Option 1 for Video 4.)
- **Python script that calls OpenAI/Claude/Azure OpenAI API directly** to auto-generate variance commentary, budget narratives, or exec briefs. Needs an API key and IT approval for outbound calls.
- **Local LLM via Ollama** — run a small model offline on the analyst's laptop for privacy-safe narrative generation. No API, no data leaves the laptop.
- **AI-assisted exception triage** — send "questionable" transactions to AI with context, get a recommended disposition ("likely Travel — approve," "likely duplicate — flag").
- **AI-generated board deck slides.** Python pulls the month's numbers → writes a slide deck skeleton → Copilot in PowerPoint polishes each slide.

### 1.2 Month-end close automation

- **"One Monday Morning" end-to-end orchestrator** — one Python script that on Monday 7am pulls bank + GL, reconciles, generates variance commentary, builds a PDF, emails it. Uses Windows Task Scheduler for the cron piece.
- **Email-driven automation** — coworker forwards an email with an Excel attachment → Python processes it → replies with results. Via Outlook's MAPI or Graph API.
- **Recurring scheduled reports** — dashboards auto-refresh and auto-email every Friday at 5pm.

### 1.3 Data quality + validation at scale

- **Watch folder script** — Python watches a shared network folder, auto-processes anything dropped in.
- **Mass file sanitizer** — take a folder of 100 CSVs, run our existing sanitize tools on all of them in one pass.
- **Cross-file referential integrity check** — Python verifies that IDs in File A also exist in File B (common reconciliation pattern).

### 1.4 VBA + Python integration

- **`xlwings` bridge** — Excel button calls Python behind the scenes. Coworker never sees Command Prompt.
- **VBA-triggered Python pipelines** — press a Command Center button → Python runs a heavy job → results return to Excel.
- **Python-generated Excel dashboards** — Python builds the entire branded workbook from scratch (charts, formatting, formulas). Useful when the same report is needed every month.

### 1.5 SQL + Python combinations

- **SQL pull → Python transform → Excel push** pattern. Automated daily data refresh from Azure SQL / SQL Server / SQLite into branded Excel dashboards.
- **dbt-lite in Python** — organize SQL transformations as versioned scripts.
- **Python-driven data validation on database queries** — verify row counts, nulls, and anomaly thresholds before a report publishes.

### 1.6 Document + PDF work

- **PDF table extraction + structured output.** Python + pdfplumber to pull tables out of vendor invoices, bank statements, regulatory filings.
- **Contract key-term extractor** — AI reads a 20-page vendor contract → extracts start date, end date, renewal terms, payment terms, penalties. Saves hours of manual review.
- **Auto-generate tailored PDF packages per stakeholder** — same data, different highlights/emphasis per audience.

### 1.7 Reporting & dashboards

- **Streamlit / Dash local dashboards** — a lightweight web app on the analyst's laptop with all the KPIs. No IT support needed; just `streamlit run`.
- **Auto-emailed dashboard snapshots** — every Monday, a PNG of the dashboard lands in leadership's inbox.
- **Interactive what-if widgets** — sliders that recompute models live.

### 1.8 Collaboration & governance

- **Git + GitHub for Finance code** — version every script, track who changed what.
- **Shared Copilot prompt library** — centralized list of the team's best Copilot prompts for repeated tasks.
- **Audit log sheet on every workbook** — auto-updated each time someone opens or modifies the file.

### 1.9 Communications + meeting automation

- **Meeting notes → action items** — Copilot processes Teams meeting transcript, extracts action items, assigns due dates, posts to a tracker.
- **Auto-draft reply bot** — reads the last N emails on a thread, drafts a reply in your voice.
- **Team status roll-up** — Python pulls updates from multiple sources (email, Teams, SharePoint), builds a single team update.

---

## Part 2 — Third-party tools, software, AI platforms

External tools worth investigating. Some are free, some licensed, some enterprise-only. Add effort/cost rough estimates as you research.

### 2.1 Workflow automation platforms

- **Microsoft Power Automate** — likely already licensed. Low-code flows for email → Excel, SharePoint triggers, approval routing.
- **Microsoft Power Automate Desktop** — free, handles GUI automation on Windows.
- **Zapier / Make** — SaaS automation, hundreds of connectors. Cloud-only.
- **n8n** — open-source self-hosted alternative to Zapier. Privacy-friendly.
- **Workato** — enterprise-grade orchestration. Expensive but powerful.

### 2.2 RPA (Robotic Process Automation)

- **UiPath** — most common enterprise RPA. Record macros across apps.
- **Automation Anywhere** — another big RPA player.
- **Microsoft Power Automate Desktop** — RPA-lite, already Microsoft.

### 2.3 Data + analytics platforms

- **Power BI** — likely already licensed. Best-in-class Microsoft dashboards.
- **Tableau** — alternative dashboarding. License cost.
- **Looker Studio (free)** — Google's free dashboard tool. Works with Google Sheets + SQL.
- **Metabase (open-source)** — self-hosted BI, free, simple to deploy.
- **Mode Analytics / Hex** — modern notebook-based BI for data teams.

### 2.4 Python + Data tooling

- **xlwings** — Python in Excel via COM. Free. Keeps coworkers in Excel.
- **PyXLL** — commercial, similar to xlwings but Excel-native.
- **Streamlit / Dash / Gradio** — Python → web app. Free.
- **pandas / polars** — data processing libraries. Free.
- **dbt** — SQL transformation framework. Community edition free.
- **Jupyter Notebooks** — interactive Python. Free.
- **VSCode + Python** — IDE. Free.

### 2.5 AI / LLM platforms

- **Microsoft 365 Copilot** — already licensed for most Finance users. UI-based.
- **Copilot Studio** — build custom Copilot bots with iPipeline data.
- **Azure OpenAI** — enterprise-grade API access if IT enables it.
- **OpenAI API (direct)** — personal credit card, requires IT approval for corporate data.
- **Anthropic Claude API** — alternative LLM provider, strong at reasoning + long documents.
- **Google Gemini API** — third option.
- **Ollama** — run LLMs locally, no cloud. Free.
- **Perplexity** — AI-powered search, good for research.
- **Humata / Chat with PDF** — point AI at a PDF, ask questions.

### 2.6 Specialized Finance tools

- **Xero / QuickBooks integrations** — APIs if iPipeline uses them for any subsidiary.
- **FloQast / BlackLine** — close automation platforms.
- **Prophix / Vena** — FP&A platforms. Enterprise.
- **Adaptive Insights / Workday Adaptive Planning** — planning + forecasting.
- **Python `finance-datareader` / `yfinance` / FRED API** — free market + macro data.

### 2.7 PDF + document processing

- **Adobe Acrobat Pro** — already licensed, has Export-to-Excel + batch automation.
- **ABBYY FineReader** — high-end OCR.
- **Amazon Textract / Azure Form Recognizer** — cloud OCR for structured documents.
- **pdfplumber / PyMuPDF / camelot-py** — open-source PDF libraries in Python.

### 2.8 Communication + meetings

- **Fireflies.ai / Otter.ai / tl;dv** — record + transcribe + summarize meetings automatically.
- **Microsoft Teams Copilot** — built-in meeting summary, already licensed probably.
- **Loom** — async video explanations. Free tier.

### 2.9 Low-code app builders

- **Microsoft Power Apps** — build custom apps without code. Already licensed.
- **Airtable** — spreadsheet + database hybrid. Free tier.
- **Notion** — all-in-one workspace. Free personal tier.
- **Retool / Appsmith** — internal tool builders. Free tier / open source.

### 2.10 Security + credentials

- **Azure Key Vault** — secure API key storage for scripts.
- **Python `keyring` library** — local password manager accessible from code.
- **1Password CLI** — programmatic credential access from scripts.

---

---

## Parked from Video 4 brainstorm (2026-04-22)

Constraints Connor flagged during V4 planning — keeping these out of V4 scope but preserving for later:

### AI API integrations (external calls to OpenAI/Claude/Gemini)
- **Python + OpenAI/Claude variance narrative generator** — script sends data to LLM API, gets back plain-English executive summary. Needs API key + IT approval for outbound AI calls.
- **AI expense classifier** — feeds transaction descriptions to an LLM, gets back categorized output ("Travel", "Software", "Professional Services").
- **AI-powered contract reader** — point an API at a multi-page vendor contract, extract key terms (dates, penalties, renewal clauses).
- **AI anomaly explainer** — existing variance flagger pipes results to an LLM for plain-English "why" reasoning.
- **Pre-req:** verify iPipeline IT policy on outbound AI API calls and corporate data sovereignty. Alternative path: Azure OpenAI if iPipeline has enterprise access (check with IT).

### Outlook / email automation
- **Outlook inbox robot** — Python reads emails, finds attachments matching a pattern, processes them, auto-replies with results. Common pattern: "send me your expense report via email, I'll reconcile and reply."
- **Scheduled email reports** — Python builds a summary, sends it via Outlook MAPI or Microsoft Graph API every Monday at 7am.
- **Email-triggered reconciliation** — coworker forwards an invoice email → Python auto-processes attachment → reply with variance report.
- **Risk flag:** company email policy + IT concerns about bot-like auto-reply. Worth a conversation with IT before building.

### Windows Task Scheduler / scheduled automation
- **Auto-refresh dashboards** — scheduler runs a Python script every hour to pull latest data.
- **Auto-run month-end pack** — scheduler triggers the consolidated reporting pipeline on the 1st of each month.
- **Auto-sanitize inbound files** — scheduler watches a SharePoint folder and runs cleanup tools on new arrivals.
- **Built-in to Windows** — no install needed. Runs any script or app on a timer/trigger.

### Azure OpenAI exploration
- **Check if iPipeline enterprise Azure has OpenAI endpoint enabled** — if yes, it's the IT-approved way to get LLM access without sending data to OpenAI directly.
- **Custom Copilot Studio bots** — Microsoft's platform for building custom AI bots with iPipeline data. Enterprise-licensed, often already available.

## How to use this document

- Add to it freely when you have a new idea. Don't worry about where it fits — rough buckets are fine.
- Rate each idea by **impact × effort** when you're ready to prioritize.
- Anything that becomes a real project gets moved to `tasks/todo.md` with a proper plan.
- This is a **scratchpad**, not a commitment. Ideas can sit here forever unmerged — that's fine.

## Proposed next review

Revisit after the 4-video project is fully delivered. Pick the 3 highest-impact ideas as the first post-project initiative.
