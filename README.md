\# Tech Show London 2026 - Resilient Exhibitor Scraper

A robust, modular web automation tool built with \*\*Python\*\*,
\*\*Playwright\*\*, and \*\*Pandas\*\* to extract exhibitor data from
the Tech Show London 2026 website.

\## Key Technical Features

\### 1. Dynamic Element Handling (Anti-Stale Architecture) Instead of
fetching a static list of elements, this script uses
\`range(total_containers)\` combined with \`.nth(i)\`. This ensures a
\"fresh\" reference to the DOM is requested at every iteration,
preventing the common \*\*StaleElementReferenceError\*\* caused by
dynamic website updates.

\### 2. Fault-Tolerant Extraction & UI Recovery The script features a
sophisticated error-handling flow. If a specific company modal fails to
load: \* The \`try-except\` block catches the exception to prevent a
total crash. \* A \*\*UI Recovery routine\*\* checks if a modal is stuck
open and closes it before proceeding. \* The loop continues to the next
item, ensuring maximum data collection even under unstable network
conditions.

\### 3. Human Behavior Simulation To bypass basic anti-bot security and
rate-limiting, the scraper implements: \* \*\*Throttling:\*\* Randomized
delays using \`random.uniform\`. \* \*\*Smart Waiting:\*\* Utilizing
\`wait_for_load_state(\"networkidle\")\` to ensure pages are fully
interactive before clicking.

\### 4. Automated Data Cleaning Integrated \*\*Pandas\*\* pipeline to:
\* Remove duplicate entries based on unique company names. \* Sort data
alphabetically (A-Z). \* Export to a clean, UTF-8 encoded CSV.

\## Installation & Usage

1\. \*\*Install Dependencies:\*\* \`\`\`bash pip install playwright
pandas playwright install chromium

2\. \*\*Run the Scraper\*\*:

\* python main.py

3\. \*\*Output\*\* The results are saved in a CSV file named
web_scrape_YYYY_MM_DD.csv containing:

\* Company Name

\* Website URL

\* LinkedIn Profile URL
