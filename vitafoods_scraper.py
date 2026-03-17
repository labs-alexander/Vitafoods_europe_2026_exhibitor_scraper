import pandas as pd
from playwright.sync_api import sync_playwright
import time
import random
import os
import logging
import data_cleaner
from curl_cffi import requests as curl_requests
from datetime import datetime

# UI and Progress Monitoring Libraries
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn, TimeRemainingColumn

console = Console()

# Error logging configuration for traceability
logging.basicConfig(
    filename='extraction_errors.log',
    level=logging.ERROR,
    format='%(asctime)s - Entity: %(message)s'
)


class VitafoodsScraper:
    """
    Main extraction engine for the Vitafoods exhibitor directory.
    Handles data fetching, localized backup, and profile scraping.
    """
    def __init__(self):
        """Initializes state variables and backup paths."""
        self.today = datetime.today().strftime('%Y-%m-%d')
        self.website_api = "https://exhibitors.vitafoods.eu.com/live/search/search_exhibition46json.jsp?site=47&eventid=598&eventid=598&types=all"
        self.master_dict = {}
        self.processed_names = set()
        self.backup_file = f"backup_real_time_{self.today}.xlsx"

    def format_report(self):
        """Formats the final Excel output via the external cleaner module."""
        df = pd.DataFrame(list(self.master_dict.values()))
        data_cleaner.format_file(df, f"Vitafoods_exhibitors_{self.today}.xlsx")

    def get_identificators(self, item):
        """Extracts and sanitizes required identification data from the JSON payload."""
        raw_name = str(item.get("name", "N/A"))
        clean_name = raw_name.replace("xx", "'")
        emp_name = clean_name.upper().strip() 
        
        if emp_name in self.processed_names:
            return None
        
        emp_id = item.get("id")
        path = self.format_id_path(emp_id)
        return path, emp_name, emp_id

    def resolve_json_response(self):
        """Fetches the main directory list via HTTP request to the target API."""
        try:
            response = curl_requests.get(self.website_api, impersonate="chrome124")
            response.raise_for_status()
            json_file = response.json()
            print("Json loaded")
        except curl_requests.exceptions.HTTPError as e:
            logging.error(f"Response Error: {e}")
            return {}
        return json_file

    def format_id_path(self, emp_id):
        """
        Constructs the specific URL path by slicing the ID. 
        This allows direct access to the data source, bypassing portal navigation limits.
        """
        s = str(emp_id).zfill(6) 
        return f"{s[0:2]}/{s[2:4]}/{s[4:6]}"

    def final_dataframe(self):
        """Generates the unformatted base Excel file as a final fallback."""
        final_df = pd.DataFrame(list(self.master_dict.values()))
        final_df.drop_duplicates(subset="Company Name", inplace=True)
        final_df.to_excel("Vitafoods_Exhibitors_2026_Final_Report.xlsx", index=False)   

    def save_data(self):
        """Saves current progress to a local Excel backup file."""
        df = pd.DataFrame(list(self.master_dict.values()))
        df.to_excel(self.backup_file, index=False)

    def resolve_backup(self):
        """Attempts to load previously saved data to resume scraping."""
        if not os.path.exists(self.backup_file):
            console.print(f"[bold yellow]⚠️ No valid backups available. Starting fresh.[/bold yellow]")
            return
            
        try:
            df_existing = pd.read_excel(self.backup_file)
            # Identify entries that already have a valid Bio to avoid redundant processing
            for _, row in df_existing.iterrows():
                name = str(row['Company Name']).upper().strip()
                bio = str(row.get('Short Description', 'N/A'))
                clean_row = row.to_dict()
                self.master_dict[name] = clean_row
                
                if bio != "N/A" and len(bio) > 10:
                    self.processed_names.add(name)
                    
            console.print(f"[bold green]✅ Backup loaded: {len(self.processed_names)} results already completed.[/bold green]")
        
        except Exception as e:
            console.print(f"[bold red]❌ Error reading backup (File open or corrupted): {e}[/bold red]")
            console.print("[bold yellow]⚠️ Ignoring backup and starting from scratch.[/bold yellow]")
        
    def scrape_profile_data(self, page, detail_url, emp_name):
        """Navigates to the specific profile page and extracts Stand Number and Description."""
        page.goto(detail_url, wait_until="domcontentloaded", timeout=25000)
        
        location = "N/A"
        if page.locator("span.stand").count() > 0:
            location = page.locator("span.stand").first.inner_text().strip()
            
        desc_text = "N/A"
        if page.locator("div.additional").count() > 0:
            raw_desc = page.locator("div.additional").first.inner_text()
            desc_text = raw_desc.split("Categories:")[0].split("Visit us at")[0].strip()
        elif page.locator(".description").count() > 0:
            desc_text = page.locator(".description").first.inner_text().strip()

        repaired = emp_name in self.master_dict
        
        if not repaired:
            self.master_dict[emp_name] = {"Company Name": emp_name}

        self.master_dict[emp_name].update({
            "Stand Number": location,
            "Description": desc_text
        })
        
        return desc_text, repaired, location

    def process_browser(self, records, results):
        """Manages the browser lifecycle and iterates through the dataset for extraction."""
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
            page = context.new_page()

            with Progress(
                SpinnerColumn(), TextColumn("[progress.description]{task.description}"),
                BarColumn(), TaskProgressColumn(), TimeRemainingColumn(), console=console
            ) as progress:
                
                main_task = progress.add_task("[cyan]Extracting & Processing...", total=records)
                progress.update(main_task, completed=len(self.processed_names))

                for item in results:
                    ids = self.get_identificators(item)
                    if not ids: continue
                    path, emp_name, emp_id = ids
                    detail_url = f"https://exhibitors.vitafoods.eu.com/net/company/{path}/result{emp_id}-598.html"
                    
                    try:
                        desc, repaired, loc = self.scrape_profile_data(page, detail_url, emp_name) 

                        status = "[green]REPAIRED[/green]" if repaired else "[blue]NEW[/blue]"
                        preview = desc[:40].replace('\n', ' ') + "..."
                        progress.console.print(f"{status} | {emp_name[:20]:<20} | {loc:<12} | [dim]{preview}[/dim]")

                        # Create a backup every 15 records
                        if len(self.master_dict) % 15 == 0:
                            self.save_data()

                        progress.update(main_task, advance=1)
                        time.sleep(random.uniform(1.2, 2.5))

                    except Exception as e:
                        logging.error(f"Error at {emp_name}: {str(e)}")
                        progress.update(main_task, advance=1)

            browser.close() 

    
def main():
    engine = VitafoodsScraper()

    # 1. Load API data
    data = engine.resolve_json_response()
    results = data.get("results", [])
    records = len(results)

    # 2. Load backup data
    engine.resolve_backup()

    # 3. Display control panel
    user_interaction(records, engine)

    # 4. Start scraping process
    engine.process_browser(records, results)

    # 5. Save and finalize output
    engine.final_dataframe()
    engine.format_report()
    display_result()
    

def display_result():
    """Displays the final completion dashboard."""
    console.print(Panel("[bold green]🏆 TASK COMPLETED: The finalized report has been generated.[/bold green]", expand=False))


def user_interaction(records, engine):
    """Renders the initial user interface console panel."""
    console.print(Panel.fit(
        f"[bold green]VITAFOODS 2026 - DATA EXTRACTION SYSTEM[/bold green]\n"
        f"[blue]Total Results Identified:[/blue] {records}\n"
        f"[yellow]Results with Valid Data:[/yellow] {len(engine.processed_names)}\n"
        f"[dim]Logs stored in: extraction_errors.log[/dim]",
        title="Control Panel"
    ))
    

if __name__ == "__main__":
    main()