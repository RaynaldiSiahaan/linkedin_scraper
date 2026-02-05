import asyncio
import random

import os
import sys
from pathlib import Path
try:
    import pandas as pd
except ImportError:
    print("Error: pandas is required. Please install it with: pip install pandas openpyxl")
    sys.exit(1)

from linkedin_scraper.scrapers.person import PersonScraper
from linkedin_scraper.core.browser import BrowserManager


async def main():
    """Scrape person profiles from Excel list"""
    # Define file path
    root_dir = Path(__file__).parent.parent
    excel_path = root_dir / "Alumni Shared.xlsx"
    
    if not excel_path.exists():
        print(f"Error: Could not find {excel_path}")
        return

    print(f"Reading data from: {excel_path}")
    
    try:
        # Read Excel file - Sheet '2013'
        df = pd.read_excel(excel_path, sheet_name="2013")
        
        # Filter: Conclusion == "Yes"
        # Using Fillna to handle potential NaN values safely
        candidates = df[df["Conclusion"].fillna("").astype(str).str.lower() == "yes"].copy()
        
        print(f"Found {len(candidates)} candidates to scrape from 2013 sheet.")
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    if candidates.empty:
        print("No candidates found with Conclusion = 'Yes'.")
        return

    # Initialize and start browser using context manager
    async with BrowserManager(headless=False,
        args=[
        '--disable-dev-shm-usage',
        '--disable-blink-features=AutomationControlled',
        '--disable-web-security',  # Careful with this
        '--disable-features=IsolateOrigins,site-per-process',
        '--no-sandbox',
    ]) as browser:
        # Load existing session (must be created first - see README for setup)
        try:
            await browser.load_session("linkedin_session.json")
            print("✓ Session loaded")
        except Exception as e:
            print(f"Warning: Could not load session ({e}). You might need to login manually.")

        # Initialize scraper with the browser page
        scraper = PersonScraper(browser.page)
        
        for index, row in candidates.iterrows():
            name = row.get("Nama Mahasiswa")
            linkedin_url = row.get("LinkedIn")
            
            if not linkedin_url or pd.isna(linkedin_url):
                print(f"Skipping {name}: No LinkedIn URL found")
                continue
                
            # Clean up URL if needed (remove query params etc usually handled by scraper but good to ensure string)
            linkedin_url = str(linkedin_url).strip()
            
            print(f"\n[{index+1}] Processing: {name}")
            print(f"🚀 Scraping: {linkedin_url}")
            
            try:
                # Give page more time to settle

                person = await scraper.scrape(linkedin_url)
                await asyncio.sleep(15)

                print(f"✓ Success! Current URL: {browser.page.url}")

                # Display person info
                print("-" * 40)
                print(f"Person: {person.name}")
                print(f"Location: {person.location} ")
                print(f"About: {person.about[:100] if person.about else 'N/A'}...")
                
                # Display work experience
                print(f"Work Experience ({len(person.experiences)} positions):")
                for exp in person.experiences[:2]:  # Show first 2
                    print(f"  - {exp.position_title} at {exp.institution_name}")
                
                 # Display education
                print(f"Education ({len(person.educations)} schools):")
                for edu in person.educations[:2]:  # Show first 2
                    print(f"  - {edu.institution_name}")
                print("-" * 40)
                
                # Optional: Add a small delay between requests to be safe
                await asyncio.sleep(random.uniform(2.2, 4.313))
                
            except Exception as e:
                print(f"❌ Failed to scrape {name}")
                print(f"🔴 Error type: {type(e).__name__}")
                print(f"🔴 Error: {str(e)}")
                print(f"🌐 Current URL in browser: {browser.page.url}")
                
                # Take a screenshot to see what's on screen
                await browser.page.screenshot(path=f"error_{index}.png")
                print(f"📸 Screenshot saved as error_{index}.png")
                
                continue
    
    print("\n✓ Batch processing done!")


if __name__ == "__main__":
    asyncio.run(main())
