import asyncio
import pandas as pd
import re
import os
from datetime import datetime
from crawl4ai import AsyncWebCrawler, CrawlerRunConfig, CacheMode
from openpyxl.styles import Alignment # Critical for the auto-formatting

def pro_clean(text):
    """
    Cleans markdown into a professional 2-sentence summary.
    """
    if not text: 
        return "Data Unavailable"
    
    # Remove markdown noise and URLs
    text = re.sub(r'\(http[^\)]+\)|[\*\[\]\|!\-\_#]', '', text)
    sentences = [s.strip() for s in text.split('.') if len(s) > 15]
    
    summary = ". ".join(sentences[:2])
    final_text = " ".join(summary.split())
    
    return (final_text[:247] + '...') if len(final_text) > 250 else final_text

async def main():
    print("🚀 Starting Professional Dual-Export with Auto-Formatting...")
    
    # 1. Setup Paths (Matching your specific file names)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(base_dir, 'targets.txt')
    csv_output = os.path.join(base_dir, 'Professional_Leads_Report.csv')
    xlsx_output = os.path.join(base_dir, 'Professional_Leads_Report.xlsx')

    if not os.path.exists(input_path):
        print(f"❌ Error: 'targets.txt' not found at {input_path}")
        return

    leads = []
    
    # 2. AI Scouting Phase
    async with AsyncWebCrawler() as crawler:
        with open(input_path, 'r') as f:
            urls = [line.strip() for line in f if line.strip()]

        for url in urls:
            print(f"🔭 Researching: {url}...")
            try:
                result = await crawler.arun(
                    url=url, 
                    config=CrawlerRunConfig(cache_mode=CacheMode.ENABLED)
                )
                
                if result.success:
                    leads.append({
                        "Company": url,
                        "Status": "Verified",
                        "Scan Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "Executive Summary": pro_clean(result.markdown)
                    })
                else:
                    leads.append({
                        "Company": url, 
                        "Status": "Failed", 
                        "Scan Date": datetime.now().strftime("%Y-%m-%d %H:%M"), 
                        "Executive Summary": "Site unreachable."
                    })
            except Exception as e:
                print(f"⚠️ Error scouting {url}: {e}")

    # 3. Professional Dual Export & Styling
    if leads:
        df = pd.DataFrame(leads)
        
        # Save the CSV (for GitHub)
        df.to_csv(csv_output, index=False, encoding='utf-8-sig')
        
        # Save and Format the XLSX (for readable reporting)
        try:
            with pd.ExcelWriter(xlsx_output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Leads')
                worksheet = writer.sheets['Leads']

                # Set Column Widths
                worksheet.column_dimensions['A'].width = 35 # Company
                worksheet.column_dimensions['B'].width = 12 # Status
                worksheet.column_dimensions['C'].width = 20 # Scan Date
                worksheet.column_dimensions['D'].width = 65 # Executive Summary

                # Apply Wrap Text and Top Alignment to every data cell
                for row in worksheet.iter_rows(min_row=2, max_row=len(leads)+1, max_col=4):
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
            
            print(f"\n✅ SUCCESS: Updated and Formatted {xlsx_output}")
            print(f"✅ SUCCESS: Updated {csv_output}")
        
        except PermissionError:
            print(f"❌ ERROR: Could not save XLSX. Please CLOSE {xlsx_output} in Excel and try again.")
    else:
        print("❌ No data was collected.")

if __name__ == "__main__":
    asyncio.run(main())