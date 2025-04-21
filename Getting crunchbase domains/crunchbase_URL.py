import os
import time
import random
import pandas as pd
import requests
from bs4 import BeautifulSoup
from duckduckgo_search import DDGS # type: ignore

def search_crunchbase_url(company_name, potential_URLs):
    try:
        crunchbase_urls = [url for url in potential_URLs if 'crunchbase.com' in url.lower()]
        if crunchbase_urls:
            return crunchbase_urls[0]
        return "Not Found"
    except Exception as e:
        print(f"Error processing URLs for {company_name}: {str(e)}")
        return None

def search_with_retry(query, max_retries=3):
    for attempt in range(max_retries):
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=5))
                return [r['href'] for r in results]
        except Exception as e:
            print(f"DDGS attempt {attempt + 1} failed: {str(e)}")
            if attempt == max_retries - 1:
                try:
                    headers = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                    }
                    url = f"https://html.duckduckgo.com/html/?q={requests.utils.quote(query)}"
                    response = requests.get(url, headers=headers, timeout=10)
                    soup = BeautifulSoup(response.text, 'html.parser')
                    return [a['href'] for a in soup.select('a.result__url')][:5]
                except Exception as fallback_error:
                    print(f"Fallback search failed: {str(fallback_error)}")
                    return []
            time.sleep(random.uniform(2, 5))
    return []

def find_crunchbase_urls(start_row, end_row):
    dataframe = pd.read_excel(INPUT_FILE)
    companies = dataframe.iloc[start_row:end_row][COMPANY_COLUMN].tolist()

    if os.path.exists(OUTPUT_FILE):
        existing_data = pd.read_excel(OUTPUT_FILE)
    else:
        existing_data = pd.DataFrame(columns=["Company Name", "Crunchbase URL"])

    results = []
    for i, company in enumerate(companies, start=1):
        print(f"Processing {i}/{len(companies)}: {company}")
        
        query = f"{company} site:crunchbase.com"
        potential_urls = search_with_retry(query)
        
        if not potential_urls:
            print(f"No results found for {company}")
            results.append({"Company Name": company, "Crunchbase URL": "Not Found"})
            continue
            
        crunchbase_url = search_crunchbase_url(company, potential_urls)
        results.append({"Company Name": company, "Crunchbase URL": crunchbase_url})
        
        pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)
        time.sleep(random.uniform(1, 3))

    return pd.DataFrame(results)

if __name__ == "__main__":
    INPUT_FILE = "C:\\JMAN Final Project - Third Party websites\\filtered_company_list.xlsx"
    OUTPUT_FILE = "C:\\JMAN Final Project - Third Party websites\\Getting crunchbase domains\\company_name_with_crunchbase_links.xlsx"
    COMPANY_COLUMN = "Company Name"

    start_row = int(input("Enter start row (0-based): "))
    end_row = int(input("Enter end row (exclusive): "))

    final_data = find_crunchbase_urls(start_row, end_row)
    final_data.to_excel(OUTPUT_FILE, index=False)
    print(f"Completed! Results saved to {OUTPUT_FILE}")