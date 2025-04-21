import os
import time
import random
import pandas as pd
from duckduckgo_search import DDGS  # type: ignore

def search_crunchbase_url(company_name, potential_URLs):
    try:
        crunchbase_urls = [url for url in potential_URLs if 'crunchbase.com' in url.lower()]
        if crunchbase_urls:
            return crunchbase_urls[0]
        return "Not Found"
    except Exception as e:
        print(f"Error processing URLs for {company_name}: {str(e)}")
        return None

def search_duckduckgo(query, num_results):
    try:
        with DDGS() as ddgs:
            search_results = list(ddgs.text(query, max_results=num_results))
            return [search_result['href'] for search_result in search_results]
    except Exception as e:
        print(f"Error searching DuckDuckGo: {str(e)}")
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
        potential_urls = search_duckduckgo(query, num_results=3)
        
        if not potential_urls:
            print(f"No results found for {company}")
            results.append({"Company Name": company, "Crunchbase URL": "Not Found"})
            continue
        
        crunchbase_url = search_crunchbase_url(company, potential_urls)
        results.append({"Company Name": company, "Crunchbase URL": crunchbase_url})
        
        updated_data = pd.DataFrame(results)
        combined_data = pd.concat([existing_data, updated_data], ignore_index=True)
        combined_data.to_excel(OUTPUT_FILE, index=False)
        print(f"Data saved to {OUTPUT_FILE}")
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
