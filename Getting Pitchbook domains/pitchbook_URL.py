import pandas as pd
import requests
from bs4 import BeautifulSoup

def fetch_pitchbook_link(company_name):
    try:
        query = f"pitchbook - {company_name}"
        url = f"https://duckduckgo.com/html/?q={query}"
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        link = soup.find('a', {'class': 'result__url'})
        return link['href'] if link else "Not Found"
    except:
        return "Error"

def main():
    input_file = "C:\\JMAN Final Project - Third Party websites\\filtered_company_list.xlsx"
    output_file = "C:\\JMAN Final Project - Third Party websites\\Getting Pitchbook domains\\company_name_with_pitchbook_links.xlsx"
    
    start_row = int(input("Enter start row (0-based indexing): "))
    end_row = int(input("Enter end row: "))
    
    df = pd.read_excel(input_file)
    selected_rows = df.iloc[start_row:end_row]
    
    results = []
    for index, row in selected_rows.iterrows():
        company_name = row.iloc[1]
        pitchbook_link = fetch_pitchbook_link(company_name)
        print(f"fetching pitchbook URL for {company_name}")
        results.append({'Company Name': company_name, 'Pitchbook Link': pitchbook_link})
    
    result_df = pd.DataFrame(results)
    result_df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

if __name__ == "__main__":
    main()