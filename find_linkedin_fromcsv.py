import pandas as pd
import stealth_requests as requests 
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
import time
import random
import re
import os


USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
]

def get_random_user_agent():
    return random.choice(USER_AGENTS)

def make_bing_site_specific_search_url(first_name, last_name, company_name):
    fn = first_name if first_name else ""
    ln = last_name if last_name else ""
    cn = company_name if company_name else ""
    query_parts = [f'site:linkedin.com/in/']
    if fn or ln: 
        query_parts.append(f'"{fn.strip()} {ln.strip()}".strip()')
    if cn:
        query_parts.append(f'"{cn.strip()}"')
    query = " ".join(filter(None, query_parts)) 
    encoded_query = quote_plus(query)
    return f"https://www.bing.com/search?q={encoded_query}&ensearch=1"

def make_bing_general_search_url(first_name, last_name, company_name):
    fn = first_name if first_name else ""
    ln = last_name if last_name else ""
    cn = company_name if company_name else ""
    query_parts = []
    if fn or ln:
        query_parts.append(f'"{fn.strip()} {ln.strip()}".strip()')
    if cn:
        query_parts.append(f'"{cn.strip()}"')
    query_parts.append('LinkedIn')
    query = " ".join(filter(None, query_parts))
    encoded_query = quote_plus(query)
    return f"https://www.bing.com/search?q={encoded_query}&ensearch=1"

def make_bing_email_search_url(email_address):
    query = f'"{email_address.strip()}" LinkedIn profile'
    encoded_query = quote_plus(query)
    return f"https://www.bing.com/search?q={encoded_query}&ensearch=1"

def extract_all_linkedin_links_from_bing(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    links = set()
    for a_tag in soup.find_all("a", href=True):
        href = a_tag['href']
        if "linkedin.com/" in href:
            if href.startswith("http://www.linkedin.com/") or href.startswith("https://www.linkedin.com/"):
                clean_url = href.split('?')[0]
                if clean_url.endswith('/'):
                   clean_url = clean_url[:-1]
                links.add(clean_url)
    return list(links)

def derive_profile_from_activity_url(activity_url):
    post_match_alt = re.search(r"linkedin\.com/posts/([^/_]+)_", activity_url)
    profile_segment = None
    if post_match_alt:
        profile_segment = post_match_alt.group(1)
    if profile_segment:
        return f"https://www.linkedin.com/in/{profile_segment}"
    return None

def _make_bing_request(search_url):
    headers = {
        "User-Agent": get_random_user_agent(),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Referer": "https://www.bing.com/", "DNT": "1", "Upgrade-Insecure-Requests": "1"
    }
    try:
        response = requests.get(search_url, headers=headers, timeout=25) 
        response.raise_for_status() 
        if "CAPTCHA" in response.text or "verify that you're not a robot" in response.text.lower() or "traffic from your computer network" in response.text.lower():
            print(f"Bing is likely blocking or asking for CAPTCHA for URL: {search_url}")
            return None
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Request error for {search_url}: {e}") 
        return None
    except Exception as e:
        print(f"An unexpected error occurred during request for {search_url}: {e}")
        return None

def _process_bing_links(found_links_list):
    direct_profile_urls = []
    derived_profile_urls = set()
    for link in found_links_list:
        if link.startswith("https://www.linkedin.com/in/"):
            if re.match(r"https://www\.linkedin\.com/in/[a-zA-Z0-9-]+(?:-[a-zA-Z0-9-]+)*$", link):
                direct_profile_urls.append(link)
        elif link.startswith("https://www.linkedin.com/posts/"):
            derived_url = derive_profile_from_activity_url(link)
            if derived_url:
                derived_profile_urls.add(derived_url)

    if direct_profile_urls:
        return direct_profile_urls[0]
    if derived_profile_urls:
        derived_list = list(derived_profile_urls)
        return derived_list[0]
    return None

def find_linkedin_profile_via_bing(first_name, last_name, company_name, email_address=None):
    profile_url = None
    if first_name and last_name and company_name:
        print(f"  Attempt 1 (Specific): '{first_name}' '{last_name}' at '{company_name}'")
        specific_search_url = make_bing_site_specific_search_url(first_name, last_name, company_name)
        html_content_specific = _make_bing_request(specific_search_url)
        if html_content_specific:
            profile_url = _process_bing_links(extract_all_linkedin_links_from_bing(html_content_specific))
            if profile_url: return profile_url
    
    if not profile_url and (first_name or last_name) and company_name:
        print(f"  Attempt 2 (General): '{first_name}' '{last_name}' at '{company_name}'")
        general_search_url = make_bing_general_search_url(first_name, last_name, company_name)
        html_content_general = _make_bing_request(general_search_url)
        if html_content_general:
            profile_url = _process_bing_links(extract_all_linkedin_links_from_bing(html_content_general))
            if profile_url: return profile_url

    if not profile_url and email_address:
        print(f"  Attempt 3 (Email): '{email_address}'")
        email_search_url = make_bing_email_search_url(email_address)
        html_content_email = _make_bing_request(email_search_url)
        if html_content_email:
            profile_url = _process_bing_links(extract_all_linkedin_links_from_bing(html_content_email))
            if profile_url: return profile_url
            
    return None
# --- End of LinkedIn Profile Finder Script ---


# --- Main Application Logic ---
if __name__ == "__main__":
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    ''' Go ahead and add the input xlsx and output xlsx over here and ensure that there is a column called First Name Last Name Company and Email Address'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''

    input_cleaned_xlsx_file = "output_cleaned_v3.xlsx" 
    output_final_xlsx_file = "output_with_linkedin_urls.xlsx" # Modified output filename
    
    # Define a sensible delay between searches
    MIN_DELAY = 1.0 # seconds
    MAX_DELAY = 2.0 # seconds 
    # Number of rows to process
    ROWS_TO_PROCESS = 211 #Just change to number of rows you wanna process

    try:
        df_full = pd.read_excel(input_cleaned_xlsx_file, engine='openpyxl')
    except FileNotFoundError:
        print(f"Error: Input file '{input_cleaned_xlsx_file}' not found.")
        exit()
    except Exception as e:
        print(f"Error reading Excel file '{input_cleaned_xlsx_file}': {e}")
        exit()

    # Take only the first N rows for processing
    df_to_process = df_full.head(ROWS_TO_PROCESS).copy() # Use .copy() to avoid SettingWithCopyWarning
    print(f"Processing the first {len(df_to_process)} rows of the input file.")


    if 'LinkedIn URL' not in df_to_process.columns:
        df_to_process['LinkedIn URL'] = "NA"
    else:
        df_to_process['LinkedIn URL'] = df_to_process['LinkedIn URL'].astype(str).fillna("NA").replace(['nan', 'None'], "NA", regex=False)

    print(f"Starting LinkedIn profile search for {len(df_to_process)} selected rows...")

    for index, row_series in df_to_process.iterrows():

        row = {k: str(v).strip() if pd.notna(v) else "" for k, v in row_series.to_dict().items()}

        first_name = row.get('First Name', "")
        last_name = row.get('Last Name', "")
        company = row.get('Company', "")
        email = row.get('Email Address', "")
        
        processing_status = row.get('Processing Status', '')
        is_emptied_row = "Emptied:" in processing_status or (not first_name and not last_name and not company and not email)

        if is_emptied_row:
            print(f"Row (original index {index}, processed index {df_to_process.index.get_loc(index) +1}): Skipping search (marked as emptied or no key data). LinkedIn URL set to NA.")
            df_to_process.loc[index, 'LinkedIn URL'] = "NA" # Update the sliced DataFrame
            continue

        if not (first_name and last_name and company) and not email:
            print(f"Row (original index {index}, processed index {df_to_process.index.get_loc(index) +1}): Insufficient data for search. LinkedIn URL set to NA.")
            df_to_process.loc[index, 'LinkedIn URL'] = "NA"
            continue

        print(f"\nProcessing Row (original index {index}, processed index {df_to_process.index.get_loc(index) +1}): Name='{first_name} {last_name}', Co='{company}', Email='{email}'")
        
        current_delay = random.uniform(MIN_DELAY, MAX_DELAY)
        
        profile_url_found = find_linkedin_profile_via_bing(
            first_name, 
            last_name,
            company,
            email
        )

        if profile_url_found:
            print(f"  >>> SUCCESS: Found LinkedIn URL: {profile_url_found}")
            df_to_process.loc[index, 'LinkedIn URL'] = profile_url_found
        else:
            print(f"  >>> FAILURE: Could not find LinkedIn profile for this row.")
            df_to_process.loc[index, 'LinkedIn URL'] = "NA"
        
        # Don't sleep after the last item in the subset
        if df_to_process.index.get_loc(index) < len(df_to_process) - 1:
            print(f"Waiting for {current_delay:.2f} seconds before next search...")
            time.sleep(current_delay)


    output_dir = os.path.dirname(output_final_xlsx_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
        
    try:
        df_to_process.to_excel(output_final_xlsx_file, index=False, engine='openpyxl')
        print(f"\nLinkedIn URL enrichment for the first {len(df_to_process)} rows complete. Output written to: {output_final_xlsx_file}")
    except Exception as e:
        print(f"Error writing to final Excel file '{output_final_xlsx_file}': {e}")

    print("\n--- Script finished ---")
