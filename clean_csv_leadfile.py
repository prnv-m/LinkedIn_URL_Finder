import pandas as pd
import re
import os

# --- Configuration (Keep GENERIC_EMAIL_DOMAINS and INVALID_NAME_KEYWORDS) ---
GENERIC_EMAIL_DOMAINS = {
    "gmail.com", "outlook.com", "yahoo.com", "hotmail.com", "aol.com", 
    "icloud.com", "me.com", "msn.com", "live.com", "protonmail.com",
    "zoho.com", "yandex.com", "gmx.com", "mail.com"
}
INVALID_NAME_KEYWORDS = {
    "admin", "administrator", "ceo", "cfo", "cto", "coo", "contact", "careers",
    "director", "desk", "exports", "enquiries", "enquiry", "executive", "finance", 
    "group", "help", "hr", "info", "information", "inquiries", "inquiry", "jobs", 
    "legal", "manager", "management", "manufacturing", "marketing", "media", 
    "office", "operations", "president", "press", "promo", "promotions", "purchase", 
    "recruitment", "sales", "secretary", "service", "services", "support", "team", 
    "technical", "technology", "webmaster", "website", "accounts", "billing",
}


def extract_company_from_email(email_address):
    if not isinstance(email_address, str) or '@' not in email_address:
        return None
    domain = email_address.split('@')[-1].lower()
    if domain in GENERIC_EMAIL_DOMAINS:
        return None
    parts = domain.split('.')
    # derive main domain segment
    if len(parts) > 2 and parts[-2] in ['co', 'com', 'org', 'net', 'ac', 'gov', 'edu']:
        name = parts[-3]
    else:
        name = parts[0]
    return name.replace('-', ' ').title()


def clean_name_part(name_part):
    if pd.isna(name_part) or not str(name_part).strip():
        return ""
    name_str = str(name_part).strip()
    words = name_str.lower().split()
    for w in words:
        if w in INVALID_NAME_KEYWORDS:
            return "INVALID_NAME_KEYWORD"
    return name_str.title()


def split_full_name_heuristic(full_name_str):
    if pd.isna(full_name_str) or not str(full_name_str).strip():
        return "", ""
    parts = str(full_name_str).strip().split()
    if len(parts) == 1:
        return clean_name_part(parts[0]), ""
    first = " ".join(parts[:-1])
    last = parts[-1]
    return clean_name_part(first), clean_name_part(last)


def derive_last_from_email(first_name, email):
    local = email.split('@')[0].lower()
    # strip trailing digits
    local = re.sub(r"\d+$", "", local)
    fn = re.escape(first_name.lower())
    m = re.match(rf"^{fn}(.+)$", local)
    if m:
        ln_candidate = m.group(1)
        return clean_name_part(ln_candidate)
    return ""


def process_excel(input_filepath, output_filepath):
    df = pd.read_excel(input_filepath, engine='openpyxl')
    # ensure status column
    df['Processing Status'] = ""

    records = []
    for _, row in df.iterrows():
        email = str(row.get('Email Address', '')).strip()
        orig_fn = str(row.get('First Name', '')).strip()
        orig_ln = str(row.get('Last Name', '')).strip()
        orig_comp = str(row.get('Company', '')).strip()

        status = 'OK'
        # Validate email
        if not email:
            continue
        # Company
        comp = orig_comp.title() if orig_comp else extract_company_from_email(email)
        if not comp:
            continue
        # Name splitting/cleaning
        fn, ln = orig_fn, orig_ln
        # split if last empty and first has space
        if not ln and ' ' in fn:
            fn, ln = split_full_name_heuristic(fn)
        # else full-name in first or duplicate
        elif fn and (not ln or ln.lower() == fn.lower() or (fn.split() and ln.lower() == fn.split()[-1].lower())):
            fsp, lsp = split_full_name_heuristic(fn)
            if lsp:
                fn, ln = fsp, lsp
        # clean parts
        fn = clean_name_part(fn)
        ln = clean_name_part(ln)
        # derive from email if still missing ln
        if fn and not ln:
            derived = derive_last_from_email(fn, email)
            if derived:
                ln = derived
        # reject invalid or missing ln
        if ln == 'INVALID_NAME_KEYWORD' or fn == 'INVALID_NAME_KEYWORD' or not ln:
            continue
        # build processed record
        record = dict(row)
        record['First Name'] = fn
        record['Last Name'] = ln
        record['Company'] = comp
        record['LinkedIn URL'] = ''
        record['Processing Status'] = 'Processed'
        records.append(record)

    # create new DataFrame of valid rows only
    out_df = pd.DataFrame(records)
    os.makedirs(os.path.dirname(output_filepath) or '.', exist_ok=True)
    out_df.to_excel(output_filepath, index=False, engine='openpyxl')
    print(f"Written {len(out_df)} valid rows to {output_filepath}")


if __name__ == '__main__':
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    ''' Go ahead and add the input xlsx and output xlsx over here and ensure that there is a column called First Name Last Name Company and Email Address'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''
    '''*************************************************************************************************************************************************'''

    inp = r'C:\Users\prana\300rows.xlsx'
    outp = 'myoutput.xlsx'
    process_excel(inp, outp)