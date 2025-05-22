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
    "dept", "department", "inc", "llc", "ltd", "corp", "corporate", "solutions"
}
# Ensure keywords are lowercase for case-insensitive matching
INVALID_NAME_KEYWORDS = {k.lower() for k in INVALID_NAME_KEYWORDS}


def extract_company_from_email(email_address):
    if not isinstance(email_address, str) or '@' not in email_address:
        return None
    domain = email_address.split('@')[-1].lower()
    if domain in GENERIC_EMAIL_DOMAINS:
        return None

    parts = domain.split('.')
    known_sld_parts = {'co', 'com', 'org', 'net', 'ac', 'gov', 'edu', 'mil', 'biz', 'info'}

    name = None
    if len(parts) >= 2:
        if len(parts) > 2 and parts[-2] in known_sld_parts:
            name = parts[-3]
        else:
            name = parts[-2]
    elif len(parts) == 1:
        name = parts[0]

    if name:
        return name.replace('-', ' ').title()
    return None


def clean_name_part(name_part):
    if name_part == "INVALID_NAME_KEYWORD":
        return "INVALID_NAME_KEYWORD"
    if pd.isna(name_part):
        return ""
    name_str = str(name_part).strip()
    if not name_str:
        return ""

    words = name_str.lower().split()
    for w in words:
        cleaned_w = re.sub(r"[.,;:!?]$", "", w)
        cleaned_w = re.sub(r"^[.,;:!?]", "", cleaned_w)
        if cleaned_w in INVALID_NAME_KEYWORDS:
            return "INVALID_NAME_KEYWORD"
    return name_str.title()


def split_full_name_heuristic(full_name_str):
    if not full_name_str or not str(full_name_str).strip():
        return "", ""
    name_as_str = str(full_name_str).strip()
    parts = name_as_str.split()
    if len(parts) == 1:
        return clean_name_part(parts[0]), ""
    first_name_candidate = " ".join(parts[:-1])
    last_name_candidate = parts[-1]
    return clean_name_part(first_name_candidate), clean_name_part(last_name_candidate)


def derive_last_from_email(first_name, email_address):
    if not first_name or first_name == "INVALID_NAME_KEYWORD":
        return ""
    if not isinstance(email_address, str) or '@' not in email_address:
        return ""

    local_part = email_address.split('@')[0].lower()
    local_part = re.sub(r"\d+$", "", local_part)
    fn_lower = first_name.lower()
    fn_escaped = re.escape(fn_lower)

    match = re.match(rf"^{fn_escaped}[._-]([a-zA-Z][a-zA-Z0-9_.-]*)$", local_part)
    if match:
        return clean_name_part(match.group(1))

    match = re.match(rf"^{fn_escaped}([a-zA-Z][a-zA-Z0-9_.-]*)$", local_part)
    if match:
        return clean_name_part(match.group(1))

    match = re.match(rf"^{fn_escaped}(.+)$", local_part)
    if match:
        ln_candidate = match.group(1)
        ln_candidate_cleaned = re.sub(r"^[._-]+", "", ln_candidate)
        if ln_candidate_cleaned and re.search(r"[a-zA-Z]", ln_candidate_cleaned):
            return clean_name_part(ln_candidate_cleaned)
    return ""


def process_excel(input_filepath, output_filepath):
    try:
        df = pd.read_excel(input_filepath, engine='openpyxl')
    except FileNotFoundError:
        print(f"Error: Input file not found at {input_filepath}")
        return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Ensure standard columns exist, initializing them if not (as per original structure)
    managed_cols = ['First Name', 'Last Name', 'Company', 'Email Address', 'LinkedIn URL', 'Processing Status']
    for col in managed_cols:
        if col not in df.columns:
            df[col] = ""
    df['Processing Status'] = "" # Initialize/reset status for this run

    records = []
    for index, row in df.iterrows():
        current_processing_status = ""

        _email_val = row.get('Email Address')
        _fn_val = row.get('First Name')
        _ln_val = row.get('Last Name')
        _comp_val = row.get('Company')

        excel_email = str(_email_val).strip() if not pd.isna(_email_val) else ""
        excel_fn = str(_fn_val).strip() if not pd.isna(_fn_val) else ""
        excel_ln = str(_ln_val).strip() if not pd.isna(_ln_val) else ""
        excel_comp = str(_comp_val).strip() if not pd.isna(_comp_val) else ""

        current_fn, current_ln, current_comp = excel_fn, excel_ln, excel_comp

        # --- Heuristic 0: Detect and attempt to fix swapped Name and Company ---
        if excel_comp and excel_fn and excel_email:
            comp_parts = excel_comp.split()
            looks_like_name_in_comp = (
                len(comp_parts) in [2, 3] and
                all(p.isalpha() or p.replace('.','').isalpha() for p in comp_parts) and
                not any(p.lower().replace('.','') in INVALID_NAME_KEYWORDS for p in comp_parts)
            )
            if looks_like_name_in_comp:
                potential_fn_from_comp = comp_parts[0]
                potential_ln_from_comp = " ".join(comp_parts[1:])
                email_local_part = excel_email.split('@')[0].lower()
                email_local_part_no_digits = re.sub(r"\d+$", "", email_local_part)
                fn_c, ln_c_concat = potential_fn_from_comp.lower(), "".join(potential_ln_from_comp.lower().split())

                name_from_comp_matches_email = False
                patterns_to_check = [
                    f"{fn_c}.{ln_c_concat}", f"{fn_c}_{ln_c_concat}", f"{fn_c}-{ln_c_concat}", f"{fn_c}{ln_c_concat}",
                ]
                if len(fn_c) > 0:
                    patterns_to_check.extend([
                        f"{fn_c[0]}.{ln_c_concat}", f"{fn_c[0]}_{ln_c_concat}", f"{fn_c[0]}-{ln_c_concat}", f"{fn_c[0]}{ln_c_concat}"
                    ])
                if len(ln_c_concat) > 0:
                     patterns_to_check.extend([
                        f"{fn_c}.{ln_c_concat[0]}", f"{fn_c}_{ln_c_concat[0]}", f"{fn_c}-{ln_c_concat[0]}", f"{fn_c}{ln_c_concat[0]}"
                    ])

                if email_local_part_no_digits in patterns_to_check:
                    name_from_comp_matches_email = True
                elif re.sub(r"[._\-]", "", email_local_part_no_digits) == f"{fn_c}{ln_c_concat}": # Fallback
                     name_from_comp_matches_email = True

                orig_fn_first_word_lower = excel_fn.split()[0].lower() if excel_fn else ""
                orig_fn_seems_unrelated = not email_local_part_no_digits.startswith(orig_fn_first_word_lower)

                if name_from_comp_matches_email and (not excel_ln or orig_fn_seems_unrelated):
                    current_comp = excel_fn
                    current_fn = potential_fn_from_comp
                    current_ln = potential_ln_from_comp
                    current_processing_status += 'Swapped Name/Company; '
        # --- End of Heuristic 0 ---

        if not excel_email:
            # df.loc[index, 'Processing Status'] = 'Skipped - Missing Email' # Original style comment
            continue

        final_comp = current_comp.title() if current_comp else extract_company_from_email(excel_email)
        if not final_comp:
            # df.loc[index, 'Processing Status'] = 'Skipped - Missing Company'
            continue

        fn_to_process, ln_to_process = current_fn, current_ln

        if not ln_to_process and ' ' in fn_to_process:
            fn_split, ln_split_from_fn = split_full_name_heuristic(fn_to_process)
            if ln_split_from_fn and ln_split_from_fn != "INVALID_NAME_KEYWORD":
                fn_to_process, ln_to_process = fn_split, ln_split_from_fn
                current_processing_status += 'Split FN to FN/LN; '
        elif fn_to_process and (not ln_to_process or ln_to_process.lower() == fn_to_process.lower() or
                     (fn_to_process.split() and len(fn_to_process.split()) > 1 and ln_to_process.lower() == fn_to_process.split()[-1].lower())):
            fn_split, ln_split_from_fn = split_full_name_heuristic(fn_to_process)
            if ln_split_from_fn and ln_split_from_fn != "INVALID_NAME_KEYWORD":
                fn_to_process, ln_to_process = fn_split, ln_split_from_fn
                current_processing_status += 'Split full name from FN; '
        
        final_fn = clean_name_part(fn_to_process)
        final_ln = clean_name_part(ln_to_process)

        if final_fn and final_fn != "INVALID_NAME_KEYWORD" and (not final_ln or final_ln == "INVALID_NAME_KEYWORD"):
            if final_ln == "INVALID_NAME_KEYWORD": final_ln = ""
            derived_ln = derive_last_from_email(final_fn, excel_email)
            if derived_ln and derived_ln != "INVALID_NAME_KEYWORD":
                final_ln = derived_ln
                current_processing_status += 'Derived LN from Email; '

        if not final_fn or final_fn == "INVALID_NAME_KEYWORD" or \
           not final_ln or final_ln == "INVALID_NAME_KEYWORD":
            # df.loc[index, 'Processing Status'] = 'Skipped - Invalid/Missing Name Parts'
            continue

        processed_record = dict(row) # Start with a copy of the original row
        processed_record['First Name'] = final_fn
        processed_record['Last Name'] = final_ln
        processed_record['Company'] = final_comp
        processed_record['Email Address'] = excel_email # Ensure original, stripped email is used
        # LinkedIn URL is preserved from input if it exists via dict(row).
        # If it wasn't in input, it was initialized to "" by the loop at the start.
        # The original example specifically set LinkedIn URL = '', so if that's needed:
        # processed_record['LinkedIn URL'] = '' 
        # For now, we preserve or use the initialized empty string from df.
        
        current_processing_status += 'Processed'
        processed_record['Processing Status'] = current_processing_status.strip().rstrip(';')
        records.append(processed_record)

    out_df = pd.DataFrame(records)

    if not out_df.empty:
        output_dir = os.path.dirname(output_filepath)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        try:
            # If you want only specific columns in the output, uncomment the next line and define `output_cols_final`
            # output_cols_final = ['First Name', 'Last Name', 'Company', 'Email Address', 'LinkedIn URL', 'Processing Status']
            # out_df = out_df[output_cols_final] 
            out_df.to_excel(output_filepath, index=False, engine='openpyxl')
            print(f"Processed {len(df)} original rows. Written {len(out_df)} valid rows to {output_filepath}")
        except Exception as e:
            print(f"Error writing output file to {output_filepath}: {e}")
    else:
        print(f"Processed {len(df)} original rows. No valid rows to write to {output_filepath}")


if __name__ == '__main__':
    '''*************************************************************************************************************************************************'''
    ''' Go ahead and add the input xlsx and output xlsx over here and ensure that there is a column called First Name Last Name Company and Email Address'''
    '''*************************************************************************************************************************************************'''

    inp = r'C:\Users\prana\300rows.xlsx' # Use raw string for Windows paths or escape backslashes
    outp = 'cleandoutput.xlsx'
    process_excel(inp, outp)