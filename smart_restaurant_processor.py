import pandas as pd
import phonenumbers
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.numbers import FORMAT_TEXT
from openpyxl.utils import get_column_letter
import os
import time
from tqdm import tqdm
import platform

# File paths
INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"

# Global state
has_changed = False
instagram_found_count = 0
instagram_fallback_count = 0
duplicates_removed = 0
partial_duplicates = []
phone_format_count = 0
ran_actions = False

# Formatting
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
bold_font = Font(bold=True)


def clear_screen():
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')


def load_data():
    try:
        df = pd.read_excel(INPUT_FILE, dtype=str)
        df.columns = [col.strip().replace("<", "").replace(">", "") for col in df.columns]
        df = df.dropna(how="all")
        return df
    except Exception as e:
        print(f"❌ Failed to load input file: {e}")
        exit(1)


def save_data(df):
    df = df.sort_values(by="COMPANY_name")
    try:
        if os.path.exists(OUTPUT_FILE):
            os.remove(OUTPUT_FILE)

        df.to_excel(OUTPUT_FILE, index=False)
        wb = openpyxl.load_workbook(OUTPUT_FILE)
        ws = wb.active

        phone_col_idx = list(df.columns).index("Company_Phone") + 1
        link_col_idx = list(df.columns).index("Instagram_link") + 1

        # Format phone column as text
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=phone_col_idx)
            cell.number_format = FORMAT_TEXT

        # Make Instagram links clickable
        for row in range(2, ws.max_row + 1):
            link = ws.cell(row=row, column=link_col_idx).value
            if link:
                ws.cell(row=row, column=link_col_idx).value = f'=HYPERLINK("{link}", "Instagram")'

        # Highlight partial duplicates
        for dup in partial_duplicates:
            for row in range(2, ws.max_row + 1):
                if ws[f"A{row}"].value == dup['name'] or ws[f"{get_column_letter(phone_col_idx)}{row}"].value == dup['phone']:
                    cell = ws[f"{get_column_letter(phone_col_idx)}{row}"]
                    cell.font = bold_font
                    cell.fill = red_fill

        wb.save(OUTPUT_FILE)
    except Exception as e:
        print(f"❌ Failed to write output file: {e}")
        exit(1)


def normalize_phone(df):
    global phone_format_count, has_changed

    def format_number(num):
        try:
            num = str(num).strip().replace(".0", "")
            digits = ''.join(filter(str.isdigit, num))
            if len(digits) == 10:
                digits = '1' + digits
            if not digits.startswith('1'):
                digits = '1' + digits
            parsed = phonenumbers.parse("+" + digits, 'US')
            if phonenumbers.is_valid_number(parsed):
                return phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
        except:
            return num
        return num

    df['Company_Phone'] = df['Company_Phone'].astype(str).apply(lambda x: format_number(x))
    phone_format_count = len(df)
    has_changed = True
    print(f"✔ {phone_format_count} phone numbers formatted.")
    return df


def remove_duplicates(df):
    global duplicates_removed, partial_duplicates, has_changed

    before = len(df)
    df.drop_duplicates(subset=['COMPANY_name', 'Company_Phone'], inplace=True)
    duplicates_removed = before - len(df)

    # Detect partial duplicates
    name_dups = df[df.duplicated(['COMPANY_name'], keep=False)]
    phone_dups = df[df.duplicated(['Company_Phone'], keep=False)]

    for _, row in name_dups.iterrows():
        partial_duplicates.append({'name': row['COMPANY_name'], 'phone': row['Company_Phone']})
    for _, row in phone_dups.iterrows():
        partial_duplicates.append({'name': row['COMPANY_name'], 'phone': row['Company_Phone']})

    has_changed = True
    print(f"✔ {duplicates_removed} exact duplicates removed.")
    print(f"⚠ {len(partial_duplicates)} partial duplicates found and marked.")
    return df


def generate_instagram_links(df):
    global instagram_found_count, instagram_fallback_count, has_changed

    print("🔍 Generating Instagram Google search links...")
    links = []
    for name in tqdm(df['COMPANY_name'], desc="Generating Instagram Google Search Links"):
        query = '+'.join(name.split())
        link = f"https://www.google.com/search?q={query}+restaurant+hawaii+site:instagram.com"
        links.append(link)
        instagram_fallback_count += 1

    df['Instagram_link'] = links
    has_changed = True
    print(f"🔗 {instagram_fallback_count} Instagram Google search links generated.")
    return df


def main():
    global has_changed, ran_actions
    clear_screen()
    df = load_data()

    while True:
        print("\nWhat would you like to do?")
        print("[Press Enter] Run all smart functions")
        print("[1] Generate Instagram links only")
        print("[2] Remove duplicates only")
        print("[3] Improve phone number formatting only")
        print("[0] Export results to output.xlsx (requires at least one action)")

        choice = input("> ").strip()

        if choice == "":
            df = remove_duplicates(df)
            df = generate_instagram_links(df)
            df = normalize_phone(df)
            ran_actions = True
        elif choice == "1":
            df = generate_instagram_links(df)
            ran_actions = True
        elif choice == "2":
            df = remove_duplicates(df)
            ran_actions = True
        elif choice == "3":
            df = normalize_phone(df)
            ran_actions = True
        elif choice == "0":
            if has_changed or ran_actions:
                save_data(df)
                print(f"✔ Output written to {OUTPUT_FILE}")
                break
            else:
                print("⚠ You must perform at least one action before exporting.")
        else:
            print("Invalid option. Try again.")


if __name__ == '__main__':
    main()