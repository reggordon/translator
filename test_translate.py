#!/usr/bin/env python3
"""
Test translation script for validating Excel files with the main translation logic.
Translates the first 3 rows of the selected file and prints the results for review.
"""
import os
import sys
import pandas as pd
from deep_translator import GoogleTranslator, LibreTranslator
import re
import string

def list_excel_files():
    return [f for f in os.listdir('.') if f.endswith('.xlsx')]

def choose_file(files, prompt):
    print(f"\n{prompt}")
    for idx, f in enumerate(files):
        print(f"  [{idx+1}] {f}")
    while True:
        choice = input("Enter number: ")
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice)-1]
        print("Invalid choice. Try again.")

def main():
    print("\n=== Test Translation Script ===")
    excel_files = list_excel_files()
    if not excel_files:
        print("No .xlsx files found in current directory.")
        sys.exit(1)
    input_file = choose_file(excel_files, "Select input Excel file:")
    # Ask user for terms to ignore (comma-separated, case-insensitive)
    ignore_terms_input = input("Enter comma-separated terms to ignore (leave blank for none): ").strip()
    if ignore_terms_input:
        ignore_terms = [t.strip() for t in ignore_terms_input.split(",") if t.strip()]
    else:
        ignore_terms = []
    # Preprocess input for clarity: clean whitespace, remove unnecessary punctuation, standardize casing
    def preprocess_text(text):
        text = ' '.join(str(text).split())
        allowed_punct = set('.!?,-:;')
        text = ''.join(ch for ch in text if ch.isalnum() or ch.isspace() or ch in allowed_punct)
        if text:
            text = text[0].upper() + text[1:]
        return text

    def extract_bold(text):
        bold_pattern = r"(\*\*([^*]+)\*\*)"
        placeholders = {}
        new_text = text
        bolds = re.findall(bold_pattern, new_text)
        for i, (full, inner) in enumerate(bolds):
            matched_ignore = None
            for term in ignore_terms:
                if inner.strip().lower() == term.lower():
                    matched_ignore = term
                    break
            if matched_ignore:
                placeholder = f"__IGNORE_{i}__"
                placeholders[placeholder] = f"[{matched_ignore}](#)"
                new_text = new_text.replace(full, placeholder, 1)
            else:
                placeholder = f"__BOLD_{i}__"
                placeholders[placeholder] = f"[{inner}](#)"
                new_text = new_text.replace(full, placeholder, 1)
        def ignore_replacer_factory(term):
            def ignore_replacer(match):
                idx = len([k for k in placeholders if k.startswith("__IGNORE_")])
                placeholder = f"__IGNORE_{idx}__"
                placeholders[placeholder] = f"[{term}](#)"
                return placeholder
            return ignore_replacer
        for term in ignore_terms:
            pattern = re.compile(re.escape(term), re.IGNORECASE)
            new_text = pattern.sub(ignore_replacer_factory(term), new_text)
        return new_text, placeholders
    def restore_bold(text, placeholders):
        for placeholder, link in placeholders.items():
            text = text.replace(placeholder, link)
        return text
    df = pd.read_excel(input_file)
    for col in df.columns:
        df[col] = df[col].astype('object')
    # Validate language codes: only keep those supported by GoogleTranslator
    # Instantiate and call get_supported_languages as an instance method (for deep_translator >=1.11.4)
    try:
        code_dict = GoogleTranslator().get_supported_languages(as_dict=True)
        supported_codes = set(code_dict.values())
    except Exception:
        try:
            supported_codes = set(GoogleTranslator().get_supported_languages())
        except Exception:
            supported_codes = set()
    language_codes = [col for col in df.columns[1:] if str(col).split("-")[0].strip() in supported_codes]
    skipped_codes = [col for col in df.columns[1:] if str(col).split("-")[0].strip() not in supported_codes]
    if skipped_codes:
        print(f"Warning: Skipping unsupported language codes: {', '.join(map(str, skipped_codes))}")
    print("\nTesting translation on first 3 rows...")
    test_df = df.head(3).copy()
    success_count = 0
    fail_count = 0
    for row_idx in range(test_df.shape[0]):
        english_text = str(test_df.iat[row_idx, 0]).strip()
        clean_text = preprocess_text(english_text)
        prepped_text, placeholders = extract_bold(clean_text)
        print(f"\nRow {row_idx+1} (EN): {english_text}")
        for col_idx, lang_code in enumerate(language_codes, start=1):
            short_code = str(lang_code).split("-")[0].strip()
            max_attempts = 3
            attempt = 0
            error = None
            translated = None
            while attempt < max_attempts:
                try:
                    translated = GoogleTranslator(source="en", target=short_code).translate(prepped_text)
                    error = None
                    break
                except Exception as e:
                    error = e
                    attempt += 1
            if error:
                test_df.iat[row_idx, col_idx] = f"[ERROR: {error}]"
                print(f"  {lang_code}: [ERROR: {error}]")
                fail_count += 1
            else:
                translated_str = restore_bold(str(translated), placeholders)
                test_df.iat[row_idx, col_idx] = translated_str
                print(f"  {lang_code}: {translated_str}")
                success_count += 1
    # Output to test-results.xlsx and test-results.csv
    test_df.to_excel("test-results.xlsx", index=False)
    test_df.to_csv("test-results.csv", index=False)
    print("\nTest complete. Results saved to test-results.xlsx and test-results.csv. If translations appear and formatting is preserved, the file is valid.")

    # Generate summary report
    summary_report = [
        f"Test translation complete. Results saved to 'test-results.xlsx'.",
        "Summary:",
        f"  Successful translations: {success_count}",
        f"  Failed translations: {fail_count}",
        f"  Total cells tested: {success_count + fail_count}"
    ]
    with open("test_translation_summary_report.txt", "w", encoding="utf-8") as summary_file:
        summary_file.write("\n".join(summary_report))
    print("Summary report saved to test_translation_summary_report.txt.")

if __name__ == "__main__":
    main()
