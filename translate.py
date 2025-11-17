

#!/usr/bin/env python3
# USAGE: Make this file executable with 'chmod +x translate.py' and run with './translate.py'

import os
import sys
import time
import csv
import threading
import pandas as pd
from deep_translator import GoogleTranslator, LibreTranslator
from tqdm import tqdm
from langdetect import detect, LangDetectException
import difflib



# List all Excel files in the current directory
def list_excel_files():
    return [f for f in os.listdir('.') if f.endswith('.xlsx')]


# Prompt user to select a file from a list
def choose_file(files, prompt):
    print(f"\n{prompt}")
    for idx, f in enumerate(files):
        print(f"  [{idx+1}] {f}")
    while True:
        choice = input("Enter number: ")
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice)-1]
        print("Invalid choice. Try again.")


# Prompt user to select translation backend
def choose_backend():
    print("\nChoose translation backend:")
    print("  [1] GoogleTranslator")
    print("  [2] LibreTranslator")
    print("  [3] Try Google, fallback to Libre")
    while True:
        choice = input("Enter number: ")
        if choice in ['1', '2', '3']:
            return choice
        print("Invalid choice. Try again.")


def main():
    print("\n=== Translation Script ===")
    excel_files = list_excel_files()
    if not excel_files:
        print("No .xlsx files found in current directory.")
        sys.exit(1)

    input_file = choose_file(excel_files, "Select input Excel file:")
    output_file = input("Enter output Excel filename (default: translated_results.xlsx): ").strip()
    if not output_file:
        output_file = "translated_results.xlsx"

    df = pd.read_excel(input_file)
    language_codes = df.columns[1:]
    backend_choice = choose_backend()

    failed_translations = []
    rows_to_translate = [
        row_idx for row_idx in range(df.shape[0])
        if pd.notna(df.iat[row_idx, 0]) and str(df.iat[row_idx, 0]).strip() != ""
    ]

    success_count = 0
    skip_count = 0
    fail_count = 0
    suspect_translations = []

    # Helper: run translation with timeout
    def translate_with_timeout(func, args=(), timeout=15):
        result = {}
        def wrapper():
            try:
                result['value'] = func(*args)
            except Exception as e:
                result['error'] = e
        thread = threading.Thread(target=wrapper)
        thread.start()
        thread.join(timeout)
        if thread.is_alive():
            return None, TimeoutError('Translation timed out')
        return result.get('value'), result.get('error')

    with tqdm(total=len(rows_to_translate), desc="Translating rows", unit="row") as pbar:
        for row_idx in rows_to_translate:
            english_text = str(df.iat[row_idx, 0]).strip()
            for col_idx, lang_code in enumerate(language_codes, start=1):
                short_code = str(lang_code).split("-")[0].strip()
                # Skip if already translated
                if pd.notna(df.iat[row_idx, col_idx]) and str(df.iat[row_idx, col_idx]).strip() != "":
                    skip_count += 1
                    continue
                try:
                    translated = None
                    backend = ""
                    if backend_choice == '1':
                        translated, error = translate_with_timeout(
                            GoogleTranslator(source="en", target=short_code).translate,
                            (english_text,), 15)
                        backend = "Google"
                    elif backend_choice == '2':
                        translated, error = translate_with_timeout(
                            LibreTranslator(source="en", target=short_code).translate,
                            (english_text,), 15)
                        backend = "Libre"
                    else:
                        translated, error = translate_with_timeout(
                            GoogleTranslator(source="en", target=short_code).translate,
                            (english_text,), 15)
                        backend = "Google"
                        if error:
                            translated, error = translate_with_timeout(
                                LibreTranslator(source="en", target=short_code).translate,
                                (english_text,), 15)
                            backend = "Libre"
                    if error:
                        raise error
                    translated_str = str(translated)
                    df.iat[row_idx, col_idx] = translated_str
                    success_count += 1

                    # --- Translation check: back-translate and language detect ---
                    try:
                        back_translated, bt_error = translate_with_timeout(
                            GoogleTranslator(source=short_code, target="en").translate,
                            (translated_str,), 15)
                        if bt_error:
                            back_translated = ""
                        similarity = difflib.SequenceMatcher(None, english_text, back_translated).ratio() if back_translated else 0.0
                        try:
                            detected_lang = detect(translated_str)
                        except LangDetectException:
                            detected_lang = "unknown"
                        if similarity < 0.7 or (
                            detected_lang != short_code and
                            detected_lang != lang_code and
                            detected_lang != "unknown"
                        ):
                            suspect_translations.append({
                                "row": row_idx,
                                "english_text": english_text,
                                "language_code": lang_code,
                                "short_code": short_code,
                                "translated_text": translated_str,
                                "back_translated": back_translated,
                                "similarity": similarity,
                                "detected_lang": detected_lang
                            })
                    except Exception:
                        pass

                except Exception as e:
                    failed_translations.append({
                        "row": row_idx,
                        "english_text": english_text,
                        "language_code": lang_code,
                        "short_code": short_code,
                        "error": str(e)
                    })
                    df.iat[row_idx, col_idx] = ""
                    backend = "FAILED"
                    fail_count += 1
                time.sleep(0.3)
            pbar.update(1)

    # Save main results and suspects to separate sheets in the same Excel file
    import openpyxl
    from pandas import ExcelWriter

    # Ensure output file is created even if there are no translations
    with ExcelWriter(output_file, engine="openpyxl") as writer:
        if df.shape[0] == 0:
            # Create empty DataFrame with expected columns
            empty_df = pd.DataFrame(columns=df.columns)
            empty_df.to_excel(writer, index=False, sheet_name="Translations")
        else:
            df.to_excel(writer, index=False, sheet_name="Translations")
        if suspect_translations:
            suspects_df = pd.DataFrame(suspect_translations)
            suspects_df.to_excel(writer, index=False, sheet_name="SuspectTranslations")

    # Still save failures to CSV for easy review
    if failed_translations:
        with open("failed_translations_log.csv", mode="w", newline="", encoding="utf-8") as log_file:
            writer = csv.DictWriter(
                log_file,
                fieldnames=["row", "english_text", "language_code", "short_code", "error"]
            )
            writer.writeheader()
            writer.writerows(failed_translations)

    print(f"\n✅ Translations complete. Results saved to '{output_file}'.")
    print(f"Summary:")
    print(f"  Successful translations: {success_count}")
    print(f"  Skipped cells (already translated): {skip_count}")
    print(f"  Failed translations: {fail_count}")
    print(f"  Suspect translations (review): {len(suspect_translations)}")
    if failed_translations:
        print(f"⚠️ {len(failed_translations)} failures logged to 'failed_translations_log.csv'")
    if suspect_translations:
        print(f"⚠️ {len(suspect_translations)} suspect translations logged to 'suspect_translations_review.csv'")

        # Save summary report to file
        summary_report = [
            f"Translations complete. Results saved to '{output_file}'.",
            "Summary:",
            f"  Successful translations: {success_count}",
            f"  Skipped cells (already translated): {skip_count}",
            f"  Failed translations: {fail_count}",
            f"  Suspect translations (review): {len(suspect_translations)}"
        ]
        if failed_translations:
            summary_report.append(f"{len(failed_translations)} failures logged to 'failed_translations_log.csv'")
        if suspect_translations:
            summary_report.append(f"{len(suspect_translations)} suspect translations logged to 'suspect_translations_review.csv'")
        with open("translation_summary_report.txt", "w", encoding="utf-8") as summary_file:
            summary_file.write("\n".join(summary_report))


if __name__ == "__main__":
    main()
