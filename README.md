# Translator Project

## Overview

This project provides scripts to automate translation and cleaning of Excel files using GoogleTranslator and LibreTranslator. It supports multiple languages and logs failed or suspect translations for review.

## Scripts

- **setup.sh**: Automates environment setup and runs the translation script. Creates a virtual environment, installs dependencies, and launches `translate.py`.
- **translate.py**: Main translation script. Prompts for an Excel file, translates text to multiple languages, performs self-checks (back-translation and language detection), and saves results and logs.
- **translation_clean.py**: Cleans and retries translations in an existing Excel file. Attempts to fix missing or placeholder translations, but does not perform self-checks.

## Features
- Translate Excel files (.xlsx) with multiple language columns
- Supports GoogleTranslator and LibreTranslator backends
- Logs failed and suspect translations
- Generates a summary report after each run
- **Ignore Terms:** You can specify a comma-separated list of terms (e.g., product names, trademarks) to be ignored during translation. These terms will be preserved as links and not translated. If you leave the input blank, all text will be translated as normal.
- **Formatting Preservation:** Bold text (markdown `**bold**`) is preserved and output as a link. Ignored terms are also output as links.
- **Test Script:** A test script (`test_translate.py`) is provided to quickly validate your Excel file and translation settings on the first 3 rows before running the full translation. The test script now also generates a summary report (`test_translation_summary_report.txt`) showing the number of successful and failed translations for your sample.

## Getting Started

## Usage Notes

- When prompted, you can enter a comma-separated list of terms to ignore (e.g., `CLICK TO PAY, VISA`). These will not be translated and will be preserved as links in the output. Leave blank to translate all text.
- The test script (`test_translate.py`) uses the same logic as the main script, including input preprocessing, retry logic, and language code validation for improved accuracy. Use it to check your file and ignore terms before running the full translation. After running, check `test_translation_summary_report.txt` for a summary of results.

### Running the Test Script

To quickly validate your Excel file and ignore terms:

1. Place your Excel file (`.xlsx`) in the project folder.
2. Run:
   ```zsh
   python test_translate.py
   ```
3. Follow the prompts to select your file and enter ignore terms. The script will show translations for the first 3 rows for all language columns.


If the output looks correct, and the summary report shows no unexpected failures, proceed to run the full translation script as described above.


## Quick Setup (Recommended)

1. Place your Excel file (`.xlsx`) in the project folder.
2. Run the setup script:
   ```zsh
   zsh setup.sh
   ```
   This will:
   - Create and activate a virtual environment
   - Install all required dependencies
   - Run the translation script
   - Guide you through selecting your input/output files and translation backend

3. Output files:
   - Translations saved to your chosen output Excel file
   - Failures logged to `failed_translations_log.csv`
   - Suspect translations logged to `suspect_translations_review.csv`
   - Summary report saved to `translation_summary_report.txt`
   - Test summary report saved to `test_translation_summary_report.txt` (when running the test script)

---

If you prefer manual setup, follow these steps:

1. Install Python (3.9+)
2. Create and activate a virtual environment
3. Install dependencies with `pip install -r requirements.txt`
4. Run the script with `python translate.py`

## File Format Tips
- **Source column:** The first column should contain the English (or source language) terms.
- **Language columns:** Each subsequent column should be named with a valid language code (e.g., `fr`, `de`, `zh-TW`).
- **No empty headers:** Avoid blank column headers; only columns with valid language codes will be translated.
- **Technical terms only:** This tool is optimized for technical terms, not full sentences or grammar-heavy content.
- **Consistent terminology:** For best results, use a glossary or termbase for key phrases.

## Example Excel Layout
| Term         | fr    | de    | zh-TW |
|--------------|-------|-------|-------|
| Checkout     |       |       |       |
| Payment      |       |       |       |
| Cancel Order |       |       |       |

## Tips
- Review suspect and failed translations in the summary report.
- Excluded columns (unsupported or empty) are listed in the report.
- For best accuracy, keep terms short and unambiguous.
- If you need to add a glossary, contact the developer.

## Supported Language Codes
- Use codes supported by Google or Libre translators (e.g., `fr`, `de`, `es`, `zh-TW`).
- Region-specific codes (like `zh-TW`) are supported if recognized by the backend.

## Troubleshooting
- If translations fail, check your internet connection and language codes.
- Review the summary report for details on skipped or failed columns.

## License
MIT

## Privacy & Security
See [PRIVACY.md](PRIVACY.md) for important privacy and security guidelines when using this tool in a company setting.
