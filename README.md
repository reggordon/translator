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
- **[BOLD] Row Support:** Both `translate.py` and `test_translate.py` support context-aware handling of `[BOLD]` rows. You can specify one or more bold words/phrases in a `[BOLD]` row immediately following a main text row. The scripts will extract, translate, and report all bold words in context, returning all translations in a single output row for review.
- **Multiple Bold Words:** If a `[BOLD]` row contains multiple bold words/phrases (comma-separated), all will be processed and reported together for that main text row.
- **Test Script:** The test script (`test_translate.py`) uses the same context-aware `[BOLD]` logic as the main script, including multi-bold support and consolidated output. It generates a summary report (`test_translation_summary_report.txt`) showing the number of successful and failed translations for your sample, and outputs all bold word translations in a dedicated sheet.

## Getting Started

**Best Practice:** To ensure all dependencies are installed and the correct Python environment is used, always run the provided setup scripts to launch the translation tools. This is the recommended and supported workflow for this project.

**For Mac/Linux users:**
- Use `zsh setup.sh` to run the main translation script (`translate.py`).
- Use `zsh test.sh` to run the test script (`test_translate.py`).

**For Windows users:**
- Use `setup.bat` to run the main translation script (`translate.py`).
- Use `test.bat` to run the test script (`test_translate.py`).

These scripts will create a virtual environment, install dependencies, and launch the appropriate tool. Keeping to this workflow avoids environment issues and ensures consistent results. You should keep the terminal open or reactivate the environment (`source venv/bin/activate` on Mac/Linux, `venv\Scripts\activate` on Windows) if you want to run other scripts manually.

## Usage Notes

- When prompted, you can enter a comma-separated list of terms to ignore (e.g., `CLICK TO PAY, VISA`). These will not be translated and will be preserved as links in the output. Leave blank to translate all text.
- The test script (`test_translate.py`) uses the same logic as the main script, including input preprocessing, retry logic, language code validation, and full support for `[BOLD]` rows and multiple bold words per row. Use it to check your file and ignore terms before running the full translation. After running, check `test_translation_summary_report.txt` for a summary of results, and review the `ContextBoldWords` sheet in the output Excel file for all bold word translations.


### Running the Scripts


#### On Mac/Linux:

**To run the main translation script:**
1. Place your Excel file (`.xlsx`) in the project folder.
2. Run:
   ```zsh
   zsh setup.sh
   ```
   This will set up the environment and launch `translate.py`.

**To run the test script:**
1. Place your Excel file (`.xlsx`) in the project folder.
2. Run:
   ```zsh
   zsh test.sh
   ```
   This will set up the environment and launch `test_translate.py`.

#### On Windows:

**To run the main translation script:**
1. Place your Excel file (`.xlsx`) in the project folder.
2. Double-click `setup.bat` or run in Command Prompt:
   ```bat
   setup.bat
   ```
   This will set up the environment and launch `translate.py`.

**To run the test script:**
1. Place your Excel file (`.xlsx`) in the project folder.
2. Double-click `test.bat` or run in Command Prompt:
   ```bat
   test.bat
   ```
   This will set up the environment and launch `test_translate.py`.

Follow the prompts in each script to select your file and enter ignore terms. The test script will show translations for the first 3 rows for all language columns. If the output looks correct, and the summary report shows no unexpected failures, proceed to run the full translation script as above.


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

## Recommended Structure for Bold Words

If you need to track and preserve bold words in your translations, use the following approach:

- Keep your main text in one row.
- Directly below each main text row, add a row that lists only the bold words or phrases from the main text, prefixed with [BOLD].

Example:

| Text                                 |
|--------------------------------------|
| This is a very important message.    |
| [BOLD] important                     |
| Please click the Confirm button.     |
| [BOLD] Confirm                       |

This structure allows you to:
- Clearly identify which words should be bold in each message.
- Programmatically process and format bold words during translation or post-processing.
- Avoid adding extra columns or complex formatting in the source file.

**Note:** You will need to adjust your translation script to recognize and handle these [BOLD] rows as indicators for formatting.

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

