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

## Getting Started


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

---

If you prefer manual setup, follow these steps:

1. Install Python (3.9+)
2. Create and activate a virtual environment
3. Install dependencies with `pip install -r requirements.txt`
4. Run the script with `python translate.py`

## Troubleshooting
- Ensure all required packages are installed (`pip install -r requirements.txt`)
- Activate your virtual environment before running the script
- If you see missing package errors, re-run the setup steps

## License
MIT
