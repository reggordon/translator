# Setup Guide

Windows (PC)

```powershell
# 1. Install Python (3.9+)
python --version
# If missing, install from https://www.python.org/downloads/windows/
winget install Python.Python.3.11
# Make sure to tick "Add Python to PATH" during install.

# 2. Create and activate a virtual environment
python -m venv venv
.\venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# requirements.txt should contain:
# pandas
# openpyxl
# deep-translator
# tqdm

# 4. Prepare files
# - Place Fraud_Rules_translations_final.xlsx in the folder
# - Save the script as translate.py

# 5. Run the script
python translate.py

# Output:
# - Translations saved to Fraud_Rules_translations_results.xlsx
# - Failures logged to failed_translations_log.csv


Mac Set Up


# 1. Install Python (3.9+)
python3 --version
# If missing, install with Homebrew:
brew install python

# 2. Create and activate a virtual environment
python3 -m venv venv
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# requirements.txt should contain:
# pandas
# openpyxl
# deep-translator
# tqdm

# 4. Prepare files

# - Ensure Fraud_Rules_translations_final.xlsx is in the same folder as the code
# - If using a different file ensure you change the title in the `input_file` field
#  - If saving to a different location ensure you change the name of the `output_file` field
# - Save the script as translate.py

# 5. Run the script
python3 translate.py

# Output:
# - Translations saved to Fraud_Rules_translations_results.xlsx
# - Failures logged to failed_translations_log.csv
