import pandas as pd
from deep_translator import GoogleTranslator
from tqdm import tqdm
import time

# Load flagged file
df = pd.read_excel("Fraud_Rules_translations_final.xlsx'")
language_codes = df.iloc[1, 1:].tolist()

# Create editable copy
df_cleaned = df.copy()
num_rows = df.shape[0] - 2
num_cols = len(language_codes)

# Progress bar
with tqdm(total=num_rows * num_cols, desc="Cleaning translations", unit="cell") as pbar:
    for row_idx in range(2, df.shape[0]):
        english_text = str(df.iat[row_idx, 0]).strip()
        if not english_text:
            pbar.update(num_cols)
            continue

        for col_idx in range(1, len(df.columns)):
            lang_code = language_codes[col_idx - 1]
            short_code = lang_code.split("_")[0].strip()
            cell_value = str(df_cleaned.iat[row_idx, col_idx]).strip()

            try:
                # Retry missing/untranslated
                if not cell_value or cell_value.lower() == english_text.lower():
                    translated = GoogleTranslator(source='en', target=short_code).translate(english_text)
                    df_cleaned.iat[row_idx, col_idx] = translated
                    time.sleep(0.3)

                # Clean bracketed placeholders
                elif "[" in cell_value and "]" in cell_value:
                    parts = []
                    for token in english_text.split():
                        if token.startswith("[") and token.endswith("]"):
                            parts.append(token)
                        else:
                            translated = GoogleTranslator(source='en', target=short_code).translate(token)
                            parts.append(translated)
                            time.sleep(0.3)
                    df_cleaned.iat[row_idx, col_idx] = " ".join(parts)

            except:
                pass  # Leave cell unchanged on error

            pbar.update(1)

        # 🔐 Auto-save after each row
        df_cleaned.to_excel("Fraud_rules.xlsx", index=False)

print("✅ Cleaned file saved as: SRCi_translations_autocleaned.xlsx")
