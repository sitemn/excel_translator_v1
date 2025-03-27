import os
import time
import logging
import pandas as pd
import numpy as np
import deepl

# Configure Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# File names
pre_file_name = "testnew.xlsx"
new_file_name = "testold.xlsx"

# File paths
base_dir = "/UpdatedExcels/"
pre_file_loc = os.path.join(base_dir, pre_file_name)
new_file_loc = os.path.join(base_dir, new_file_name)
output_file = "output.xlsx"

# DeepL Translator
auth_key = "yourkeyhere"
translator = deepl.Translator(auth_key)

# Columns to remove from all sheets
rem_list = ['1001总表', '829主图', '1001主图', '汇总', '401总表', '409主图', '5332', '25549']

# Expected column structure
expected_columns = ['Product', 'ASIN', 'Model_Requirements', 'Total_Video', 'Scene', 'Pets', 'Requirements', 'Comments']


def read_excel(file_path, columns):
    """Read an Excel file and handle errors."""
    try:
        return pd.read_excel(file_path, sheet_name=None, usecols=columns)
    except FileNotFoundError:
        logger.error(f"File {file_path} not found.")
        exit(1)
    except Exception as e:
        logger.error(f"Error reading {file_path}: {e}")
        exit(1)


def preprocess_sheets(new_df):
    """Remove unwanted columns and delete unwanted sheets."""
    return {sheet: df.drop(columns=[col for col in rem_list if col in df.columns], errors='ignore')
            for sheet, df in new_df.items() if sheet not in rem_list}  # Remove the entire sheet if its name is in rem_list


def translate_column(df, column_name, target_lang='EN-US'):
    """Batch translate a column using DeepL API while handling empty values."""

    # Ensure column exists before processing
    if column_name not in df.columns:
        logger.warning(f"Column {column_name} not found, skipping translation.")
        return df

    # Convert all values to strings and replace NaNs with an empty string
    df[column_name] = df[column_name].astype(str).fillna('')

    # Filter out empty strings before sending to DeepL
    mask = df[column_name] != ""
    texts_to_translate = df.loc[mask, column_name].tolist()

    try:
        if texts_to_translate:  # Ensure we don't send empty requests
            translations = translator.translate_text(texts_to_translate, target_lang=target_lang)
            df.loc[mask, column_name] = [t.text for t in translations]
    except Exception as e:
        logger.error(f"Error translating {column_name}: {e}")

    return df


# def process_sheet(sheet, df, pre_df, new_added_worksheets):
#     """Process a single worksheet."""
#     logger.info(f"Processing {sheet}...")
#
#     if len(df.columns) < len(expected_columns):
#         logger.warning(f"Skipping {sheet} due to missing columns.")
#         return None
#
#     df = df.copy()
#
#     df.columns = expected_columns
#
#     if sheet not in new_added_worksheets:
#         if sheet in pre_df:
#             last_cell = pre_df[sheet].replace('', np.nan).ffill().iloc[-1, -1]
#             filtered_df = df[df['Product'] == last_cell]
#             if not filtered_df.empty:
#                 skip_rows = filtered_df.index[0]
#                 df = df.iloc[skip_rows + 1 :].copy()
#             else:
#                 logger.info(f"No rows to translate in {sheet}.")
#                 return None
#
#     # Fix FutureWarning: Explicit column assignment
#     df['Model_Requirements'] = df['Model_Requirements'].fillna('N/A').copy()
#     df['Scene'] = df['Scene'].fillna('N/A').astype(str).copy()
#     df['Shooting_Requirements'] = (df['Comments'].fillna('').astype(str) + '\r' + df['Requirements'].fillna('').astype(str)).copy()
#
#     # Drop unnecessary columns
#     df = df.drop(columns=['Requirements', 'Comments'], errors='ignore')
#
#     # Translate columns safely
#     df = translate_column(df, 'Product')
#     df = translate_column(df, 'Scene')
#     df = translate_column(df, 'Shooting_Requirements')
#
#     return df


def process_sheet(sheet, df, pre_df, new_added_worksheets):
    """Process a single worksheet, skipping old rows and translating new data."""
    logger.info(f"Processing {sheet}...")

    if len(df.columns) < len(expected_columns):
        logger.warning(f"Skipping {sheet} due to missing columns.")
        return None

    df = df.copy()  # Prevent modifications to original DataFrame
    df.columns = expected_columns

    if sheet not in new_added_worksheets and sheet in pre_df:
        last_cell = pre_df[sheet]['Product'].dropna().iloc[-1]  # Get last processed value

        # Find the first occurrence of last_cell in new_df using .idxmax()
        mask = df['Product'] == last_cell
        if mask.any():  # Ensure last_cell exists in new_df
            first_new_row_index = mask.idxmax()  # Get first occurrence of last_cell
            df = df.iloc[first_new_row_index + 1:].copy()  # Skip processed rows efficiently
        else:
            logger.info(f"No new rows to translate in {sheet}.")
            return None

    # Fill missing values and concatenate fields
    df['Model_Requirements'] = df['Model_Requirements'].fillna('N/A').copy()
    df['Scene'] = df['Scene'].fillna('N/A').astype(str).copy()
    df['Shooting_Requirements'] = (df['Comments'].fillna('').astype(str) + '\r' + df['Requirements'].fillna('').astype(str)).copy()

    # Drop unnecessary columns
    df = df.drop(columns=['Requirements', 'Comments'], errors='ignore')

    # Translate columns safely
    df = translate_column(df, 'Product')
    df = translate_column(df, 'Scene')
    df = translate_column(df, 'Shooting_Requirements')

    return df


def main():
    """Main execution function."""
    logger.info("Starting Excel processing...")

    # Read Excel files
    pre_df = read_excel(pre_file_loc, "D")
    new_df = read_excel(new_file_loc, "D:J,L")
    # Rename the column in all sheets
    for sheet in pre_df:
        pre_df[sheet].columns = ['Product']

    # Preprocess: Remove unnecessary sheets
    new_df = preprocess_sheets(new_df)

    # Compare worksheet names
    pre_worksheets = set(pre_df.keys())
    new_worksheets = set(new_df.keys())
    new_added_worksheets = new_worksheets - pre_worksheets
    deleted_worksheets = pre_worksheets - new_worksheets  # Identify deleted worksheets

    logger.info(f"Previous worksheets: {pre_worksheets}")
    logger.info(f"Latest worksheets: {new_worksheets}")
    logger.info(f"Newly added worksheets: {new_added_worksheets}")
    logger.info(f"Deleted worksheets: {deleted_worksheets}")

    # Process each worksheet
    with pd.ExcelWriter(output_file) as writer:
        for sheet, df in new_df.items():
            processed_df = process_sheet(sheet, df, pre_df, new_added_worksheets)
            if processed_df is not None:
                processed_df.to_excel(writer, sheet_name=sheet, index=False)
                logger.info(f"{sheet} processing complete.")

    logger.info(f"Processing completed. Output saved to {output_file}")


# Run the script
if __name__ == "__main__":
    main()
