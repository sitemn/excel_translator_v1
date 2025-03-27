import pandas as pd
import logging

def read_excel(file_path, columns=None):
    try:
        return pd.read_excel(file_path, sheet_name=None, usecols=columns)
    except FileNotFoundError:
        logging.error(f"File {file_path} not found.")
        return None
    except Exception as e:
        logging.error(f"Error reading {file_path}: {e}")
        return None

def preprocess_sheets(new_df, rem_list):
    return {sheet: df.drop(columns=[col for col in rem_list if col in df.columns], errors='ignore')
            for sheet, df in new_df.items() if sheet not in rem_list}
