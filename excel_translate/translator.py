import deepl
import logging

def translate_column(df, column_name, translator, target_lang='EN-US'):
    """Batch translate a column using DeepL API while handling empty values."""
    if column_name not in df.columns:
        logging.warning(f"Column {column_name} not found, skipping translation.")
        return df

    df[column_name] = df[column_name].astype(str).fillna('')
    mask = df[column_name] != ""
    texts_to_translate = df.loc[mask, column_name].tolist()

    try:
        if texts_to_translate:
            translations = translator.translate_text(texts_to_translate, target_lang=target_lang)
            df.loc[mask, column_name] = [t.text for t in translations]
    except Exception as e:
        logging.error(f"Error translating {column_name}: {e}")

    return df
