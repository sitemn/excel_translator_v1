# Excel Translate

**Excel Translate** is a desktop application that processes Excel spreadsheets and automatically translates specified columns using the [DeepL API](https://www.deepl.com/).

It compares a previously processed file with a new one, skips already processed rows, and exports the translated result to a new Excel file.

---

## 🚀 Features

- 🧾 Load "previous" and "new" Excel files
- 🧼 Skip previously processed rows
- 🌍 Translate content with DeepL API
- 🪄 Clean and user-friendly GUI with logging
- ⚙️ Customizable columns/sheets to exclude

---

## 🖥️ Requirements

- Python 3.8+
- [DeepL API Key](https://www.deepl.com/pro-api)
- Dependencies (install below)

---

## 📦 Installation

1. **Clone the repo:**

   ```bash
   git clone https://github.com/yourusername/excel-translate.git
   cd excel-translate
   ```

2. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the app:**

   ```bash
   python main.py
   ```

---

## 📂 Project Structure

```
excel-translate/
├── excel_translate/
│   ├── __init__.py
│   ├── gui.py               # GUI logic
│   ├── translator.py        # Translation helper
│   └── excel_utils.py       # Excel file helpers
├── main.py                  # App entry point
├── requirements.txt         # Dependencies
└── README.md                # This file
```

---

## 🔐 DeepL API

You’ll need a DeepL API key to use the translation feature.

- Sign up: https://www.deepl.com/pro
- Paste your API key in the app’s GUI when prompted

---

## ✅ Example Usage

1. Select the “Previous Excel File” (used to skip already-processed rows)
2. Select the “New Excel File” to process
3. Set the output file name
4. Enter your DeepL API key
5. Click **Process Excel Files**

---

## 📃 License

MIT License. Free for personal and commercial use.

---

## 🤝 Contributions

PRs welcome! Feel free to open issues for bugs or feature suggestions.