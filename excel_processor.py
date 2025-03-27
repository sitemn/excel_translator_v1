import os
import logging
import pandas as pd
import deepl
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
import threading
import queue


class LogHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor with Translation")
        self.root.geometry("900x700")
        self.root.minsize(900, 700)

        # Create queue for logging
        self.log_queue = queue.Queue()

        # Configure logging
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)

        # Create console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)

        # Create queue handler for GUI
        queue_handler = LogHandler(self.log_queue)
        queue_handler.setLevel(logging.INFO)
        queue_handler.setFormatter(console_formatter)
        self.logger.addHandler(queue_handler)

        # Variables
        self.pre_file_path = tk.StringVar()
        self.new_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.deepl_key = tk.StringVar()
        self.remove_columns = tk.StringVar(value='1001总表,829主图,1001主图,汇总,401总表,409主图,5332,25549')

        self.expected_columns = ['Product', 'ASIN', 'Model_Requirements', 'Total_Video', 'Scene', 'Pets',
                                 'Requirements', 'Comments']

        # Create GUI elements
        self.create_widgets()

        # Set up periodic log check
        self.after_id = None
        self.check_logs()

    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding=10)
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        # Previous file
        ttk.Label(file_frame, text="Previous Excel File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.pre_file_path, width=50).grid(row=0, column=1, sticky=tk.W + tk.E,
                                                                              padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_pre_file).grid(row=0, column=2, padx=5, pady=5)

        # New file
        ttk.Label(file_frame, text="New Excel File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.new_file_path, width=50).grid(row=1, column=1, sticky=tk.W + tk.E,
                                                                              padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_new_file).grid(row=1, column=2, padx=5, pady=5)

        # Output file
        ttk.Label(file_frame, text="Output File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_file_path, width=50).grid(row=2, column=1, sticky=tk.W + tk.E,
                                                                                 padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output_file).grid(row=2, column=2, padx=5, pady=5)

        # Configuration frame
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding=10)
        config_frame.pack(fill=tk.X, padx=5, pady=5)

        # DeepL API key
        ttk.Label(config_frame, text="DeepL API Key:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(config_frame, textvariable=self.deepl_key, width=50, show="*").grid(row=0, column=1,
                                                                                      sticky=tk.W + tk.E, padx=5,
                                                                                      pady=5)

        # Columns to remove
        ttk.Label(config_frame, text="Columns/Sheets to Remove:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(config_frame, textvariable=self.remove_columns, width=50).grid(row=1, column=1, sticky=tk.W + tk.E,
                                                                                 padx=5, pady=5)

        # Process button
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame, text="Process Excel Files", command=self.start_processing).pack(pady=10)

        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, width=80, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)

        # Configure grid expansions
        file_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(1, weight=1)

    def browse_pre_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.pre_file_path.set(file_path)

    def browse_new_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.new_file_path.set(file_path)
            # Set default output path to the same directory
            if not self.output_file_path.get():
                dir_name = os.path.dirname(file_path)
                self.output_file_path.set(os.path.join(dir_name, "output.xlsx"))

    def browse_output_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_file_path.set(file_path)

    def check_logs(self):
        # Check for new log messages
        while not self.log_queue.empty():
            try:
                message = self.log_queue.get_nowait()
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
            except queue.Empty:
                break

        # Schedule the next check
        self.after_id = self.root.after(100, self.check_logs)

    def read_excel(self, file_path, columns=None):
        """Read an Excel file and handle errors."""
        try:
            return pd.read_excel(file_path, sheet_name=None, usecols=columns)
        except FileNotFoundError:
            self.logger.error(f"File {file_path} not found.")
            return None
        except Exception as e:
            self.logger.error(f"Error reading {file_path}: {e}")
            return None

    def preprocess_sheets(self, new_df, rem_list):
        """Remove unwanted columns and delete unwanted sheets."""
        return {sheet: df.drop(columns=[col for col in rem_list if col in df.columns], errors='ignore')
                for sheet, df in new_df.items() if sheet not in rem_list}

    def translate_column(self, df, column_name, translator, target_lang='EN-US'):
        """Batch translate a column using DeepL API while handling empty values."""
        # Ensure column exists before processing
        if column_name not in df.columns:
            self.logger.warning(f"Column {column_name} not found, skipping translation.")
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
            self.logger.error(f"Error translating {column_name}: {e}")

        return df

    def process_sheet(self, sheet, df, pre_df, new_added_worksheets, translator):
        """Process a single worksheet, skipping old rows and translating new data."""
        self.logger.info(f"Processing {sheet}...")

        if len(df.columns) < len(self.expected_columns):
            self.logger.warning(f"Skipping {sheet} due to missing columns.")
            return None

        df = df.copy()  # Prevent modifications to original DataFrame
        df.columns = self.expected_columns

        if sheet not in new_added_worksheets and sheet in pre_df:
            if not pre_df[sheet].empty:
                last_cell = pre_df[sheet]['Product'].dropna().iloc[-1]  # Get last processed value

                # Find the first occurrence of last_cell in new_df using .idxmax()
                mask = df['Product'] == last_cell
                if mask.any():  # Ensure last_cell exists in new_df
                    first_new_row_index = mask.idxmax()  # Get first occurrence of last_cell
                    df = df.iloc[first_new_row_index + 1:].copy()  # Skip processed rows efficiently
                else:
                    self.logger.info(f"No new rows to translate in {sheet}.")
                    return None

        # Fill missing values and concatenate fields
        df['Model_Requirements'] = df['Model_Requirements'].fillna('N/A').copy()
        df['Scene'] = df['Scene'].fillna('N/A').astype(str).copy()
        df['Shooting_Requirements'] = (
                    df['Comments'].fillna('').astype(str) + '\r' + df['Requirements'].fillna('').astype(str)).copy()

        # Drop unnecessary columns
        df = df.drop(columns=['Requirements', 'Comments'], errors='ignore')

        # Translate columns safely
        df = self.translate_column(df, 'Product', translator)
        df = self.translate_column(df, 'Scene', translator)
        df = self.translate_column(df, 'Shooting_Requirements', translator)

        return df

    def process_files(self):
        """Process the Excel files based on GUI inputs."""
        pre_file_loc = self.pre_file_path.get()
        new_file_loc = self.new_file_path.get()
        output_file = self.output_file_path.get()
        auth_key = self.deepl_key.get()

        # Parse remove list
        rem_list = [item.strip() for item in self.remove_columns.get().split(',')]

        self.logger.info("Starting Excel processing...")

        # Validate inputs
        if not pre_file_loc or not new_file_loc:
            self.logger.error("Please select both previous and new Excel files.")
            return

        if not output_file:
            self.logger.error("Please specify an output file location.")
            return

        if not auth_key:
            self.logger.error("Please enter a DeepL API key.")
            return

        try:
            # Initialize DeepL translator
            translator = deepl.Translator(auth_key)

            # Read Excel files
            self.logger.info(f"Reading previous file: {pre_file_loc}")
            pre_df = self.read_excel(pre_file_loc, "D")
            if pre_df is None:
                return

            self.logger.info(f"Reading new file: {new_file_loc}")
            new_df = self.read_excel(new_file_loc, "D:J,L")
            if new_df is None:
                return

            # Rename the column in all sheets
            for sheet in pre_df:
                pre_df[sheet].columns = ['Product']

            # Preprocess: Remove unnecessary sheets
            self.logger.info("Preprocessing sheets...")
            new_df = self.preprocess_sheets(new_df, rem_list)

            # Compare worksheet names
            pre_worksheets = set(pre_df.keys())
            new_worksheets = set(new_df.keys())
            new_added_worksheets = new_worksheets - pre_worksheets
            deleted_worksheets = pre_worksheets - new_worksheets  # Identify deleted worksheets

            self.logger.info(f"Previous worksheets: {pre_worksheets}")
            self.logger.info(f"Latest worksheets: {new_worksheets}")
            self.logger.info(f"Newly added worksheets: {new_added_worksheets}")
            self.logger.info(f"Deleted worksheets: {deleted_worksheets}")

            # Process each worksheet
            self.logger.info(f"Writing output to: {output_file}")
            with pd.ExcelWriter(output_file) as writer:
                for sheet, df in new_df.items():
                    processed_df = self.process_sheet(sheet, df, pre_df, new_added_worksheets, translator)
                    if processed_df is not None and not processed_df.empty:
                        processed_df.to_excel(writer, sheet_name=sheet, index=False)
                        self.logger.info(f"{sheet} processing complete.")

            self.logger.info(f"Processing completed. Output saved to {output_file}")

        except Exception as e:
            self.logger.error(f"Error processing files: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())

    def start_processing(self):
        """Start processing in a separate thread to keep GUI responsive."""
        # Clear log area
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        # Disable the process button
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.config(state=tk.DISABLED)

        # Start processing thread
        processing_thread = threading.Thread(target=self.run_processing)
        processing_thread.daemon = True
        processing_thread.start()

    def run_processing(self):
        """Run processing and re-enable GUI when finished."""
        try:
            self.process_files()
        finally:
            # Re-enable the process button
            self.root.after(0, self.enable_buttons)

    def enable_buttons(self):
        """Re-enable all buttons in the GUI."""
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()