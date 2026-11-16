"""
DATA CLEANING GUI APPLICATION
Libraries: pandas, tkinter (built into Python), openpyxl (for Excel files)

Install required packages:
    pip install pandas openpyxl

Run:
    python data_cleaning_gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path


class DataCleaningApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Cleaning and Preprocessing Application")
        self.root.geometry("1180x760")
        self.root.minsize(950, 620)

        self.original_df = None
        self.df = None
        self.file_path = None

        self.setup_style()
        self.create_widgets()

    def setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Title.TLabel", font=("Segoe UI", 19, "bold"),
                        foreground="#17365D")
        style.configure("Sub.TLabel", font=("Segoe UI", 10),
                        foreground="#526477")
        style.configure("Action.TButton", font=("Segoe UI", 10, "bold"),
                        padding=7)
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=25)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def create_widgets(self):
        header = ttk.Frame(self.root, padding=(18, 14))
        header.pack(fill="x")

        ttk.Label(header, text="DATA CLEANING & PREPROCESSING",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(header,
                  text="Load, inspect, clean and export your dataset",
                  style="Sub.TLabel").pack(anchor="w", pady=(3, 0))

        toolbar = ttk.Frame(self.root, padding=(15, 5))
        toolbar.pack(fill="x")

        buttons = [
            ("Load Dataset", self.load_dataset),
            ("Preview Original", self.show_original),
            ("Preview Cleaned", self.show_cleaned),
            ("Save Cleaned Data", self.save_dataset),
        ]
        for label, command in buttons:
            ttk.Button(toolbar, text=label, command=command,
                       style="Action.TButton").pack(side="left", padx=4)

        self.file_label = ttk.Label(self.root, text="No dataset loaded",
                                    padding=(19, 3))
        self.file_label.pack(anchor="w")

        self.stats_frame = ttk.LabelFrame(self.root, text="Dataset Information",
                                          padding=10)
        self.stats_frame.pack(fill="x", padx=18, pady=8)

        self.stats_text = tk.StringVar(
            value="Load a CSV or Excel file to see its information."
        )
        ttk.Label(self.stats_frame, textvariable=self.stats_text,
                  justify="left", wraplength=1080).pack(anchor="w")

        main = ttk.Frame(self.root, padding=(18, 5))
        main.pack(fill="both", expand=True)

        left = ttk.LabelFrame(main, text="Cleaning Operations", padding=10)
        left.pack(side="left", fill="y", padx=(0, 12))

        operations = [
            ("Inspect Missing Values", self.inspect_missing),
            ("Remove Duplicate Rows", self.remove_duplicates),
            ("Handle Invalid Values", self.handle_invalid_values),
            ("Fill Missing Values", self.fill_missing),
            ("Trim & Clean Text", self.trim_text),
            ("Convert Numeric Columns", self.convert_numeric),
            ("Handle Numeric Outliers (IQR)", self.handle_outliers),
            ("Clean Column Names", self.clean_column_names),
            ("Reset to Original Data", self.reset_data),
        ]

        for label, command in operations:
            ttk.Button(left, text=label, command=command,
                       width=29).pack(fill="x", pady=4)

        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True)

        preview_frame = ttk.LabelFrame(right, text="Dataset Preview",
                                       padding=8)
        preview_frame.pack(fill="both", expand=True)

        self.preview_title = ttk.Label(preview_frame, text="No data to display")
        self.preview_title.pack(anchor="w", pady=(0, 5))

        table_container = ttk.Frame(preview_frame)
        table_container.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(table_container, show="headings")
        y_scroll = ttk.Scrollbar(table_container, orient="vertical",
                                 command=self.tree.yview)
        x_scroll = ttk.Scrollbar(preview_frame, orient="horizontal",
                                 command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set,
                            xscrollcommand=x_scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(fill="x")

        log_frame = ttk.LabelFrame(right, text="Cleaning Activity Log",
                                   padding=8)
        log_frame.pack(fill="x", pady=(10, 0))

        self.log = tk.Text(log_frame, height=7, wrap="word",
                           font=("Consolas", 9), state="disabled")
        self.log.pack(fill="both", expand=True)

        self.write_log("Application started. Load your dataset to begin.")

    def write_log(self, message):
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def require_data(self):
        if self.df is None:
            messagebox.showwarning("No Dataset",
                                   "Please load a CSV or Excel dataset first.")
            return False
        return True

    def load_dataset(self):
        path = filedialog.askopenfilename(
            title="Select a dataset",
            filetypes=[
                ("Supported files", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        if not path:
            return

        try:
            extension = Path(path).suffix.lower()
            if extension == ".csv":
                try:
                    loaded = pd.read_csv(path)
                except UnicodeDecodeError:
                    loaded = pd.read_csv(path, encoding="latin-1")
            elif extension in (".xlsx", ".xls"):
                loaded = pd.read_excel(path)
            else:
                messagebox.showerror("Unsupported File",
                                     "Please select a CSV or Excel file.")
                return

            self.file_path = path
            self.original_df = loaded.copy(deep=True)
            self.df = loaded.copy(deep=True)

            self.file_label.config(text=f"Loaded file: {Path(path).name}")
            self.update_stats()
            self.display_dataframe(self.original_df, "Original Dataset")
            self.write_log(
                f"Loaded {Path(path).name}: "
                f"{len(self.df)} rows and {len(self.df.columns)} columns."
            )
            messagebox.showinfo("Dataset Loaded",
                                "Your dataset has been loaded successfully.")
        except Exception as error:
            messagebox.showerror("Loading Error", str(error))

    def display_dataframe(self, dataframe, title):
        self.preview_title.config(
            text=f"{title} — showing up to 200 rows"
        )

        self.tree.delete(*self.tree.get_children())
        columns = [str(column) for column in dataframe.columns]
        self.tree["columns"] = columns

        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=140, minwidth=90, stretch=True)

        preview = dataframe.head(200).copy()
        preview = preview.astype(object).where(pd.notna(preview), "Missing")

        for _, row in preview.iterrows():
            values = [str(value)[:250] for value in row.tolist()]
            self.tree.insert("", "end", values=values)

    def update_stats(self):
        if self.df is None:
            return

        missing = int(self.df.isna().sum().sum())
        duplicates = int(self.df.duplicated().sum())
        numeric_count = len(self.df.select_dtypes(include="number").columns)
        text_count = len(self.df.select_dtypes(include="object").columns)

        self.stats_text.set(
            f"Rows: {len(self.df)}     "
            f"Columns: {len(self.df.columns)}     "
            f"Missing cells: {missing}     "
            f"Duplicate rows: {duplicates}     "
            f"Numeric columns: {numeric_count}     "
            f"Text columns: {text_count}"
        )

    def show_original(self):
        if self.original_df is not None:
            self.display_dataframe(self.original_df, "Original Dataset")
        else:
            self.require_data()

    def show_cleaned(self):
        if self.require_data():
            self.display_dataframe(self.df, "Current Cleaned Dataset")

    def inspect_missing(self):
        if not self.require_data():
            return

        report = self.df.isna().sum()
        report = report[report > 0]

        if report.empty:
            message = "No missing values were found."
            messagebox.showinfo("Missing Values", message)
        else:
            details = "\n".join(
                f"{column}: {count}" for column, count in report.items()
            )
            message = "Missing values by column:\n\n" + details
            messagebox.showinfo("Missing Values", message)

        self.write_log("Missing-value inspection completed.")
        self.update_stats()

    def remove_duplicates(self):
        if not self.require_data():
            return

        before = len(self.df)
        self.df = self.df.drop_duplicates().reset_index(drop=True)
        removed = before - len(self.df)

        self.write_log(f"Duplicate removal: {removed} duplicate row(s) removed.")
        self.update_stats()
        self.display_dataframe(self.df, "Current Cleaned Dataset")
        messagebox.showinfo("Duplicates Removed",
                            f"{removed} duplicate row(s) removed.")

    def handle_invalid_values(self):
        if not self.require_data():
            return

        before_rows = len(self.df)

        # 1. 'unavailable' jaise invalid text ko numeric NaN mein convert karna
        if "price" in self.df.columns:
            self.df["price"] = pd.to_numeric(self.df["price"], errors="coerce")

        # 2. Negative ya zero area values (area_sqft <= 0) ko remove karna
        if "area_sqft" in self.df.columns:
            self.df = self.df[self.df["area_sqft"] > 0].reset_index(drop=True)

        removed_rows = before_rows - len(self.df)
        self.write_log(
            f"Invalid values handled: 'unavailable' converted to NaN and {removed_rows} invalid row(s) removed."
        )
        self.update_stats()
        self.display_dataframe(self.df, "Current Cleaned Dataset")
        messagebox.showinfo(
            "Invalid Values Handled",
            f"Invalid text converted to Missing and {removed_rows} row(s) with invalid area removed."
        )

    def fill_missing(self):
        if not self.require_data():
            return

        before = int(self.df.isna().sum().sum())
        if before == 0:
            messagebox.showinfo("Missing Values", "No missing values to fill.")
            return

        for column in self.df.columns:
            if not self.df[column].isna().any():
                continue

            if pd.api.types.is_numeric_dtype(self.df[column]):
                median_value = self.df[column].median()
                if pd.notna(median_value):
                    self.df[column] = self.df[column].fillna(median_value)
            else:
                modes = self.df[column].mode(dropna=True)
                if not modes.empty:
                    self.df[column] = self.df[column].fillna(modes.iloc[0])
                else:
                    self.df[column] = self.df[column].fillna("Unknown")

        after = int(self.df.isna().sum().sum())
        self.write_log(
            f"Missing-value handling: {before - after} cell(s) filled. "
            "Numeric columns use median; text columns use mode or 'Unknown'."
        )
        self.update_stats()
        self.display_dataframe(self.df, "Current Cleaned Dataset")
        messagebox.showinfo(
            "Missing Values Handled",
            f"Missing cells before: {before}\nMissing cells remaining: {after}"
        )

    def trim_text(self):
        if not self.require_data():
            return

        changed = 0
        city_map = {
            "la": "Los Angeles",
            "l.a.": "Los Angeles",
            "nyc": "New York City",
            "sf": "San Francisco",
            "chicgo": "Chicago",
            "bostan": "Boston",
        }

        for column in self.df.select_dtypes(include=["object", "string"]).columns:
            old = self.df[column].copy()
            self.df[column] = self.df[column].apply(
                lambda value: value.strip() if isinstance(value, str) else value
            )
            if column.lower() == "location":
                self.df[column] = self.df[column].apply(
                    lambda val: city_map.get(val.lower(), val.title()) if isinstance(val, str) else val
                )
            changed += int((old != self.df[column]).fillna(False).sum())

        self.write_log(f"Text cleaning completed: {changed} value(s) standardized.")
        self.update_stats()
        self.display_dataframe(self.df, "Current Cleaned Dataset")
        messagebox.showinfo("Text Cleaning",
                            f"Text standardized for {changed} value(s).")

    def convert_numeric(self):
        if not self.require_data():
            return

        converted_columns = []
        for column in self.df.select_dtypes(include=["object", "string"]).columns:
            original_non_missing = self.df[column].notna().sum()
            converted = pd.to_numeric(self.df[column], errors="coerce")
            converted_non_missing = converted.notna().sum()

            # Convert only when all existing non-missing values can be parsed.
            if original_non_missing > 0 and converted_non_missing == original_non_missing:
                self.df[column] = converted
                converted_columns.append(str(column))

        if converted_columns:
            message = "Converted columns: " + ", ".join(converted_columns)
        else:
            message = "Numeric columns are already in numeric format."

        self.write_log("Data type check: " + message)
        self.update_stats()
        self.display_dataframe(self.df, "Current Cleaned Dataset")
        messagebox.showinfo("Data Type Conversion", message)

    def handle_outliers(self):
        if not self.require_data():
            return

        numeric_columns = list(
            self.df.select_dtypes(include="number").columns
        )
        if not numeric_columns:
            messagebox.showinfo("Outliers",
                                "No numeric columns are available.")
            return

        changed_total = 0
        for column in numeric_columns:
            series = self.df[column].dropna()
            if len(series) < 4:
                continue

            q1 = series.quantile(0.25)
            q3 = series.quantile(0.75)
            iqr = q3 - q1

            if pd.isna(iqr) or iqr == 0:
                continue

            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr

            mask = self.df[column].notna() & (
                (self.df[column] < lower) | (self.df[column] > upper)
            )
            count = int(mask.sum())

            if count:
                self.df.loc[mask, column] = self.df.loc[
                    mask, column
                ].clip(lower=lower, upper=upper)
                changed_total += count

        self.write_log(
            f"Outlier handling: {changed_total} value(s) capped "
            "using the IQR method."
        )
        self.update_stats()
        self.display_dataframe(self.df, "Current Cleaned Dataset")
        messagebox.showinfo(
            "Outlier Handling",
            f"{changed_total} outlier value(s) capped using the IQR method.\n"
            "Review the results to ensure this treatment is suitable "
            "for your dataset."
        )

    def clean_column_names(self):
        if not self.require_data():
            return

        old_columns = list(self.df.columns)
        new_columns = []
        for column in old_columns:
            name = str(column).strip().lower()
            name = "_".join(name.split())
            name = name.replace("-", "_")
            new_columns.append(name)

        counts = {}
        unique_columns = []
        for name in new_columns:
            counts[name] = counts.get(name, 0) + 1
            unique_columns.append(
                name if counts[name] == 1 else f"{name}_{counts[name]}"
            )

        self.df.columns = unique_columns
        self.write_log("Column names cleaned: trimmed, lowercased and standardized.")
        self.update_stats()
        self.display_dataframe(self.df, "Current Cleaned Dataset")
        messagebox.showinfo("Column Names",
                            "Column names have been standardized.")

    def reset_data(self):
        if self.original_df is None:
            self.require_data()
            return

        confirm = messagebox.askyesno(
            "Reset Dataset",
            "This will discard all cleaning changes and restore the original data. Continue?"
        )
        if confirm:
            self.df = self.original_df.copy(deep=True)
            self.update_stats()
            self.display_dataframe(self.df, "Restored Original Dataset")
            self.write_log("Dataset reset to the original uploaded version.")

    def save_dataset(self):
        if not self.require_data():
            return

        path = filedialog.asksaveasfilename(
            title="Save cleaned dataset",
            defaultextension=".csv",
            filetypes=[
                ("CSV file", "*.csv"),
                ("Excel workbook", "*.xlsx")
            ],
            initialfile="cleaned_dataset.csv"
        )
        if not path:
            return

        try:
            extension = Path(path).suffix.lower()
            if extension == ".csv":
                self.df.to_csv(path, index=False, encoding="utf-8-sig")
            elif extension == ".xlsx":
                self.df.to_excel(path, index=False)
            else:
                messagebox.showerror(
                    "Unsupported Format",
                    "Please save as CSV or XLSX."
                )
                return

            self.write_log(f"Cleaned dataset saved to: {path}")
            messagebox.showinfo("Saved",
                                f"Cleaned dataset saved successfully:\n{path}")
        except Exception as error:
            messagebox.showerror("Save Error", str(error))


if __name__ == "__main__":
    root = tk.Tk()
    app = DataCleaningApp(root)
    root.mainloop()