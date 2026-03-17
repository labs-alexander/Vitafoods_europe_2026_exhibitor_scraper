import os
import math
import pandas as pd
import win32com.client  # For resolving Windows shortcut (.lnk) paths
from tkinter import Tk, filedialog

# ================== Configuration ===================
CONFIG = {
    "MAX_COLUMN_WIDTH": 100,
    "COLUMN_PADDING": 5,
    "ROW_HEIGHT": 15,
    "HEADER_BG": "#2F5597",
    "HEADER_FG": "white",
    "ZEBRA_BG": "#FFFFFF",
    "NORMAL_ROW_BG": "#D9E1F2",
    "FONT_NAME": "Arial",
    "FONT_SIZE": 12
}

# ================== Main Loop ======================
def main():
    # Continuous menu loop until the user selects exit
    while True:
        print("\n===== DATA CLEANER =====")
        print("1 - Process File")
        print("2 - Exit")
        choice = input("Choose an option: ").strip()

        if choice == "1":
            process_file()
        elif choice == "2":
            break
        else:
            print("Invalid option.")


# ================== File Processing =================
def process_file():
    try:
        file_path = resolve_input_path()  # Handles drag & drop, file picker, or shortcuts
        df, output_file = load_data(file_path)  # Reads Excel or CSV
        format_file(df, output_file)  # Writes formatted Excel output
        print(f"✅ Report generated successfully: {output_file}")
    except FileNotFoundError:
        print("❌ File not found.")
    except PermissionError:
        print("❌ Close Excel before generating the report.")
    except Exception as error:
        print(f"❌ Unexpected error: {error}")


# ================== Input Resolution =================
def resolve_input_path():
    # Accepts either drag & drop path or file picker selection
    print("\nDrag file here or press ENTER to open file dialog.")
    path = input("> ").strip().strip('"').strip("'")

    if not path:
        path = open_file_dialog()

    path = os.path.abspath(path)
    path = resolve_windows_shortcut(path)  # Resolve .lnk to actual target

    if not os.path.exists(path):
        raise FileNotFoundError(path)

    return path


def open_file_dialog():
    # Opens a GUI dialog to select Excel/CSV file
    root = Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Select Excel or CSV file",
        filetypes=[("Excel or CSV", "*.xlsx *.xls *.csv")]
    )
    return path


def resolve_windows_shortcut(path):
    # Converts Windows shortcut (.lnk) to actual file path
    if not path.lower().endswith(".lnk"):
        return path
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(path)
    return shortcut.Targetpath


# ================== Data Loading ===================
def load_data(file_path):
    # Determines file type and loads into pandas DataFrame
    name, extension = os.path.splitext(file_path)
    extension = extension.lower()

    if extension == ".csv":
        df = pd.read_csv(file_path)
    elif extension in (".xlsx", ".xls"):
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format.")

    # Generate output filename in the same folder
    output_file = f"{name}_premium_report.xlsx"
    return df, output_file


# ================== Excel Formatting =================
def format_file(dataframe, output_file):
    # Uses xlsxwriter to apply formatting, column widths, and row heights
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, sheet_name="premium_report", index=False)
        workbook = writer.book
        worksheet = writer.sheets["premium_report"]

        # Define cell formats
        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": CONFIG["HEADER_BG"],
            "font_color": CONFIG["HEADER_FG"],
            "border": 1
        })
        zebra_fmt = workbook.add_format({
            "font_name": CONFIG["FONT_NAME"],
            "font_size": CONFIG["FONT_SIZE"],
            "bg_color": CONFIG["ZEBRA_BG"],
            "border": 1,
            "valign": "top",
            "text_wrap": True
        })
        normal_fmt = workbook.add_format({
            "font_name": CONFIG["FONT_NAME"],
            "font_size": CONFIG["FONT_SIZE"],
            "bg_color": CONFIG["NORMAL_ROW_BG"],
            "border": 1,
            "valign": "top",
            "text_wrap": True
        })

        # Dynamically calculate column widths based on content
        column_widths = {}
        for col_idx, col_name in enumerate(dataframe.columns):
            worksheet.write(0, col_idx, col_name, header_fmt)
            max_len = max(dataframe[col_name].astype(str).str.len().max(), len(col_name))
            width = min(CONFIG["MAX_COLUMN_WIDTH"], max_len + CONFIG["COLUMN_PADDING"])
            worksheet.set_column(col_idx, col_idx, width)
            column_widths[col_idx] = width

        # Dynamically calculate row heights based on cell text wrapping
        for row_idx in range(1, len(dataframe) + 1):
            row_fmt = zebra_fmt if row_idx % 2 == 0 else normal_fmt
            max_height = CONFIG["ROW_HEIGHT"]

            for col_idx in range(len(dataframe.columns)):
                text = str(dataframe.iloc[row_idx - 1, col_idx])
                width = column_widths[col_idx]
                lines = math.ceil(len(text) / width)
                height = max(CONFIG["ROW_HEIGHT"], lines * CONFIG["ROW_HEIGHT"])
                if height > max_height:
                    max_height = height

            worksheet.set_row(row_idx, max_height, row_fmt)


# ================== Entry Point ====================
if __name__ == "__main__":
    main()