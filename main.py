import tkinter as tk
from tkinter import filedialog
from rr_word_html import DocxToHtml


def browse_and_convert(full=False):
    # Initialize Tkinter root widget
    root = tk.Tk()
    root.withdraw()  # We don't want a full GUI, so keep the root window from appearing

    # Show an "Open" dialog box and return the path to the selected file
    filepath = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])

    # Check if a file was selected
    if filepath:
        print(f"File selected: {filepath}")
        # Here you would call your conversion function, e.g., docx_to_html(filepath)
        word_to_html = DocxToHtml(full)
        word_to_html.convert(filepath)

    else:
        print("No file selected.")


def main():
    # full means include html tags, link to stylesheet and body tags

    full = False
    browse_and_convert(full)


if __name__ == "__main__":
    main()
