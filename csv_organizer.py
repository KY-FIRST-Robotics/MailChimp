import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os


def launch_gui():
    root = tk.Tk()
    root.withdraw() # Hides default window

    file_path = filedialog.askopenfilename( # Opens file selection and returns path to the selected file as a string
        title="Select CSV file from FIRST Tableu to modify for MailChimp",
        filetypes=[("CSV Files", "*.csv")]
    )
    if not file_path:
        return # Prevents program from crashing when dialog is cancelled

if __name__ == "__main__":
    launch_gui()