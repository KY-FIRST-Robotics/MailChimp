import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def split_name(name):
    if pd.isna(name): return "", "" # Returns empty string if LC1, LC2, or Admin name is missing
    parts = name.strip().split() # Separates first and last name, strips whitespace
    return parts[0], " ".join(parts[1:]) if len(parts) > 1 else "" # Returns first and last name, returns last name as empty if there is only the first name

def process_file(filepath):
    df = pd.read_csv(filepath, encoding="ISO-8859-1") # Initalizes DataFrame for working with spreadsheet, ensures file can be encoded and opened
    active_teams = df[df["Active Team"].str.strip() == "Active"] # Filters out inactive teams

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