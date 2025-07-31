import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from collections import defaultdict


def split_name(name):
    if pd.isna(name):
        return "", ""  # Returns empty string if LC1, LC2, or Admin name is missing
    parts = name.strip().split()  # Separates first and last name, strips whitespace
    return parts[0], (
        " ".join(parts[1:]) if len(parts) > 1 else ""
    )  # Returns first and last name, returns last name as empty if there is only the first name


def process_file(filepath): # Processes CSV File
    df = pd.read_csv(
        filepath, encoding="ISO-8859-1"
    )  # Initalizes DataFrame for working with spreadsheet, ensures file can be encoded and opened
    active_teams = df[
        df["Active Team"].str.strip() == "Active"
    ]  # Filters out inactive teams

    contacts = []

    for _, row in active_teams.iterrows():  # Iterates through Excel spreadsheet rows
        program = str(row.get("Program", "")).strip().upper()
        team_num = str(int(row["Team Number"])) if pd.notna(row["Team Number"]) else ""
        team_id = f"{program}{team_num}"

        for role in [
            ("LC1", "LC1 Name", "LC1 Email"),
            ("LC2", "LC2 Name", "LC2 Email"),
            ("Admin", "Team Admin Name", "Team Admin Email"),
        ]:
            _, name_col, email_col = role
            if pd.isna(row.get(email_col)) or row[email_col] == "":
                continue

            email = row[email_col]
            first, last = split_name(row.get(name_col, ""))
            contact = {
                "Email Address": email,
                "First Name": first,
                "Last name": last,
                "City": row.get("Team City", ""),
                "County": row.get("Team County", ""),
                "Zip Code": row.get("Team Postal Code", ""),
                "Country": row.get("Team Country", ""),
                "State": row.get("Team State Province", ""),
                "Team IDs": [team_id],
                "Programs": [program]
            }
            contacts.append(contact)

    email_dict = defaultdict(lambda: { # Ensures contacts aren't added multiple times
        "Email Address": "",
        "First Name": "",
        "Last name": "",
        "City": "",
        "County": "",
        "Zip Code": "",
        "Country": "",
        "State": "",
        "Team IDs": set(), # Sets ensure no duplicates are added to contact's information
        "Programs": set()
    })

    for c in contacts: # Adds contacts to record
        email = c["Email Address"]
        record = email_dict[email] # References one's email as the start of a new contact's record
        record["Email Address"] = email
        record["First Name"] = c["First Name"]
        record["Last name"] = c["Last name"]
        record["City"] = c["City"]
        record["County"] = c["County"]
        record["Zip Code"] = c["Zip Code"]
        record["Country"] = c["Country"]
        record["State"] = c["State"]
        record["Team IDs"].update(c["Team IDs"]) # Updates Team IDs so multiple team numbers can be added without duplication of existing ones
        record["Programs"].update(c["Programs"]) # Updates existing program to avoid duplications
    
    output_rows = []
    for contact in email_dict.values(): #Loop through every contact's information (record)
        teams = list(contact["Team IDs"]) # Converts Team IDs to a list
        programs = list(contact["Programs"])
        tags = []
        affiliations = []
        for prog in programs:
            tags.extend([prog, f"{prog} coach"]) # Adds coach to end of their program (FTC or FRC) for MailChimp tags
            affiliations.append(f"{prog} Coach/Mentor") # Adds Coach/Mentor to program(s) for MailChimp Affiliations
            
            row = {
    "Email Address": contact["Email Address"],
    "First Name": contact["First Name"],
    "Last name": contact["Last name"],
    "Affiliation": ", ".join(affiliations), # Join strings together in case of multiple affiliations
    "City": contact["City"],
    "County": contact["County"],
    "Zip Code": contact["Zip Code"],
    "Team 1 Type & Number": teams[0] if len(teams) > 0 else "", # Adds multiple teams, leaves nonexistent teams empty
    "Team 2 Type & Number": teams[1] if len(teams) > 1 else "",
    "Team 3 Type & Number": teams[2] if len(teams) > 2 else "",
    "Team 4 Type & Number": teams[3] if len(teams) > 3 else "",
    "Country": contact["Country"],
    "State": contact["State"],
    "Tags": ", ".join(tags)
}

        output_rows.append(row) # Adds this row to list of all rows

    output_df = pd.DataFrame(output_rows) # Converts rows into panda DataFrame
    output_file = os.path.join(os.path.dirname(filepath), "mailchimp_contacts.csv") # Constructs path for saving file in the same folder as original file, but renamed
    output_df.to_csv(output_file, index=False) # Saves DataFrame as CSV file
    return output_file # Shows success message and where file was saved

def process_txt_file(filepath):
    try:
        df = pd.read_csv(filepath, sep="\t", encoding="utf-16")
        df.columns = df.columns.str.strip()  # Remove spaces from column headers

        email_dict = defaultdict(lambda: { # Stores records of contacts
            "Email Address": "",
            "First Name": "",
            "Last name": "",
            "Affiliation": "Volunteer",
            "City": "",
            "Zip Code": "",
            "Country": "",
            "State": "",
            "Tags": set()
        })

        for _, row in df.iterrows():
            email = str(row.get("Email", "")).strip().lower()
            if not email: # Ensures no duplication of contacts
                continue

            record = email_dict[email]
            record["Email Address"] = email
            record["First Name"] = row.get("Preferred Name", "").strip()
            record["Last name"] = row.get("Last Name", "").strip()
            record["City"] = row.get("City", "").strip()
            record["Zip Code"] = str(row.get("Postalcode", "")).strip()
            record["Country"] = row.get("Country", "").strip()
            record["State"] = row.get("State/Province", "").strip()

            employer = row.get("Current Employer", "")
            if pd.notna(employer) and employer.strip():
                record["Affiliation"] = f"Volunteer, {employer.strip()}" # Adds volunteer and current employer to affiliation

            # Tags
            program = row.get("Program", "").strip().upper()
            if program:
                print(f"PROGRAM VALUE: '{row.get('Program')}' → '{program}'")
                record["Tags"].add(program)

            roles = str(row.get("Volunteer Roles", "")).lower() # Gets tags from Volunteer Roles column
            if "judge" in roles:
                record["Tags"].add("Judge")
            if "referee" in roles:
                record["Tags"].add("Referee")
            if "fta" in roles:
                record["Tags"].add("FTA")
            if "csa" in roles:
                record["Tags"].add("CSA")


        output_rows = [] # Final outputted rows
        for record in email_dict.values():
            output_rows.append({
                "Email Address": record["Email Address"],
                "First Name": record["First Name"],
                "Last name": record["Last name"],
                "Affiliation": record["Affiliation"],
                "City": record["City"],
                "Zip Code": record["Zip Code"],
                "Country": record["Country"],
                "State": record["State"],
                "Tags": ", ".join(sorted(record["Tags"]))
            })

        output_df = pd.DataFrame(output_rows)
        output_file = os.path.join(os.path.dirname(filepath), "mailchimp_volunteers.csv")
        output_df.to_csv(output_file, index=False)
        return output_file

    except Exception as e:
        raise RuntimeError(f"Failed to process volunteer file: {e}")









    



def launch_gui():
    root = tk.Tk()
    root.withdraw()  # Hides default window

    file_path = filedialog.askopenfilename(  # Opens file selection and returns path to the selected file as a string
        title="Select CSV file from FIRST Tableu or .txt volunteer file to modify for MailChimp",
        filetypes=[("CSV or TXT Files", "*.csv *.txt"), ("CSV Files", "*.csv"), ("Text Files", "*.txt")],
    )
    if not file_path:
        return  # Prevents program from crashing when dialog is cancelled
    
    try:
        if file_path.lower().endswith(".txt"):
            out_file = process_txt_file(file_path)
        else:
            out_file = process_file(file_path)
        messagebox.showinfo("Success", f"Mailchimp contacts saved to:\n{out_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process file:\n{e}")


if __name__ == "__main__":
    launch_gui()
