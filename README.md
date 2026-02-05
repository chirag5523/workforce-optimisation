Workforce Optimisation Data Processor


📌 Overview

This repository contains a specialized ETL (Extract, Transform, Load) tool designed to automate the consolidation of workforce performance reports. The script handles the common "real-world" problem of inconsistent Excel formatting by dynamically locating headers and cleaning data on the fly.


🛠️ Key Technical Features
1. Dynamic Folder Selection (Tkinter UI)

Instead of hardcoding file paths which change between users, the script uses a tkinter folder dialog.

    Benefit: Ensures the script is portable and "plug-and-play" for any team member.

    Implementation: The script initializes a "hidden" root window to bring the folder picker to the foreground.

2. Intelligent Header Detection

Excel exports often have empty rows or metadata at the top.

    The Logic: The script scans every row until it finds a cell containing "Agent Name".

    The Result: It dynamically resets the dataframe header to that row, ensuring data alignment regardless of the export version.

3. Automated Subfolder Filtering

To prevent processing old data, the script uses Regular Expressions (Regex) to identify folder names starting with a month number (e.g., 09. September). It currently filters for folders where the month index is 9 or higher.
4. Data Sanitization & Split Export

    Sanitization: Removes rows where "Agent Name" or "Team Name" are blank.

    Category Splitting: The script identifies unique values in the Call Productivity (Talk+Wait) column and generates a separate Excel file for each category, which is essential for distributed team reporting.




🚀 Getting Started
Prerequisites

    Python 3.8+

    Pandas & OpenPyXL libraries

Installation

    Clone the repository:
    Bash

    git clone https://github.com/your-username/workforce-optimisation.git
    cd workforce-optimisation

    Install dependencies:
    Bash

    pip install -r requirements.txt

Running the Script

    Execute the script:
    Bash

    python workforce_process.py

    Select Folder: A Windows explorer window will appear. Select the Data folder containing your monthly subfolders.

    Check Outputs: The script will print progress to the console and save the new Excel files (prefixed with Output_) in the directory above your data folder.




# Filtering rows with valid content only
df = df[
    (df["Agent Name"].notna()) & (df["Agent Name"].astype(str).str.strip() != "")
]

This handles "dirty" data by stripping whitespace and removing null rows that often appear at the end of Excel exports.
🔒 Security & Privacy

Important: This tool is designed for corporate environments.

    No Data in Cloud: The .gitignore file is strictly configured to ensure that no .xlsx files are ever pushed to GitHub.

    Anonymized Logic: All internal server paths and usernames have been replaced with dynamic system calls (os.path.expanduser).
