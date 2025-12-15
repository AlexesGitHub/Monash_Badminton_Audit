# 🏸 Monash Badminton Club Membership Auditor

This Python script is designed to automate the process of cross-referencing member registration data between the current participant list (2025 Audit Sheet) and the next registration list (2026 Registration Sheet).

It updates the 2025 Audit Sheet by flagging re-registered members and performing a multi-level verification of their student membership status based on email, student ID, and user type.

## 🚀 Setup and Installation

This script requires Python 3 and the `pandas` library for efficient Excel file handling.

### 1. Prerequisites

You need to have Python installed on your system.

### 2. Create the Virtual Environment (.venv)

A virtual environment isolates this project's dependencies from your main Python installation.

1.  **Open your terminal** and navigate to your project folder (`badminton-audit/`).
2.  Run the following command to create the environment named `venv`:

```bash
python -m venv venv
```

### 3. Activate the Virtual Environment

You must activate the environment every time you want to run the script or install packages.

* **If you are using Windows PowerShell:**
```bash
.\venv\Scripts\Activate.ps1
```
* **If you are using Windows Command Prompt (CMD) or Git Bash:**
```bash
.\venv\Scripts\activate
```
*Note: Your prompt should now start with `(venv)`.*

### 4. Install Dependencies

With the `(venv)` active, install the required libraries:

```bash
pip install pandas openpyxl
```
### 5. Create the file structure
Place the Python script (matchChecker.py) and your two Excel files in a single, dedicated folder (e.g., badminton-audit/).

Your folder structure must look like this for the script to find the files:
```
badminton-audit/
├── matchChecker.py
├── 202512140830_MonashBadmintonClubMembers.xlsx    <-- 2026 Registration Sheet
├── MONASHBADDY Member Audit Sheet.xlsx             <-- 2025 Audit Sheet
├── venv/ <-- Created in Step 2
```

### ⚙️ Expected Sheet Setup and Column Headers
The script relies on exact column header names for the validation logic to work. Please ensure your Excel sheets use the following headers (case-sensitive):

### A. 2026 Registration Sheet (`202512140830_MonashBadmintonClubMembers.xlsx`)

| Column Header | Purpose in Script |
| :--- | :--- |
| `First Name` | Used to create the unique Name Key for cross-referencing. |
| `Last Name` | Used to create the unique Name Key for cross-referencing. |
| `Email` | Used to check for the student domain (`@student.monash.edu`). |
| `Student ID` | Used to confirm student status (must be present). |
| `User Type` | Used to identify members marked as `Monash Student` for special cases. |

### B. 2025 Audit Sheet (`MONASHBADDY Member Audit Sheet.xlsx`)

| Column Header | Purpose in Script | Statuses Set by Script |
| :--- | :--- | :--- |
| `name` | The full name of the member (used to create the Name Key). | N/A |
| `On UniOne?` | **[UPDATE TARGET 1]** Tracks re-registration status. | `yes`, `no` (preserves existing custom values). |
| `Selected correct membership type (student/general)?` | **[UPDATE TARGET 2]** Tracks student verification status. | `yes`, `Missing Student ID`, `Requires Validation` |

### ▶️ How to Run the Script
1. Open your terminal and navigate to the badminton-audit folder.

2. Activate your virtual environment (the prompt should start with (venv)):
. .\venv\Scripts\Activate.ps1

3. Execute the Python script:
python matchChecker.py

### Script Output
The script executes two key actions:

1. Saves a New File: 
It creates a new Excel file named MONASHBADDY_Audit_Sheet_UPDATED.xlsx in the same directory,
containing all the updates.

2. Prints an Audit Summary: It prints a detailed, case-insensitive summary to the console, 
showing exactly which members were updated in each of the verification categories 
(Yes, Missing Student ID, Requires Validation).

### 📝 Validation Logic Summary

The script prioritizes the verification status based on the following checks for all members
who have re-registered (On UniOne? is set to yes):

| Status Applied | Conditions Met (in 2026 Sheet) | Priority |
| :--- | :--- | :--- |
| `Missing Student ID` | Email is `@student.monash.edu` AND Student ID is missing (empty/NaN). | Highest |
| `Requires Validation` | User Type is `Monash Student` BUT the Email does NOT end in `@student.monash.edu`. | Medium |
| `yes` | Email is `@student.monash.edu` AND Student ID is present. | Lowest (Default Success) |

