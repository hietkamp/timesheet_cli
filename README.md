# Timesheet App (CLI)

A robust command-line interface (CLI) application written in Rust for tracking daily work hours, managing project templates, and exporting professional monthly timesheets to Excel.

## 🚀 Features

* **Template Management**: Create, edit, and delete default daily hour templates for recurring projects.
* **Time Logging**: Track hours per project on a weekly basis (ISO weeks). Support for auto-filling from templates.
* **Monthly Overview**: View a matrix report (Projects vs. Days) for any given month directly in the terminal.
* **Excel Export**: Generate a formatted, professional Excel timesheet (Dutch format: *Urenstaat*) ready for invoicing or signing.
* **Database**: Uses SQLite (`timesheet.db`) for persistent local storage.

## 🛠️ Prerequisites

Before running the application, ensure you have the following installed:

* **Rust & Cargo**: [Install Rust](https://www.rust-lang.org/tools/install)

## ⚙️ Configuration & Setup

**Crucial Step:** This application requires specific external files to function correctly, particularly for the Excel export feature.

### 1. The `.env` File
Create a `.env` file in the root directory of your project to configure the employee details used in the Excel export.

**Example `.env` content:**
```env
EMPLOYEE_NAME="John Doe"
EMPLOYEE_TITLE="Software Engineer"
EMPLOYEE_PHONE="+31 6 12345678"
PATH="/Users/<username>/Downloads"
```

### 2. Image Assets

The Excel export function looks for two specific images in the project root directory. You must add these files or the export may fail/look incomplete.

- logo.jpg: Company logo (displayed at the top of the timesheet).
- signature.png: Your digital signature (placed at the bottom of the timesheet).

Note: The code attempts to scale these images to fit specific dimensions (300x200).