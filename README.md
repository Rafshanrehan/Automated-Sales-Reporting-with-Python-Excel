# Automated-Sales-Reporting-with-Python-Excel
This project demonstrates a simple Python-based automation workflow for processing sales data and generating Excel reports with pivot tables and bar charts.

## Features

- Reads raw sales data (`supermarket_sales.xlsx`) with Pandas.
- Creates a pivot table summarizing **Total Sales by Gender and Product Line**.
- Automatically calculates column totals.
- Generates a **bar chart** visualizing sales by product line.
- Supports running as a **standalone executable** for users who don't want to open Python.

## Technologies Used

- **Python** – core scripting and data manipulation.
- **Pandas** – for reading Excel and creating pivot tables.
- **OpenPyXL** – for Excel file handling, formatting, and chart creation.
- **Tkinter (optional)** – for GUI input dialogs in the executable version.
