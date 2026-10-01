# Sales Data Analysis & Automation

A Python project that analyzes sales data from an Excel file and automatically generates a sales report.

## About the Project

This project uses Python, Pandas, and OpenPyXL to process sales data and generate useful business insights automatically.

The program reads sales data from an Excel file, calculates sales, analyzes products and customers, and creates an automated Excel report.

## Features

- Reads sales data from Excel
- Cleans product names and customer names
- Calculates Sales = Quantity × Price
- Calculates total sales
- Calculates sales by product
- Calculates sales by customer
- Identifies the best-selling product
- Identifies the top customer
- Generates an automated Excel report
- Creates a separate Summary sheet
- Supports re-running without duplicate Summary sheet errors

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Excel / WPS Office

## Input Data

The input Excel file (`sales.xlsx`) contains:

- Date
- Product
- Quantity
- Price
- Customer

## Output

The program generates:

`automated_sales_report.xlsx`

The report contains:

- Sales data
- Summary sheet
- Total sales
- Best-selling product
- Best product sales
- Top customer
- Top customer sales

## Example Results

Using the sample dataset:

- Total Sales: 308,900
- Best-Selling Product: Laptop
- Best Product Sales: 220,000
- Top Customer: Customer A
- Top Customer Sales: 173,000

## How to Run

Install the required libraries:

```bash
pip install pandas openpyxl
```

Run the program:

```bash
python business_sales.py
```

Make sure `sales.xlsx` is in the same folder as the Python script.

## What I Learned

Through this project, I practiced:

- Reading and writing Excel files using Python
- Data cleaning with Pandas
- Grouping and aggregation using `groupby()`
- Calculating business metrics
- Creating automated reports
- Working with OpenPyXL
- Writing reusable automation scripts

## Project Goal

The goal of this project is to understand how Python can automate repetitive Excel-based data analysis and reporting tasks.
