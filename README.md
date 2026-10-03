# Automated Sales Intelligence System

A Python-based sales automation system that processes sales data from Excel files and external APIs, performs data analysis, and generates formatted Excel reports.

## Features

- Read and process multiple Excel sales files
- Validate required columns
- Combine sales data using pandas
- Calculate total sales
- Generate sales summaries
- Generate sales-by-person reports
- Retrieve product data from an external REST API
- Generate formatted Excel reports
- Automatic column width adjustment
- Excel number formatting
- Logging
- Error handling
- Modular project structure

## Technologies Used

- Python
- pandas
- openpyxl
- requests
- REST API
- Git & GitHub

## Project Structure

```text
automated_sales_system/
│
├── api/
│   └── api_client.py
│
├── config/
│   └── config.py
│
├── excel_writer/
│   ├── excel_writer_api.py
│   └── excel_writer_sales.py
│
├── sales/
│   └── sales_data.py
│
├── utils/
│   └── logger.py
│
├── data/
│   └── Excel input files
│
├── output/
│   └── Generated Excel reports
│
├── logs/
│   └── Application log files
│
├── main.py
├── .gitignore
└── README.md