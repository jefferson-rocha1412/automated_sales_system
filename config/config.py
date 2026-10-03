class Config:
    # API SETTINGS
    API_URL = "https://dummyjson.com/products"
    API_TIMEOUT = 10

    # PRODUCTS DATA
    PRODUCTS_DATA = ['id', 'title', 'description', 'category', 'price', 'discountPercentage', 'brand']

    # LOGGER SETTINGS
    LOGGER_FILENAME = 'program_logs.log'
    LOGGER_FOLDER = 'logs'
    LOGGER_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

    # EXCEL SETTINGS
    API_WORKBOOK_FILENAME = 'api_products.xlsx'
    API_SHEET_FILENAME = 'products_summary'
    SALES_REPORT_WORKBOOK_FILENAME = 'sales_report.xlsx'
    SALES_REPORT_SUMMARY_SHEET = 'summary'
    SALES_REPORT_SALESBYPERSON_SHEET = 'sales_by_person'

    # FILE EXTENSIONS
    EXCEL_EXT = ('.xls', '.xlsx')

    # FOLDERS
    DATA_FOLDER = 'data'
    OUTPUT_FOLDER = 'output'

    # SALES COLUMN
    SALES_REQUIRED_COLUMNS = ['No.', 'Date', 'Product ID', 'Quantity', 'Unit Price', 'Sales Person']