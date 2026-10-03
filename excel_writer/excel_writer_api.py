from openpyxl.styles import Font, Alignment
from api.api_client import get_api_data
from utils.logger import logger_setup
from config.config import Config
from openpyxl.utils import get_column_letter
import os
import pandas as pd

# CREATE THE LOGGER OBJECT
logger = logger_setup()

# CREATE OUTPUT FOLDER
os.makedirs(Config.OUTPUT_FOLDER, exist_ok=True)

# CREATE THE FILE SAVE PATH
save_path = os.path.join(Config.OUTPUT_FOLDER, Config.API_WORKBOOK_FILENAME)

def get_products_data():
    api_data = get_api_data()

    if api_data is not None:
        data = pd.DataFrame(api_data['products'])
        data = data[Config.PRODUCTS_DATA]

        # SAVE THE DATA TO EXCEL
        try:
            with pd.ExcelWriter(save_path) as writer:
                data.to_excel(writer, sheet_name=Config.API_SHEET_FILENAME, index=False)

                # ACCESS THE WORKSHEET
                worksheet = writer.sheets[Config.API_SHEET_FILENAME]

                # FORMATTING THE WORKSHEET USING OPENPYXL
                for cell in worksheet[1]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')

                    # FIND THE PRICE COLUMN NUMBER
                    if cell.value == 'price':
                        price_column_letter = get_column_letter(cell.column)

                    # FIND THE DESCRIPTION COLUMN NUMBER
                    if cell.value == 'description':
                        description_column_letter = get_column_letter(cell.column)

                # NUMBER FORMATTING IN PRICE COLUMN
                for cell in worksheet[price_column_letter][1:]:
                        cell.number_format = "#,##0.00"

                # WRAPPING FORMATTING IN DESCRIPTION COLUMN
                for cell in worksheet[description_column_letter][1:]:
                    cell.alignment = Alignment(wrap_text=True)

                # GETTING THE MAX LENGTH IN EACH COLUMN
                for column_number, column in enumerate(worksheet.columns, start=1):
                    max_length = 0

                    for cell in column:
                        if cell.value is not None:
                            length = len(str(cell.value))
                            max_length = max(max_length, length)

                    # GET THE COLUMN LETTER
                    column_letter = get_column_letter(column_number)

                    # ADJUST THE COLUMN WIDTH
                    worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

        except PermissionError:
            print(f'Unable to save {Config.API_SHEET_FILENAME}. Please close the file and try again.')
    else:
        logger.info('Cannot continue processing, API data is empty.')

if __name__ == "__main__":
    get_products_data()