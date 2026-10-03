import pandas as pd
import os
from config.config import Config
from utils.logger import logger_setup

# CREATE LOGGER OBJECT
logger = logger_setup()

def get_excel_data():
    # READ THE FILES INSIDE THE DATA FOLDER
    file_lists = os.listdir(Config.DATA_FOLDER)

    # EMPTY DATAFRAME LIST
    data_frames = []

    # PROCESS THE EXCEL FILES
    if file_lists:
        for file in file_lists:
        
            if file.endswith(Config.EXCEL_EXT):
                
                # CREATE THE FILE PATH
                file_path = os.path.join(Config.DATA_FOLDER, file)

                # READ THE FILE DATA
                try:
                    data_frame = pd.read_excel(file_path)

                    actual_columns = data_frame.columns.to_list()

                    if all(column in actual_columns for column in Config.SALES_REQUIRED_COLUMNS):
                        data_frames.append(data_frame)
                    else:
                        logger.info(f'File {file} has missing columns.')

                except FileNotFoundError:
                    logger.error(f'File {file} cannot be found and processed.')
                except PermissionError:
                    logger.error(f'File {file} cannot be accessed and processed.')

            else:
                logger.info(f'File {file} is not an excel file and cannot be processed.')

    else:
        logger.info('Empty data folder.')

    # COMBINE THE DATA FRAMES
    if data_frames:
        combine_data = pd.concat(data_frames, ignore_index=True)

        # CREATE A SALES COLUMN
        combine_data['Sales'] = combine_data['Quantity'] * combine_data['Unit Price']

        # GROUP BY SALES OF SALESPERSON
        sales_by_person = combine_data.groupby('Sales Person')['Sales'].sum().reset_index()

        return combine_data, sales_by_person
    else:
        logger.info('No data frame can be processed.')
        return None, None

if __name__ == "__main__":
    get_excel_data()