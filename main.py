from utils.logger import logger_setup
from excel_writer.excel_writer_api import get_products_data
from excel_writer.excel_writer_sales import sales_report

# MAIN PROGRAM

# SET THE LOGGER OBJECT
logger = logger_setup()

def main():
    success = True

    logger.info('Main program has started.')

    # RUN THE EXCEL WRITER API MODULE
    try:
        get_products_data()
    except Exception as e:
        logger.error(f'Running in main program, get_products_data() has error.'
                     f'Error details: {e}.')
        success = False

    # RUN THE EXCEL WRITER SALES
    try:
        sales_report()
    except Exception as e:
        logger.error(f'Running in main program, sales_report() has error.'
                     f'Error details: {e}.')
        success = False

    # LOG THE PROGRAM IF SUCCESSGFUL OR NOT
    if success:
        logger.info('Program successfully finished.')
    else:
        logger.error(f'Program finished with errors')

if __name__ == "__main__":
    main()