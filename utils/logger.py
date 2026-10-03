import logging
import os
from config.config import Config

def logger_setup():
    # CREATE THE FOLDER IF NOT YET EXISTED
    os.makedirs(Config.LOGGER_FOLDER, exist_ok=True)

    # CREATE THE FILE PATH OF THE LOG FILE
    logger_file_path = os.path.join(Config.LOGGER_FOLDER, Config.LOGGER_FILENAME)

    logging.basicConfig(
        filename=logger_file_path,
        level=logging.INFO,
        format=Config.LOGGER_FORMAT
    )

    return logging.getLogger(__name__)