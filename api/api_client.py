import requests
from config.config import Config
from utils.logger import logger_setup

# CREATE A LOGGER OBJECT
logger = logger_setup()

def get_api_data():
    try:
        response = requests.get(Config.API_URL, timeout=Config.API_TIMEOUT)
        response.raise_for_status()
        return response.json()
        
    except requests.exceptions.ConnectionError as e:
        logger.error(f"Failed to connect to {Config.API_URL}. "
                     f"Network or server down. Details {e}.")
        return
    except requests.exceptions.Timeout as e:
        logger.error(f"API request to {Config.API_URL} failed. "
                     f"Timeout threshold of {Config.API_TIMEOUT}s reached. "
                     f"Details: {e}.")
        return
    except requests.exceptions.JSONDecodeError as e:
        logger.error(f"Failed to decode JSON from data source."
                     f"Error details: {e}.")
        return
    except requests.exceptions.RequestException as e:
        logger.error(e)
        return

if __name__ == '__main__':
    get_api_data()