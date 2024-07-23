import logging
from assistant import handle_request

# Set up logging
logging.basicConfig(level=logging.INFO)

# Define the folder IDs
attachments_folder_id = "01EDGXCHBGKCGUXZB7VFB2RILVACUQIZAU"
response_folder_id = "01EDGXCHD3NDRYZ5AQOJAYFK4MIPAJGNJS"

# Call the function
try:
    handle_request(attachments_folder_id, response_folder_id)
    print("Request processed successfully.")
except Exception as e:
    print(f"An error occurred: {e}")
