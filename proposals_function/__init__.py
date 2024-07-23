import azure.functions as func
import logging
import json
import threading
from assistant import handle_request

def handle_request_in_background(attachments_folder_id, response_folder_id):
    try:
        logging.info(f"Background processing started for Attachments Folder ID: {attachments_folder_id}, Response Folder ID: {response_folder_id}")
        handle_request(attachments_folder_id, response_folder_id)
        logging.info(f"Background processing completed for Attachments Folder ID: {attachments_folder_id}, Response Folder ID: {response_folder_id}")
    except Exception as e:
        logging.error(f"Error processing request in background: {e}", exc_info=True)

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
        # Parse the request body to get folder IDs
        req_body = req.get_json()
        attachments_folder_id = req_body.get('attachments_folder_id')
        response_folder_id = req_body.get('response_folder_id')

        if not attachments_folder_id or not response_folder_id:
            logging.error("Missing folder IDs in the request body.")
            return func.HttpResponse(
                "Missing folder IDs in the request body.",
                status_code=400
            )

        # Start the handle_request function in a background thread
        threading.Thread(target=handle_request_in_background, args=(attachments_folder_id, response_folder_id)).start()

        # Immediately return a response to Power Automate
        return func.HttpResponse(
            "The function is processing in the background. Check logs for details.",
            status_code=202
        )
    except Exception as e:
        logging.error(f"Error processing request: {e}", exc_info=True)
        return func.HttpResponse(
            "An error occurred. Check logs for details.",
            status_code=500
        )
