import azure.functions as func
import azurefunctions.extensions.connectors.office365 as office365
import logging
import json
from typing import List

app = func.FunctionApp()


@app.function_name(name="OnNewEmail")
@app.connector_trigger(arg_name="emails")
def on_new_email(emails: List[office365.ClientReceiveMessage]) -> None:
    logging.info("Received Microsoft 365 OnNewEmail trigger")
    try:
        for email in emails:
            logging.info("Email received from: %s", email.from_)
            logging.info("Email subject: %s", email.subject)
        # Best-effort full-payload dump for parity with the .NET sample
        dump = [e.__dict__ if hasattr(e, "__dict__") else str(e) for e in emails]
        logging.info("Full payload: %s", json.dumps(dump, default=str, indent=2))
    except Exception:
        logging.exception("Failed to process email")
