import os
import json
import uuid
from flask import Flask, request
from google.cloud import pubsub_v1
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import bot_logic # This imports your script!

app = Flask(__name__)

PROJECT_ID = os.getenv("GOOGLE_CLOUD_PROJECT")
TOPIC_ID = os.getenv("PUBSUB_TOPIC_ID", "calendar-updates")

# Recreate credentials.json from secret (just like your old GitHub Action did)
cred_content = os.getenv("GOOGLE_CREDENTIALS_JSON_CONTENT")
if cred_content and not os.path.exists("credentials.json"):
    with open("credentials.json", "w") as f:
        f.write(cred_content)

@app.route('/webhook', methods=['POST'])
def calendar_webhook():
    # Google Calendar requires this quick sync response
    if request.headers.get('X-Goog-Resource-State') == 'sync':
        return 'Sync successful', 200

    # Push a ticket to Pub/Sub to trigger the bot in the background
    if PROJECT_ID and TOPIC_ID:
        publisher = pubsub_v1.PublisherClient()
        topic_path = publisher.topic_path(PROJECT_ID, TOPIC_ID)
        publisher.publish(topic_path, b"Trigger Bot")
    
    return 'Webhook received, triggering bot', 200

@app.route('/process', methods=['POST'])
def process_meetings():
    try:
        print("Starting Real-Time Bot Logic...")
        bot_logic.main() # Runs your main script!
        return 'Processed Successfully', 200
    except Exception as e:
        print(f"Error: {e}")
        return f"Internal Error: {e}", 500

@app.route('/setup-webhook', methods=['GET'])
def setup_webhook():
    try:
        # Connect to Calendar using your existing token
        token_info = json.loads(os.getenv("GOOGLE_TOKEN_JSON_CALENDAR"))
        creds = Credentials.from_authorized_user_info(token_info)
        calendar_service = build('calendar', 'v3', credentials=creds)

        # Tell Google Calendar to send webhooks to this server
        webhook_url = f"{request.host_url.rstrip('/')}/webhook"
        request_body = {
            "id": str(uuid.uuid4()),
            "type": "web_hook",
            "address": webhook_url
        }
        response = calendar_service.events().watch(
            calendarId='primary', body=request_body
        ).execute()

        return f"Success! Connected to Calendar. Webhook URL: {webhook_url}", 200
    except Exception as e:
        return f"Failed to register webhook: {e}", 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
