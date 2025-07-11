import datetime
import os.path
import time
import base64
import traceback
from email.mime.text import MIMEText
import io # For GDrive downloads
import re
import markdown 
import json
import fitz
from google import genai
from google.genai import types
import enum

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from dotenv import load_dotenv
from pydantic import BaseModel, ValidationError

# For parsing Office documents if downloaded from Drive
from pptx import Presentation
import openpyxl

# Load the .env file
env = os.getenv("ENV", "dev")
env_file = f".env.{env}"

if not os.getenv("GITHUB_ACTIONS"):
    load_dotenv(env_file)

if not os.getenv("GITHUB_ACTIONS"):
    load_dotenv(env_file)
# --- Configuration ---
# For Google Workspace APIs (Calendar, Gmail, Drive)
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
]
CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json") # Downloaded from GCP
TOKEN_FILE_PREFIX = 'token_brandvmeet' # Will generate token_brandvmeet_calendar.json etc.

# For Gemini API
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") # Set this environment variable

# Google Drive Folder ID containing NBH data
NBH_GDRIVE_FOLDER_ID = os.getenv("NBH_GDRIVE_FOLDER_ID") # Set env var or replace placeholder

AGENT_EMAIL = "brand.vmeet@nobroker.in" # Email of the agent account
ADMIN_EMAIL_FOR_NOTIFICATIONS = "ajay.saini@nobroker.in" # REPLACE with your actual email
leadership_emails = ["sristi.agarwal@nobroker.in", "rohit.c@nobroker.in"] # Add the second email

EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP = {
    AGENT_EMAIL.lower().split('@')[0],
    "pia.brand","pia"
}

PROCESSED_EVENTS_FILE = 'processed_event_ids.txt' # Simple file-based tracking for local runs

# Cache file for inferred industries
# INFERRED_INDUSTRIES_CACHE_FILE = 'inferred_industries_cache.json' # Cache for brand industry inference



# Specific File Names (you might make these configurable or discover them)
FILE_NAME_PITCH_DECK_PDF = "NBH Monetization Pitch Deck.pdf" # From your image
# FILE_NAME_CASE_STUDIES_GSLIDES = "National Campaigns_case studies" # From your image
FILE_NAME_CASE_STUDIES_PDF = "National_Campaigns_case_studies.pdf" # New or alternative
FILE_NAME_PHYSICAL_CAMPAIGNS_GSHEET = "Physical_campaigns_live_sheet" # From your image
FILE_NAME_DIGITAL_CAMPAIGNS_GSHEET = "Digital_Campaigns_live_sheet" # From your image (note "live sheet")
FILE_NAME_COM_DATA_GSHEET = "NoBroker_Overall_Data" # From your image
FILE_NAME_NBH_PREVIOUS_MEETINGS_GSHEET = "NBH_Previous_meetings"


def parse_names_from_cell_helper(cell_value_str):
    names = set()
    if not cell_value_str or str(cell_value_str).strip().lower() == 'n/a':
        return names
    
    cell_value_str = str(cell_value_str) # Ensure it's a string
    # Remove content in parentheses (e.g., "(NoBrokerHood)", "(Brand Representative)")
    cleaned_cell = re.sub(r'\s*\([^)]*\)', '', cell_value_str)
    cleaned_cell = cleaned_cell.replace('*', '').strip() # Remove asterisks and leading/trailing spaces

    # Split by common delimiters. \s* around delimiters handles spaces.
    potential_names = re.split(r'\s*[,;/&\n]\s*|\s+\band\b\s+|\s+\bwith\b\s+', cleaned_cell)
    
    for name_part in potential_names:
        final_name = name_part.strip().lower()
        # Filter for meaningful names (e.g., length > 2) and exclude common role descriptors
        if final_name and len(final_name) > 2 and \
           "nbh sales" not in final_name and \
           "brand representative" not in final_name and \
           "nobrokerhood" not in final_name and \
           "stay vista" not in final_name: # Example: add other generic terms to exclude
            names.add(final_name)
    return names

# --- ADD THESE NEW HELPER FUNCTIONS NEAR THE TOP OF YOUR SCRIPT ---

def normalize_attendee_name(name_str):
    """
    Takes a raw name string and converts it into a standardized set of name parts.
    - "Shubham Chandrakant Dakhane" -> {'shubham', 'chandrakant', 'dakhane'}
    - "shubham.chandrakant" -> {'shubham', 'chandrakant'}
    - "trisha.bagchi7" -> {'trisha', 'bagchi'}
    - "mary" -> {'mary'}
    """
    if not isinstance(name_str, str) or not name_str.strip():
        return set()
    
    # Lowercase and replace common delimiters with spaces
    processed_name = name_str.lower().replace('.', ' ').replace('_', ' ')
    
    # Remove all characters that are not letters or spaces
    processed_name = re.sub(r'[^a-z\s]', '', processed_name)
    
    # Split into parts and filter out any empty strings resulting from multiple spaces
    name_parts = {part for part in processed_name.split() if part}
    
    return name_parts

def find_common_attendees(attendee_set_1_raw, attendee_set_2_raw):
    """
    Compares two sets of raw name strings and finds common individuals
    using a flexible, normalization-based approach.
    Returns a list of the matched raw names from the first set.
    """
    # Normalize all names in both sets
    # Each item in these lists will be a set of name parts, e.g., [{'shubham', 'chandrakant'}, {'trisha', 'bagchi'}]
    normalized_attendees_1 = [normalize_attendee_name(name) for name in attendee_set_1_raw]
    normalized_attendees_2 = [normalize_attendee_name(name) for name in attendee_set_2_raw]

    common_attendees_raw_names = []
    
    # Keep track of which attendees from set 2 have already been matched to avoid double counting
    matched_indices_in_set_2 = set()

    for i, norm_set_1 in enumerate(normalized_attendees_1):
        if not norm_set_1:  # Skip empty normalized names
            continue
        
        for j, norm_set_2 in enumerate(normalized_attendees_2):
            if j in matched_indices_in_set_2:
                continue # This person from set 2 is already matched
            if not norm_set_2:
                continue

            # THE CORE LOGIC: Check for equality or if one is a subset of the other.
            # This handles cases like {'shubham', 'chandrakant'} being a subset of {'shubham', 'chandrakant', 'dakhane'}
            if norm_set_1.issubset(norm_set_2) or norm_set_2.issubset(norm_set_1):
                # We have a match!
                # Add the original raw name from the first set to our list of common attendees
                common_attendees_raw_names.append(list(attendee_set_1_raw)[i])
                matched_indices_in_set_2.add(j)
                break # Move to the next person in set 1

    return common_attendees_raw_names

# --- END OF NEW HELPER FUNCTIONS ---

class Industry(enum.Enum):
    FMCG = "FMCG"
    AUTOMOTIVE_AND_TRANSPORT = "Automotive & Transportation"
    MEMBERSHIP_AND_LOCAL_SERVICES="Membership & Local Services"
    MARKETING_ADVERTISING_AND_MEDIA="Marketing, Advertising & Media"
    APPAREL_AND_FASHION="Apparel & Fashion"
    FOOD_AND_BEVERAGE="Food & Beverage"
    HEALTHCARE="Healthcare"
    FINANCE_AND_FINTECH="Finance & Fintech"
    BEAUTY_AND_PERSONAL_CARE="Beauty & Personal Care"
    JEWELLERY="Jewellery"
    REAL_ESTATE_AND_CONSTRUCTION="Real Estate & Construction"
    ENERGY_RENEWABLES_AND_MINING="Energy, Renewables & Mining"
    WELLNESS_AND_FITNESS="Wellness & Fitness"
    EDUCATION_AND_TRAINING="Education & Training"
    HOME_GOODS_AND_ELECTRONINCS="Home Goods & Electronics"
    HOSPITALITY_AND_TRAVEL="Hospitality & Travel"
    TECHNOLOGY_AND_BUSINESS_SERVICES="Technology & Business Services"
    E_COMMERCE="E-Commerce"
    RETAIL="Retail"
    PETS_AND_PETS_SERVICES="Pets & Pet Services"
    GAMING="Gaming"
    LOGISTICS_AND_WAREHOUSING="Logistics & Warehousing"
    OTHER_UNKNOWN="Other / Unknown"
    MANUFACTURING_AND_INDUSTRIAL="Manufacturing & Industrial"
    QUICk_COMMERCE="Quick Commerce"
    PHARMA = "Pharma"
    OTT = "OTT"


class Brand_Details(BaseModel):
    brand_name: str
    industry:   Industry

    class Config:
        use_enum_values = True  # Use enum values instead of names in JSON output

Allowed_Industries = [industry.value for industry in Industry]

BRAND_EXTRACTION_PROMPT_TEMPLATE = """
You are an expert administrative assistant working for NoBrokerHood (NBH) responsible for parsing meeting titles of meetings between NBH and different companies to extract key business information about those companies.
Your task is to analyze the provided meeting title and return a JSON object with two specific keys: "brand_name" and "industry".

Follow these rules precisely:
1.  **brand_name**: Identify the primary brand or company being met. If a title follows the pattern 'Parent Company (Brand)', the 'Brand' inside the parentheses is the primary `brand_name`. The parent company should be ignored for this task.
2.  **industry**: Infer the most likely industry for the primary `brand_name` strictly from **Allowed_Industries** mentioned below.
3.  If the title is ambiguous or you cannot identify a clear brand, return "Unknown" for both fields.
4.  Your response MUST be ONLY the JSON object, with no other text or markdown fences.

Allowed_Industries: {Allowed_Industries}
---
Here are some examples:

Title: "TCPL (Tetley) X NoBrokerHood/Partnership, 11am"
{{
  "brand_name": "Tetley",
  "industry": "Beverages"
}}

Title: "NBH X GIVA Digital _ June Discussion, 12pm"
{{
  "brand_name": "Giva",
  "industry": "Jewellery"
}}

Title: "Physical meeting - Posterscope X NBH, 3:15pm"
{{
  "brand_name": "Posterscope",
  "industry": "Media Agency"
}}

Title: "Campaign Discussion | Aashirvaad Svasti x NBH, 5pm"
{{
  "brand_name": "Aashirvaad Svasti",
  "industry": "Dairy"
}}

Title: "Internal Team Sync"
{{
  "brand_name": "Unknown",
  "industry": "Unknown"
}}
---

Now, analyze the following title:

Title: "{MEETING_TITLE}"
"""

def get_brand_details_from_title_with_llm(gemini_llm_client, meeting_title):
    """
    Uses a single LLM call to extract brand name and industry from a meeting title.
    Returns a dictionary with the extracted info or defaults if parsing fails.
    """
    default_response = {
        "brand_name": "Unknown Brand",
        "industry": "Unknown"
    }
    if not gemini_llm_client:
        print("  LLM model not available for brand extraction.")
        return default_response

    prompt = BRAND_EXTRACTION_PROMPT_TEMPLATE.format(MEETING_TITLE=meeting_title, Allowed_Industries=Allowed_Industries)

    # Define the grounding tool
    grounding_tool = types.Tool(
        google_search=types.GoogleSearch()
    )

    # Configure generation settings
    config = types.GenerateContentConfig(
        tools=[grounding_tool]
    )
    try:
        response = gemini_llm_client.models.generate_content(model="gemini-2.5-flash",contents=prompt, config=config)
        raw_text = response.candidates[0].content.parts[0].text

        cleaned_json_str = re.sub(r'```json\s*|\s*```', '', raw_text).strip()
        data = json.loads(cleaned_json_str)

        # Validate the simplified response structure
        if "brand_name" in data and "industry" in data:
            if not data["brand_name"] or data["brand_name"].lower() == 'unknown':
                print(f"  LLM identified title '{meeting_title}' as ambiguous.")
                return default_response
            return data
        else:
            print(f"  Error: LLM response for '{meeting_title}' was missing 'brand_name' or 'industry'.")
            return default_response

    except (json.JSONDecodeError, IndexError, AttributeError, Exception) as e:
        print(f"⚠️ First pass failed ({e}), retrying with strict JSON schema…")

        # Build a cleanup prompt + strict JSON enforcement (no tools)
        cleanup_prompt = (
            "The text below is supposed to be a JSON object matching this schema:\n\n"
            f"{Brand_Details.schema_json(indent=2)}\n\n"
            "But it wasn’t valid JSON. Please reformat exactly as JSON (no extra text):\n\n"
            f"{raw_text}"
        )

        cleanup_config = types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=Brand_Details
        )
        retry = gemini_llm_client.models.generate_content(
            model="gemini-2.5-flash",
            contents=cleanup_prompt,
            config=cleanup_config
        )

        # the SDK will give you a .parsed attribute when you supply response_schema
        try:
            parsed: Brand_Details = retry.parsed
            return parsed.model_dump()   # or parsed.dict()
        except Exception as e2:
            print(f"❌ Retry still failed: {e2}")
            return default_response

    except Exception as e:
        # anything else
        print(f"❌ Unexpected error: {e}")
        return default_response


# --- Google Authentication and Service Building ---
def get_google_service(service_name, version, scopes_list, token_filename_base_for_local_storage): # Changed last param name for clarity
    creds = None
    # Construct the specific local token filename (e.g., token_brandvmeet_calendar.json)
    # This is used for local development fallback and saving tokens locally.
    local_token_file_path = token_filename_base_for_local_storage

    # --- Attempt 1: Load from specific environment variable for the service (for CI) ---
    # Construct the expected environment variable name, e.g., GOOGLE_TOKEN_JSON_CALENDAR
    token_env_var_name = f"GOOGLE_TOKEN_JSON_{service_name.upper()}"
    token_json_string_from_env = os.getenv(token_env_var_name)

    if os.getenv('CI') == 'true' and token_json_string_from_env: # Check 'CI' env var and if token string exists
        print(f"CI environment detected. Attempting to load credentials for {service_name} from env var: {token_env_var_name}")
        try:
            token_info = json.loads(token_json_string_from_env)
            # The token_info from your stored JSON should contain client_id, client_secret, and refresh_token,
            # which are needed for the refresh mechanism by the Credentials object.
            creds = Credentials.from_authorized_user_info(token_info, scopes_list)
            print(f"Successfully loaded credentials for {service_name} from environment variable.")
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON from {token_env_var_name} for {service_name}: {e}")
            creds = None
        except Exception as e:
            print(f"Generic error loading credentials for {service_name} from environment variable {token_env_var_name}: {e}")
            creds = None # Fallback

    # --- Attempt 2: Load from local token file (for local development or if CI load failed) ---
    if not creds and os.path.exists(local_token_file_path):
        print(f"Loading credentials for {service_name} from local file: {local_token_file_path}")
        try:
            creds = Credentials.from_authorized_user_file(local_token_file_path, scopes_list)
        except Exception as e:
            print(f"Error loading credentials from {local_token_file_path} for {service_name}: {e}")
            creds = None

    # --- Attempt 3 & 4: Refresh or Run Interactive Flow ---
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print(f"Token for {service_name} is expired. Refreshing...")
            try:
                creds.refresh(Request())
                print(f"Token for {service_name} refreshed successfully.")
                # If refreshed successfully in a non-CI environment, save it back to the local file
                if not os.getenv('CI') == 'true':
                    try:
                        with open(local_token_file_path, 'w') as token_file:
                            token_file.write(creds.to_json())
                        print(f"Saved refreshed token for {service_name} to {local_token_file_path}")
                    except Exception as e_save:
                        print(f"Error saving refreshed token for {service_name} to {local_token_file_path}: {e_save}")
            except Exception as e_refresh:
                print(f"Error refreshing token for {service_name}: {e_refresh}")
                # If refresh fails in CI, it's a critical issue as there's no interactive fallback.
                if os.getenv('CI') == 'true':
                    print(f"FATAL: Token refresh failed for {service_name} in CI. Stored token might be invalid (e.g., revoked) or scopes changed.")
                    return None
                creds = None # Force re-authentication locally if refresh failed and not in CI
        
        # This block should ideally NOT be reached in a CI environment if the token env var is set correctly.
        if not creds:
            if os.getenv('CI') == 'true':
                print(f"FATAL: No valid credentials for {service_name} in CI and interactive flow is disabled.")
                print(f"       Ensure {token_env_var_name} secret is set correctly and contains valid token JSON.")
                return None # Critical failure in CI

            print(f"No valid credentials for {service_name}. Attempting interactive local server flow...")
            # Ensure CREDENTIALS_FILE ('credentials.json' with OAuth client_id/secret) exists for the flow
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"FATAL: {CREDENTIALS_FILE} not found. Cannot run interactive auth flow for {service_name}.")
                return None
            try:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, scopes_list)
                creds = flow.run_local_server(port=0) # This opens a browser for user consent
                # Save the newly obtained credentials to the local token file for future runs
                with open(local_token_file_path, 'w') as token_file:
                    token_file.write(creds.to_json())
                print(f"Saved new token for {service_name} (from interactive flow) to {local_token_file_path}")
            except Exception as e_flow:
                print(f"Error during interactive auth flow for {service_name}: {e_flow}")
                return None
    
    if not creds: # Should not happen if all paths above are handled
        print(f"Ultimately failed to obtain credentials for {service_name}.")
        return None

    # --- Build the Google API Service ---
    try:
        service = build(service_name, version, credentials=creds)
        print(f"{service_name.capitalize()} service initialized successfully.")
        return service
    except HttpError as error:
        print(f'An HTTP error occurred building {service_name} service: {error}')
        return None
    except Exception as e: # Catch other potential errors during build
        print(f'A general error occurred building {service_name} service: {e}')
        return None
    
# --- Google Drive Functions ---
def list_files_in_gdrive_folder(drive_service, folder_id):
    # ... (Implementation from previous thought block - list files) ...
    # Ensure it handles empty folder_id
    if not folder_id or folder_id == "YOUR_GDRIVE_FOLDER_ID_HERE":
        print("Google Drive Folder ID for NBH data is not configured.")
        return []
    try:
        results = drive_service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            # Include mimeType to decide how to read
            fields="nextPageToken, files(id, name, mimeType)"
        ).execute()
        items = results.get('files', [])
        return items
    except HttpError as error:
        print(f"An error occurred listing GDrive files for folder {folder_id}: {error}")
        return []

    

# THIS IS THE PRIMARY FUNCTION TO GET FILE DATA
def get_structured_gdrive_file_data(drive_service, sheets_service, file_id, file_name, mime_type):
    """
    Extracts and parses structured content from a Google Drive file based on its MIME type.
    
    Supports Google Slides (as PPTX), Microsoft PowerPoint, Google Sheets, Microsoft Excel, Google Docs, PDFs, and plain text files. Returns structured data as a list of dictionaries for presentations and spreadsheets, or as a string for text-based files. If parsing fails or the file type is unsupported, returns a descriptive error message.
    """

    print(f"    Attempting to read structured data for: {file_name} (MIME: {mime_type})")

    MIMETYPE_GOOGLE_SHEET = 'application/vnd.google-apps.spreadsheet'
    MIMETYPE_GOOGLE_DOC = 'application/vnd.google-apps.document'
    MIMETYPE_GOOGLE_PRESENTATION = 'application/vnd.google-apps.presentation'
    MIMETYPE_MS_EXCEL = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    MIMETYPE_MS_POWERPOINT = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    MIMETYPE_PDF = 'application/pdf'

    try:
        # --- Google Slides (Presentations) ---
        if mime_type == MIMETYPE_GOOGLE_PRESENTATION:
            print(f"    Exporting Google Slides '{file_name}' as PPTX for slide-level parsing...")
            # Export Google Slides as PPTX
            request_pptx = drive_service.files().export_media(
                fileId=file_id,
                mimeType=MIMETYPE_MS_POWERPOINT # Export as PPTX
            )
            fh_pptx = io.BytesIO()
            downloader_pptx = MediaIoBaseDownload(fh_pptx, request_pptx)
            done_pptx = False
            while not done_pptx:
                status, done_pptx = downloader_pptx.next_chunk()
            fh_pptx.seek(0)

            try:
                prs = Presentation(fh_pptx)
                slides_data = []
                for i, slide in enumerate(prs.slides):
                    slide_text_runs = []
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if run.text:
                                        slide_text_runs.append(run.text)
                    slides_data.append({
                        "slide_number": i + 1,
                        "text": "\n".join(slide_text_runs),
                        "source_type": "Google Presentation (exported as PPTX)"
                    })
                if not slides_data and prs.slides: # Had slides but no text extracted
                     return f"Google Slides '{file_name}': Exported as PPTX, but no text content found in slides."
                elif not prs.slides: # Exported PPTX was empty
                    return f"Google Slides '{file_name}': Exported as PPTX, but it contained no slides."
                return slides_data                          
            except Exception as e_pptx_parse:
                print(f"    Warning: Failed to parse Google Slides '{file_name}' exported as PPTX: {e_pptx_parse}. Falling back to plain text export.")
                # Fallback: Export as plain text if PPTX export/parse fails
                request_text = drive_service.files().export_media(fileId=file_id, mimeType='text/plain')
                fh_text = io.BytesIO()
                downloader_text = MediaIoBaseDownload(fh_text, request_text)
                done_text = False
                while not done_text:
                    status, done_text = downloader_text.next_chunk()
                fh_text.seek(0)
                full_text = fh_text.read().decode('utf-8', errors='replace')
                if not full_text.strip():
                    return f"Google Slides '{file_name}': Fallback to plain text export, but no content found."
                return [{"slide_number": 1, "text": full_text, "source_type": "Google Presentation (all text fallback)"}]


        # --- Microsoft PowerPoint (.pptx) ---
        elif mime_type == MIMETYPE_MS_POWERPOINT: # Handles native PPTX files
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: status, done = downloader.next_chunk()
            fh.seek(0)
            try:
                prs = Presentation(fh)
                slides_data = []
                for i, slide in enumerate(prs.slides):
                    slide_text_runs = []
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame") and shape.text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if run.text: slide_text_runs.append(run.text)
                    slides_data.append({"slide_number": i + 1, "text": "\n".join(slide_text_runs), "source_type": "PPTX (native)"}) # Clarified source
                if not slides_data and prs.slides:
                    return f"PPTX File '{file_name}': Contained slides, but no text content found in them."
                elif not prs.slides:
                    return f"PPTX File '{file_name}': Contained no slides."
                return slides_data
            except Exception as e:
                return f"Could not parse native PPTX content for '{file_name}': {e}"
            

        # --- Google Sheets & Microsoft Excel ---
        elif mime_type == MIMETYPE_GOOGLE_SHEET or mime_type == MIMETYPE_MS_EXCEL:
            all_rows_data = []
            try:
                if mime_type == MIMETYPE_GOOGLE_SHEET:
                    spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=file_id).execute()
                    for sheet_meta in spreadsheet.get('sheets', []):
                        sheet_title = sheet_meta['properties']['title']
                        # Read a significant number of rows, e.g., 1000, and all columns up to ZZ
                        range_str = f"'{sheet_title}'!A1:ZZ5000" 
                        result = sheets_service.spreadsheets().values().get(
                            spreadsheetId=file_id, range=range_str
                        ).execute()
                        rows = result.get('values', [])
                        if rows:
                            header = rows[0]
                            for row_idx, row_values in enumerate(rows): # (row_idx is 0-based here)
                                all_rows_data.append({"sheet_name": sheet_title, 
                                                      "row_index": row_idx + 1, # 1-based for display
                                                      "header": header, 
                                                      "values": row_values,
                                                      "source_type": "Google Sheet"})
                elif mime_type == MIMETYPE_MS_EXCEL:
                    request = drive_service.files().get_media(fileId=file_id)
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done: status, done = downloader.next_chunk()
                    fh.seek(0)
                    workbook = openpyxl.load_workbook(fh)
                    for sheet_title in workbook.sheetnames:
                        sheet = workbook[sheet_title]
                        header = [cell.value for cell in sheet[1]] # Assuming header is 1st row
                        for row_idx, row_obj in enumerate(sheet.iter_rows(min_row=1, max_col=50, max_row=1000, values_only=True)): # max_col to limit width
                            # row_idx is 0-based here from iter_rows(min_row=1)
                            all_rows_data.append({"sheet_name": sheet_title,
                                                  "row_index": row_idx + 1, # 1-based for display
                                                  "header": header, 
                                                  "values": list(row_obj), # Ensure it's a list
                                                  "source_type": "Excel Sheet"})
                if not all_rows_data: # If loops completed but no data (e.g. all sheets were empty)
                    return f"Spreadsheet file '{file_name}' ({mime_type}) processed, but no data rows found in any sheet."
                return all_rows_data
            except Exception as e:
                return f"Could not parse Spreadsheet content for '{file_name}' ({mime_type}): {e}"
        
        # --- Google Documents ---
        elif mime_type == MIMETYPE_GOOGLE_DOC:
            request = drive_service.files().export_media(fileId=file_id, mimeType='text/plain')
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: status, done = downloader.next_chunk()
            fh.seek(0)
            return fh.read().decode('utf-8', errors='replace') # Returns string

        # --- PDF Files ---
        elif mime_type == MIMETYPE_PDF:
            try:
                request = drive_service.files().get_media(fileId=file_id)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                pdf_text = ""
                with fitz.open(stream=fh.read(), filetype="pdf") as doc:
                    for page in doc:
                        pdf_text += page.get_text() + "\n"
                if not pdf_text.strip():
                    return f"PDF File: '{file_name}'. No extractable text found."
                return pdf_text
            except Exception as e:
                return f"Could not parse PDF content: {e}" 

        # --- Plain Text Files ---
        elif mime_type.startswith('text/'):
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: status, done = downloader.next_chunk()
            fh.seek(0)
            return fh.read().decode('utf-8', errors='replace') # Returns string
            
        else:
            return f"File type ({mime_type}) for '{file_name}' not configured for structured data extraction. Its name and existence might be relevant." # Returns string

    except HttpError as error:
        return f"An HTTP error occurred reading GDrive file {file_name} (ID: {file_id}): {error}" # Returns string
    except Exception as e:
        return f"A general error occurred reading GDrive file {file_name} (ID: {file_id}): {e}" # Returns string


def summarize_file_content_with_gemini(gemini_llm_client, file_name, mime_type, file_content):
    """
    Uses Gemini LLM to summarize the content of a file for inclusion in the meeting brief.
    """
    if not gemini_llm_client:
        return "Error: Gemini model not available for summarization."

    prompt = (
        f"Summarize the following content from the file '{file_name}' (type: {mime_type}) in 50-70 concise bullet points, "
        "focusing on key facts, data, or insights that would be useful for a marketing/sales meeting. "
        "Do not include generic statements. If the content is not relevant, say 'No relevant content found.'\n\n"
        f"---\n{file_content}\n---"
    )
    try:
        response = gemini_llm_client.models.generate_content(model="gemini-2.5-flash",contents=prompt)
        if response and response.candidates and response.candidates[0].content.parts:
            summary = response.candidates[0].content.parts[0].text.strip()
            return summary
        else:
            return "No summary generated."
    except Exception as e:
        print(f"Error during Gemini summarization: {e}")
        return f"Error: Exception during Gemini summarization: {e}"





# --- Modify get_internal_nbh_data_for_brand ---
def get_internal_nbh_data_for_brand(drive_service, sheets_service, gemini_llm_client, 
                                    current_target_brand_name,target_brand_industry,current_meeting_data,EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP, AGENT_EMAIL):
    # ... (initial checks for services and folder ID remain) ...
    """
    Fetches and processes internal NBH data relevant to a target brand for an upcoming meeting, aggregating summaries, campaign data, and previous meeting insights.
    
    This function collects and analyzes various internal files from Google Drive (pitch decks, case studies, campaign sheets, platform metrics, and previous meetings) to build a comprehensive context string for LLM-based brief generation. It extracts and cleans attendee names, identifies relevant campaign and case study data, and processes previous meeting records to determine if the upcoming meeting is a direct follow-up or if there are other past interactions with the brand. The function returns a structured summary for LLM input, flags for follow-up and past interactions, and a condensed summary of past meetings for leadership alerts.
    Returns:
        dict: {
            "llm_summary_string": Aggregated context string for LLM brief generation,
            "is_overall_direct_follow_up": True if the meeting is a direct follow-up based on attendee overlap,
            "has_other_past_interactions": True if there are other past meetings with the brand,
            "condensed_past_meetings_for_alert": List of summarized past meetings for leadership notification
        }
    """
    print(f"Fetching and processing internal NBH data for target brand '{current_target_brand_name}'...")
    all_files_in_folder = list_files_in_gdrive_folder(drive_service, NBH_GDRIVE_FOLDER_ID) # Ensure this is called
    
    if not all_files_in_folder: # Add check for empty list
        # Return structure consistent with success case but indicating no data
        return {
            "llm_summary_string": f"--- Targeted Internal NBH Data & Summaries for {current_target_brand_name} ---\n\nNo files found in the NBH GDrive folder to process.\n\n--- End of Targeted Internal NBH Data & Summaries ---\n",
            "is_overall_direct_follow_up": False,
            "has_previous_interactions": False,
            "condensed_past_meetings_for_alert": []
        }
    
    print(f"  Target Brand: '{current_target_brand_name}', Inferred Industry: '{target_brand_industry}'")

    # Get clean names from current meeting data
    #def get_clean_names_from_attendee_list(attendee_list_param):
    #    names = set()
    #    for att in attendee_list_param:
    #        name_to_process = att_dict.get('name', '') # Get the pre-extracted name
    #        if not name_to_process and 'email' in att_dict: # Fallback to email prefix if 'name' is empty
    #            name_to_process = att_dict['email'].split('@')[0]
    #        if not name_to_process: # Skip if still no name
    #            continue
            
             # Your existing cleaning logic from parse_names_from_cell_helper can be used here
             # Remove content in parentheses, asterisks, split by delimiters, etc.
    #        cleaned_name_to_process = re.sub(r'\s*\([^)]*\)', '', str(name_to_process))
    #        cleaned_name_to_process = cleaned_name_to_process.replace('*', '').strip().lower()
    
             # Split by common delimiters (if needed, but usually display names are clean)
             # For simplicity, let's assume display names are mostly single entities for now
             # If they can contain commas etc., you'd re-apply parts of parse_names_from_cell_helper's split logic
            
    #        if cleaned_name_to_process and len(cleaned_name_to_process) > 2: # Basic filter
    #           # Further filtering like in parse_names_from_cell_helper
    #            if "nbh sales" not in cleaned_name_to_process and \
    #                "brand representative" not in cleaned_name_to_process and \
    #                "nobrokerhood" not in cleaned_name_to_process:
    #                names.add(cleaned_name_to_process)
    #     return names

    current_nbh_attendees_list = current_meeting_data.get('nbh_attendees', []) # List of {'name': 'Display Name', 'email': '...'}
    current_brand_attendees_list = current_meeting_data.get('brand_attendees_info', []) # List of {'name': 'Display Name / Email', 'email': '...'}

    # For NBH attendees, prioritize the 'name' field (which should be displayName)
    nbh_names_to_parse = []
    for att in current_nbh_attendees_list:
        # Prefer 'name' (displayName), fallback to email prefix
        name_candidate = att.get('name', '')
        if not name_candidate or "@" in name_candidate: # If name is empty or looks like an email
            email_prefix = att.get('email', '').split('@')[0]
            if email_prefix:
                name_candidate = email_prefix.replace(".", " ") # Convert sreetoma.lodh to sreetoma lodh
        if name_candidate: # Only add if we have something
            name_candidate = name_candidate.replace(".", " ") 
            nbh_names_to_parse.append(name_candidate)
    
    # For Brand attendees, it's trickier as displayName might be just the email
    brand_names_to_parse = []
    for att in current_brand_attendees_list:
        name_candidate = att.get('name', '') # This is often the email itself for externals
        # If it looks like an email, try to get the part before @
        if "@" in name_candidate and name_candidate == att.get('email',''): # If name is identical to email
            email_prefix = name_candidate.split('@')[0]
            name_candidate = email_prefix.replace(".", " ").replace("_", " ") # spandeyh, akssuren
        elif not name_candidate and 'email' in att: # Fallback if name is empty
             name_candidate = att.get('email', '').split('@')[0].replace(".", " ").replace("_", " ")

        if name_candidate:
            name_candidate = name_candidate.replace(".", " ") 
            brand_names_to_parse.append(name_candidate)

    # Now use parse_names_from_cell_helper with these strings
    # We pass a single string of names joined by a common delimiter like a comma,
    # as parse_names_from_cell_helper expects a string.
    current_nbh_attendee_names = parse_names_from_cell_helper(", ".join(nbh_names_to_parse))
    current_brand_attendee_names = parse_names_from_cell_helper(", ".join(brand_names_to_parse))
                                        
    current_nbh_attendee_names_for_followup_check = {
    name for name in current_nbh_attendee_names if name not in EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP
    }
    #print(f"  DEBUG Current NBH Names for Follow-up Check: {current_nbh_attendee_names_for_followup_check}")
    #print(f"  DEBUG Current Brand Names for Follow-up Check: {current_brand_attendee_names}")

    final_context_parts_for_llm = [] # Accumulates strings for the LLM from ALL files
    
    # This list will store parsed data from the NBH_Previous_meetings sheet
    matching_previous_meetings_details_accumulator = [] 

    # --- Column indices for the NBH_Previous_meetings sheet ---
    # Initialize to -1, they will be set when the sheet is processed
    prev_meetings_brand_col_idx = -1
    prev_meetings_date_col_idx = -1
    prev_meetings_client_participants_col_idx = -1
    prev_meetings_nbh_participants_col_idx = -1
    prev_meetings_key_discussion_col_idx = -1
    prev_meetings_action_items_col_idx = -1
    prev_meetings_key_questions_col_idx = -1
    prev_meetings_brand_traits_col_idx = -1
    prev_meetings_customer_needs_col_idx = -1
    prev_meetings_client_pain_points_col_idx = -1
    # --- Add any other indices you need from this sheet ---

    # --- Process each specific, known file ---
    for item in all_files_in_folder: # Main loop starts here
        file_name = item.get('name', 'Unknown File')
        file_id = item['id']
        mime_type = item.get('mimeType', '')
        print(f"  Considering file: {file_name} ({mime_type})")

        # Call the consolidated and correctly named function
        file_data_object = get_structured_gdrive_file_data(
            drive_service, sheets_service, file_id, file_name, mime_type
        )

        # In get_internal_nbh_data_for_brand, after calling get_structured_gdrive_file_data:
        print(f"    DEBUG: file_data_object for {file_name} (Type: {type(file_data_object)}):")
        if isinstance(file_data_object, str):
            print(f"      Content/Error: {file_data_object[:500]}{'...' if len(file_data_object) > 500 else ''}")
        elif isinstance(file_data_object, list) and file_data_object:
            print(f"      Content: List of {len(file_data_object)} items. First item: {str(file_data_object[0])[:200]}...")
        elif isinstance(file_data_object, list) and not file_data_object:
            print(f"      Content: Empty list.")
        # Then proceed with your if/elif for file types

        # 1. NBH Monetization Pitch Deck.pdf (General Pitch Context)
        if FILE_NAME_PITCH_DECK_PDF.lower() in file_name.lower() and mime_type == 'application/pdf':
            if isinstance(file_data_object, str): # PDF content or error string
                is_error_or_no_content = (
                    (file_data_object.startswith("PDF File:") and "No extractable text found." in file_data_object) or
                    file_data_object.startswith("Could not parse PDF content:") or
                    file_data_object.startswith("An HTTP error occurred") or # General errors
                    file_data_object.startswith("A general error occurred") or
                    not file_data_object.strip() # Empty content
                )

                if not is_error_or_no_content and gemini_llm_client:
                     summary_of_pdf_content = summarize_file_content_with_gemini(gemini_llm_client, file_name, mime_type, file_data_object[:20000])
                     final_context_parts_for_llm.append(f"## General NBH Pitch Overview (from '{file_name}'):\n{summary_of_pdf_content}\n")
                elif not is_error_or_no_content: # Has text, but no LLM
                     final_context_parts_for_llm.append(f"## General NBH Pitch Overview (from '{file_name}'):\n{file_data_object[:1000]}...\n(Full content available, LLM summarization skipped)\n")
                else: # It's an error string or "no content" message
                     final_context_parts_for_llm.append(f"## General NBH Pitch Context (from '{file_name}'):\n{file_data_object}\n")
            continue

        # 2. National Campaigns_case studies (Now as PDF)
        elif FILE_NAME_CASE_STUDIES_PDF.lower() in file_name.lower() and mime_type == 'application/pdf':
            print(f"    Processing Case Studies PDF: {file_name}")
            if isinstance(file_data_object, str) and file_data_object.strip(): # file_data_object is the extracted PDF text
                # Check if it's an error message from PDF parsing
                is_error_or_no_content = (
                    (file_data_object.startswith("PDF File:") and "No extractable text found." in file_data_object) or
                    file_data_object.startswith("Could not parse PDF content:") or
                    file_data_object.startswith("An HTTP error occurred") or
                    file_data_object.startswith("A general error occurred")
                )
                if is_error_or_no_content:
                    final_context_parts_for_llm.append(f"## Case Studies (from PDF: '{file_name}'):\n{file_data_object}\n") # Report error                
                else:
                    # Successfully got text from the PDF.
                    # The PDF text will likely be one large block.
                    # Your LLM prompt needs to be good at segmenting this.
                    # We can't easily identify "slides" from a PDF's raw text dump.
                    # So, we provide the whole text and instruct the LLM.
                    # You might want to truncate if the PDF text is excessively long. MAX_PDF_TEXT_LEN = 50000      
                     
                    truncated_pdf_text = file_data_object
                    if len(truncated_pdf_text) > 50000: # Example truncation
                        truncated_pdf_text = file_data_object[:50000] + "\n... [Case Study PDF content truncated]"
                        print(f"    Truncated case study PDF text for '{file_name}' to 50000 chars.")                          

                    case_study_context = (
                        f"## Case Studies (from PDF document: '{file_name}'):\n"
                        f"The following is the extracted text content from the PDF. "
                        f"Please analyze this text to identify distinct case studies. For each, try to find "
                        f"the Brand, Objective, Activities, Key Results, and explain its relevance to '{current_target_brand_name}' "
                        f"(Industry: '{target_brand_industry}'). Pay attention to headings, lists, and common case study structures.\n\n"
                        f"---\nBEGIN PDF CASE STUDY TEXT\n---\n"
                        f"{truncated_pdf_text}\n"
                        f"---\nEND PDF CASE STUDY TEXT\n---\n"
                    )
                    final_context_parts_for_llm.append(case_study_context)
            else: # Should not happen if get_structured_gdrive_file_data returns string or error
                final_context_parts_for_llm.append(f"## Case Studies (from PDF: '{file_name}'):\nNo text data extracted or unexpected format.\n")
            continue # Move to next file

        # 3. & 4. Physical & Digital Campaigns (GSheets - Extract relevant rows)
        # THIS IS THE CORRECTED INDENTATION LEVEL FOR THE ELIF
        elif (FILE_NAME_PHYSICAL_CAMPAIGNS_GSHEET.lower() in file_name.lower() or \
              FILE_NAME_DIGITAL_CAMPAIGNS_GSHEET.lower() in file_name.lower()) and \
             mime_type == 'application/vnd.google-apps.spreadsheet':
            
            #rows_data_from_sheet = read_gdrive_file_content_with_slides(drive_service, sheets_service, file_id, file_name, mime_type)
            
            if isinstance(file_data_object, list) and file_data_object:
                # ... (the rest of your detailed logic for processing campaign GSheet rows) ...
                # ... (finding brand name column, iterating rows, checking direct match, checking industry match, etc.) ...
                relevant_rows_for_llm = [f"## Historical Campaign Data (from '{file_name}'):\n"]
                rows_added_count = 0
                MAX_CAMPAIGN_ROWS_TO_INCLUDE = 15 

                # Assuming first item in file_data_object (if it's a list of rows) contains the header
                first_row_dict = file_data_object[0] if isinstance(file_data_object[0], dict) else {}
                header_values = first_row_dict.get("header", [])
                
                brand_name_col_idx = -1
                industry_col_idx = -1 # NEW: For reading "Industry" column from sheet                
                

                if header_values:
                    try:
                        lower_header = [str(h).strip().lower() for h in header_values]
                        brand_name_col_idx = lower_header.index("brand name")
                        industry_col_idx = lower_header.index("industry") # Ensure this column exists in your sheets!
                    except ValueError:
                        missing_cols = []
                        if "brand name" not in lower_header: missing_cols.append("'Brand Name'")
                        if "industry" not in lower_header: missing_cols.append("'Industry'")
                        print(f"    Warning: Required column(s) {', '.join(missing_cols)} not found in header of {file_name}. Skipping detailed campaign processing for this file.")
                        final_context_parts_for_llm.append(f"## Historical Campaign Data (from '{file_name}'):\nRequired columns missing in sheet header. Cannot process for relevant campaigns.\n")
                        continue # Skip to next file if essential columns are missing

                for row_info in file_data_object:
                    if not isinstance(row_info, dict) or 'values' not in row_info or not row_info['values']: continue
                    if row_info.get("values") == header_values : # Skip header row
                        continue

                    row_values = row_info['values']
                    is_relevant_row = False
                    relevance_reason = ""
                    row_brand_name = ""
                    row_campaign_industry_from_sheet = "" # NEW


                    # Attempt to extract Brand Name from the current row    
                    if brand_name_col_idx != -1 and len(row_values) > brand_name_col_idx and row_values[brand_name_col_idx] is not None:
                        row_brand_name = str(row_values[brand_name_col_idx]).strip()
                        if current_target_brand_name.lower() == row_brand_name.lower():
                            is_relevant_row = True
                            relevance_reason = f"Direct match for '{current_target_brand_name}'."
                    
                    # Attempt to extract Industry from the current row
                    if industry_col_idx != -1 and len(row_values) > industry_col_idx and row_values[industry_col_idx] is not None:
                        row_campaign_industry_from_sheet = str(row_values[industry_col_idx]).strip()

                    # Relevance check based on industry
                    if not is_relevant_row and target_brand_industry != "Unknown" and row_campaign_industry_from_sheet: # Check if sheet industry is not empty    
                       if row_campaign_industry_from_sheet.lower() == target_brand_industry.lower():
                               is_relevant_row = True
                               relevance_reason = f"Industry match for '{target_brand_industry}' in row industry '{row_campaign_industry_from_sheet}'."


                    # ROW DEBUG print statements would go here, after all checks
                    #print(f"    ROW DEBUG: TargetBrand='{current_target_brand_name.lower()}', RowBrand='{row_brand_name.lower()}'")
                    #print(f"    ROW DEBUG: TargetIndustry='{target_brand_industry.lower()}', RowSheetIndustry='{row_campaign_industry_from_sheet.lower()}'")
                    #print(f"    ROW DEBUG: is_relevant_row after all checks: {is_relevant_row}, Reason: '{relevance_reason}'")


                    if is_relevant_row and rows_added_count < MAX_CAMPAIGN_ROWS_TO_INCLUDE:
                        print(f"      MATCH FOUND for row {row_info.get('row_index')}: {relevance_reason}")
                        # ... your existing logic to add to relevant_rows_for_llm ...
                        row_details = []
                        for h_idx, h_val in enumerate(header_values): # Use header_values from the top of this file's processing block
                            cell_val = str(row_values[h_idx]) if len(row_values) > h_idx else "N/A"
                            if cell_val.strip() and cell_val != "N/A": # Only add if there's a value
                                row_details.append(f"{h_val}: {cell_val.strip()}")
                        if row_details: # Only add if we have some details  
                            relevant_rows_for_llm.append(f"### Campaign Row {row_info.get('row_index', 'N/A')} (Relevance: {relevance_reason}):\n" + ", ".join(row_details) + "\n---\n")
                        rows_added_count +=1
                
                if rows_added_count > 0:
                    final_context_parts_for_llm.append("".join(relevant_rows_for_llm))
                else:
                    final_context_parts_for_llm.append(f"## Historical Campaign Data (from '{file_name}'):\nNo campaigns found directly matching '{current_target_brand_name}' or its inferred industry '{target_brand_industry}'.\n")

            elif isinstance(file_data_object, str): 
                final_context_parts_for_llm.append(f"## Note on '{file_name}':\nError processing campaign data: {file_data_object}\n")
            else: 
                 final_context_parts_for_llm.append(f"## Historical Campaign Data (from '{file_name}'):\nNo data rows found or file not in expected sheet format.\n")
            continue

        # 5. com data _ Cross vertical Services_ West (GSheet - Database Numbers/Leads - MULTI-TAB)
        elif FILE_NAME_COM_DATA_GSHEET.lower() in file_name.lower() and mime_type == 'application/vnd.google-apps.spreadsheet':
            com_data_structured_output = [
                f"## NoBroker.com Platform Metrics (from GSheet: '{file_name}'):\n"
                f"This Google Sheet contains multiple tabs, each with specific NoBroker platform metrics (e.g., lead volumes, user data, city-wise breakdowns, service-specific numbers). The agent preparing this brief should analyze the data from the relevant tabs described below to identify compelling metrics for '{current_target_brand_name}' (Industry: '{target_brand_industry}').\n\n"
            ]
            
            # `read_gdrive_file_content_with_slides` already returns a list of dicts, 
            # where each dict can represent a row from a specific sheet (tab) within the GSheet.
            # Each dict has: {"sheet_name": S, "row_index": R, "header": H, "values": V}
            if isinstance(file_data_object, list) and file_data_object and "sheet_name" in file_data_object[0]:

            #all_rows_from_com_data_sheet = read_gdrive_file_content_with_slides(
                drive_service, sheets_service, file_id, file_name, mime_type
            #)

            if isinstance(file_data_object, list) and file_data_object:
                processed_tabs = {} # Key: tab_name, Value: list of row strings

                for row_info in file_data_object:
                    if 'sheet_name' not in row_info or 'values' not in row_info or not row_info['values']:
                        continue # Skip malformed row_info

                    tab_name = row_info['sheet_name']
                    
                    if tab_name not in processed_tabs:
                        processed_tabs[tab_name] = []
                        # Add header to the tab's data if available and it's the first time we see this tab
                        if "header" in row_info and row_info["header"]:
                             processed_tabs[tab_name].append(f"  Columns: {', '.join(map(str, row_info['header']))}\n")
                    
                    # Limit the number of sample rows per tab
                    MAX_SAMPLE_ROWS_PER_TAB = 7 # Show up to 7 data rows per tab (plus header)
                    if len(processed_tabs[tab_name]) < (MAX_SAMPLE_ROWS_PER_TAB + 1): # +1 for potential header
                        # Only add the row if it's not identical to the header (simple check)
                        if not (row_info.get("row_index", 0) == 1 and row_info.get("values") == row_info.get("header")):
                            processed_tabs[tab_name].append(f"  Data: {', '.join(map(str, row_info['values']))}\n")
                
                # Now assemble the output for the LLM
                for tab_name, rows_as_strings in processed_tabs.items():
                    if rows_as_strings: # Only if we have content for this tab
                        com_data_structured_output.append(f"### Data from Tab: '{tab_name}'\n")
                        com_data_structured_output.extend(rows_as_strings)
                        if len(rows_as_strings) > MAX_SAMPLE_ROWS_PER_TAB : # Check if we hit the row limit
                             com_data_structured_output.append("  ... (more rows available in this tab)\n")
                        com_data_structured_output.append("\n") # Separator between tabs

            elif isinstance(file_data_object, str): # Error reading sheet
                com_data_structured_output.append(f"Error processing the '{file_name}' GSheet: {file_data_object}\n") # CORRECTED: use file_data_object
            else: # No data or unexpected format
                com_data_structured_output.append(f"No data or tabs could be extracted from '{file_name}', or it was not in the expected sheet format.\n")
            
            final_context_parts_for_llm.append("".join(com_data_structured_output))
            continue 
        
        # 6. NBH Previous Meetings Data
        elif FILE_NAME_NBH_PREVIOUS_MEETINGS_GSHEET.lower() in file_name.lower() and \
             mime_type == 'application/vnd.google-apps.spreadsheet':
            print(f"    Processing NBH Previous Meetings Sheet: {file_name}")
            previous_meeting_notes_list = [] # Use a list to build strings

            if isinstance(file_data_object, list) and file_data_object:              
                # --- Define expected header names for this sheet (adjust as per your sheet) ---
                # It's crucial these match your sheet's column headers exactly (case-insensitive)
                # Based on your sample output: 'Meeitng ID', 'Meeting Date', 'Brand Name', 'Meeting Title',
                # 'Industry Type', 'Brand Size', 'Meeting Type', 'Meeting Agenda', 'NBH Participants', 'Client Participants',
                # 'Key Discussion Points', 'Action Items/Next Steps', 'Is Follow Up?'
                # Let's assume a few key ones for matching:
                HEADER_PREV_BRAND_NAME = "brand name"
                HEADER_PREV_MEETING_DATE = "meeting date" # For sorting
                HEADER_PREV_CLIENT_PARTICIPANTS_NAMES = "client participants" # Column with client names
                HEADER_PREV_NBH_PARTICIPANTS_NAMES = "nbh participants"       # Column with NBH names
                HEADER_PREV_KEY_DISCUSSION = "key discussion points"
                HEADER_PREV_ACTION_ITEMS = "action items"
                #HEADER_PREV_IS_FOLLOW_UP_TAG = "is follow up?" # A column you might add to explicitly tag follow-ups
                HEADER_PREV_KEY_QUESTIONS = "key questions" # Adjust to your exact column name
                HEADER_PREV_BRAND_TRAITS = "brand traits"   # Adjust
                HEADER_PREV_CUSTOMER_NEEDS = "customer needs" # Adjust
                HEADER_PREV_CLIENT_PAIN_POINTS = "client pain points" # Adjust

                first_row_dict = file_data_object[0] if isinstance(file_data_object[0], dict) else {}
                header_values = first_row_dict.get("header", [])
                lower_header = [str(h).strip().lower() for h in header_values]

                # Add more column indices as needed based on what you want to extract
                print(f"    DEBUG NBH_Previous_meetings - Parsed Headers (lower_header): {lower_header}") # <--- ADD THIS

                try:
                            #print(f"Attempting to find: '{HEADER_PREV_BRAND_NAME}'")
                            prev_meetings_brand_col_idx = lower_header.index(HEADER_PREV_BRAND_NAME)
                            #print(f"Found '{HEADER_PREV_BRAND_NAME}' at index {prev_brand_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_MEETING_DATE}'")
                            prev_meetings_date_col_idx = lower_header.index(HEADER_PREV_MEETING_DATE)
                            #print(f"Found '{HEADER_PREV_MEETING_DATE}' at index {prev_date_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_CLIENT_PARTICIPANTS_NAMES}'")
                            prev_meetings_client_participants_col_idx = lower_header.index(HEADER_PREV_CLIENT_PARTICIPANTS_NAMES)
                            #print(f"Found '{HEADER_PREV_CLIENT_PARTICIPANTS_NAMES}' at index {client_participants_names_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_KEY_DISCUSSION}'")
                            prev_meetings_key_discussion_col_idx = lower_header.index(HEADER_PREV_KEY_DISCUSSION)
                            #print(f"Found '{HEADER_PREV_KEY_DISCUSSION}' at index {key_discussion_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_ACTION_ITEMS}'")
                            prev_meetings_action_items_col_idx = lower_header.index(HEADER_PREV_ACTION_ITEMS)
                            #print(f"Found '{HEADER_PREV_ACTION_ITEMS}' at index {action_items_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_KEY_QUESTIONS}'")
                            prev_meetings_key_questions_col_idx = lower_header.index(HEADER_PREV_KEY_QUESTIONS)
                            #print(f"Found '{HEADER_PREV_KEY_QUESTIONS}' at index {key_questions_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_BRAND_TRAITS}'")
                            prev_meetings_brand_traits_col_idx = lower_header.index(HEADER_PREV_BRAND_TRAITS)
                            #print(f"Found '{HEADER_PREV_BRAND_TRAITS}' at index {brand_traits_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_CUSTOMER_NEEDS}'")
                            prev_meetings_customer_needs_col_idx = lower_header.index(HEADER_PREV_CUSTOMER_NEEDS)
                            #print(f"Found '{HEADER_PREV_CUSTOMER_NEEDS}' at index {customer_needs_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_CLIENT_PAIN_POINTS}'")
                            prev_meetings_client_pain_points_col_idx = lower_header.index(HEADER_PREV_CLIENT_PAIN_POINTS)
                            #print(f"Found '{HEADER_PREV_CLIENT_PAIN_POINTS}' at index {client_pain_points_col_idx}")

                            #print(f"Attempting to find: '{HEADER_PREV_NBH_PARTICIPANTS_NAMES}'")
                            prev_meetings_nbh_participants_col_idx = lower_header.index(HEADER_PREV_NBH_PARTICIPANTS_NAMES)
                            #print(f"Found '{HEADER_PREV_NBH_PARTICIPANTS_NAMES}' at index {nbh_participants_names_col_idx}")

                except ValueError as e_val: # Catch the ValueError and print its specific message
                    print(f"    VALUE ERROR while finding column index: {e_val}") # This will tell you *which string* was not found
                    print(f"    Warning: Essential columns for previous meeting analysis not found in '{file_name}'. Skipping.")
                    # ... your existing error handling in except block ...
                    previous_meeting_notes_list.append(f"Could not process previous meetings sheet due to missing essential columns (e.g., Brand Name, Meeting Date, Key Discussion Points).\n")
                    final_context_parts_for_llm.append("".join(previous_meeting_notes_list))
                    # Re-initialize all indices to -1 within the except block if one fails, so subsequent code doesn't use stale (but valid looking) indices from a previous file
                    prev_meetings_brand_col_idx = -1
                    prev_meetings_date_col_idx = -1
                    prev_meetings_client_participants_col_idx = -1
                    prev_meetings_nbh_participants_col_idx = -1
                    prev_meetings_key_discussion_col_idx = -1
                    prev_meetings_action_items_col_idx = -1
                    prev_meetings_key_questions_col_idx = -1
                    prev_meetings_brand_traits_col_idx = -1
                    prev_meetings_customer_needs_col_idx = -1
                    prev_meetings_client_pain_points_col_idx = -1
                    continue # Skip further processing for this file
            
                all_past_meetings_for_brand = []

                for row_info in file_data_object:
                    if not isinstance(row_info, dict) or 'values' not in row_info or not row_info['values']:
                        continue
                    if row_info.get("values") == header_values:  # Skip header row
                        continue
                    
                    row_values = row_info['values']# Initialize for current row
                    
                    # Extract brand name from the previous meeting row
                    prev_meeting_brand_name_from_sheet = ""

                    if prev_meetings_brand_col_idx != -1 and len(row_values) > prev_meetings_brand_col_idx:
                        # Use .strip() to remove whitespace AND non-breaking spaces
                        prev_meeting_brand_name_from_sheet = str(row_values[prev_meetings_brand_col_idx] or "").replace('\xa0', ' ').strip()

                    if not prev_meeting_brand_name_from_sheet:
                        continue


                    # --- START OF TARGETED DEBUG BLOCK ---
                    # Only print if the sheet brand name (lowercase) potentially contains the target brand name (lowercase)
                    # or vice-versa. This helps narrow down the debug output significantly.
                    # This condition is intentionally broad to catch near misses or partial matches for debugging.
                    target_brand_lower = current_target_brand_name.lower()
                    sheet_brand_lower = prev_meeting_brand_name_from_sheet.lower()

                    #if target_brand_lower in sheet_brand_lower or \
                    #   (sheet_brand_lower and target_brand_lower.startswith(sheet_brand_lower)) or \
                    #   (target_brand_lower and sheet_brand_lower.startswith(target_brand_lower)) or \
                    #   (target_brand_lower == "chitale" and "chitale" in sheet_brand_lower) or \
                    #   (target_brand_lower == "giva" and "giva" in sheet_brand_lower) : # Add more specific checks if needed

                        #print(f"    TARGETED_PREV_MTG_DEBUG (Row {row_info.get('row_index', 'N/A')} in sheet '{file_name}'):")
                        #print(f"        Current Target Brand (Original) : '{current_target_brand_name}'")
                        #print(f"        Sheet Brand Name (Original)   : '{prev_meeting_brand_name_from_sheet}'")
                        #print(f"        --- For Comparison ---")
                        #print(f"        Current Target Brand (lower)  : '{target_brand_lower}' (len {len(target_brand_lower)})")
                        #print(f"        Sheet Brand Name (lower)    : '{sheet_brand_lower}' (len {len(sheet_brand_lower)})")
                        # Optional: Byte representation for very tricky cases
                        # print(f"        Bytes Current Target (lower): {target_brand_lower.encode('utf-8', 'surrogateescape')}")
                        # print(f"        Bytes Sheet Brand (lower)   : {sheet_brand_lower.encode('utf-8', 'surrogateescape')}")
                        #is_match = (sheet_brand_lower == target_brand_lower)
                        #print(f"        Exact lowercase match?        : {is_match}")
                        #print(f"        ------------------------------------")
                    # --- END OF TARGETED DEBUG BLOCK ---

                    is_a_match = False


                    # Minimum length for a brand name to be considered for matching to avoid trivial matches (e.g., 'a', 'an')
                    MIN_BRAND_NAME_LEN = 3 
                    
                    if len(target_brand_lower) < MIN_BRAND_NAME_LEN or len(sheet_brand_lower) < MIN_BRAND_NAME_LEN:
                        # If either name is too short, only an exact match is allowed.
                        if target_brand_lower == sheet_brand_lower:
                            is_a_match = True
                    else:
                        # Rule 1: Always prioritize an exact match. This is the strongest signal.
                        if target_brand_lower == sheet_brand_lower:
                            is_a_match = True
                        else:
                            # Rule 2: Use regex with word boundaries (\b) to check if one name is a
                            # "whole word" component of the other. This prevents matching "advan" inside "advances".
                            
                            # Check if the (potentially shorter) target brand is a whole word in the sheet brand.
                            # e.g., target="Chitale", sheet="Chitale Bandhu" -> MATCH
                            regex_target_in_sheet = r'\b' + re.escape(target_brand_lower) + r'\b'
                            if re.search(regex_target_in_sheet, sheet_brand_lower):
                                is_a_match = True
                            
                            # Check if the (potentially shorter) sheet brand is a whole word in the target brand.
                            # e.g., target="Meeting with Chitale Bandhu", sheet="Chitale Bandhu" -> MATCH
                            else:
                                regex_sheet_in_target = r'\b' + re.escape(sheet_brand_lower) + r'\b'
                                if re.search(regex_sheet_in_target, target_brand_lower):
                                    is_a_match = True
                    

                    if is_a_match:
                        # This previous meeting was with the same brand
                        print(f"    *** VERIFIED MATCH: Target='{target_brand_lower}' <==> Sheet='{sheet_brand_lower}'")
                        print(f"      MATCH FOUND for previous meeting row {row_info.get('row_index')}: Brand match for '{current_target_brand_name}'")
                        meeting_details_dict = {
                            "original_row_index": row_info.get("row_index", "N/A"), # Get it here
                            "original_row_info": row_info
                        }    
                        try:
                            # Extract date for sorting (handle potential date parsing errors)
                            date_str = str(row_values[prev_meetings_date_col_idx]) if len(row_values) > prev_meetings_date_col_idx else None
                            if date_str:
                                # Attempt to parse common date formats, be robust
                                for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%y"): # Add more if needed
                                    try:
                                        meeting_details_dict["date_obj"] = datetime.datetime.strptime(date_str.split(" ")[0], fmt).date() # Take only date part
                                        break
                                    except ValueError:
                                        pass
                                if "date_obj" not in meeting_details_dict:
                                     meeting_details_dict["date_obj"] = datetime.date.min # Fallback for unparseable dates
                            else:
                                meeting_details_dict["date_obj"] = datetime.date.min
                        except Exception:
                            meeting_details_dict["date_obj"] = datetime.date.min # Fallback
                        all_past_meetings_for_brand.append(meeting_details_dict)

                if all_past_meetings_for_brand:
                    matching_previous_meetings_details_accumulator.extend(all_past_meetings_for_brand)               

            elif isinstance(file_data_object, str): # Error reading sheet
                final_context_parts_for_llm.append(f"## Previous NBH Meetings (from '{file_name}'):\nError processing the sheet: {file_data_object}\n")
            else: # No data
                final_context_parts_for_llm.append(f"## Previous NBH Meetings (from '{file_name}'):\nNo data rows found in the sheet.\n")
            # This 'continue' ensures we don't fall through to generic file handling for this specific file type
            continue 

        # Fallback for other files or unrecognized structures
        elif isinstance(file_data_object, str) and file_data_object.strip():
            # Could be a simple text file, or an error message we haven't specifically handled
            # You might want to summarize these too if they are documents
            if "File type" in file_data_object or "Error" in file_data_object or "Could not parse" in file_data_object:
                 final_context_parts_for_llm.append(f"## Note on file '{file_name}':\n{file_data_object}\n")
            elif gemini_llm_client: # It's some text content
                 summary = summarize_file_content_with_gemini(gemini_llm_client, file_name, mime_type, file_data_object[:15000])
                 final_context_parts_for_llm.append(f"## Information from '{file_name}':\n{summary}\n")
            else: # No LLM, just pass truncated raw text
                 final_context_parts_for_llm.append(f"## Information from '{file_name}':\n{file_data_object[:500]}...\n")# --- Other files (not specifically handled) ---


    # --- Process Accumulated Previous Meeting Data (FINAL, COMPLETE VERSION) ---
    is_overall_direct_follow_up = False
    has_other_past_interactions = False
    condensed_past_meetings_for_alert = []
    previous_meeting_notes_for_llm_list = []

    # Helper to safely get cell values from the row
    def get_cell_val_from_row(row_values, col_idx, default="N/A"):
        """
        Retrieve and clean the value from a specified column index in a row, returning a default if the value is missing or empty.
        
        Parameters:
            col_idx (int): The index of the column to retrieve.
            default (str): The value to return if the cell is missing or empty.
        
        Returns:
            str: The stripped string value from the specified column, or the default if not present.
        """
        if col_idx != -1 and len(row_values) > col_idx and row_values[col_idx] is not None and str(row_values[col_idx]).strip():
            return str(row_values[col_idx]).strip()
        return default

    if matching_previous_meetings_details_accumulator:
        # Sort the entire collection of meetings by date, descending.
        matching_previous_meetings_details_accumulator.sort(key=lambda x: x.get("date_obj", datetime.date.min), reverse=True)

        # Set the number of recent meetings to analyze
        MAX_PREVIOUS_MEETINGS_TO_ANALYZE = 5
        latest_meetings_to_analyze = matching_previous_meetings_details_accumulator[:MAX_PREVIOUS_MEETINGS_TO_ANALYZE]




        # --- STEP 1: Populate full details and determine follow-up status for each recent meeting ---
        for mtg_data in latest_meetings_to_analyze:
            row_info = mtg_data['original_row_info']
            row_values = row_info['values']


            # THIS IS THE COMPLETE POPULATION OF ALL REQUIRED FIELDS
            mtg_data["discussion"] = get_cell_val_from_row(row_values, prev_meetings_key_discussion_col_idx)
            mtg_data["actions"] = get_cell_val_from_row(row_values, prev_meetings_action_items_col_idx)
            mtg_data["key_questions"] = get_cell_val_from_row(row_values, prev_meetings_key_questions_col_idx)
            mtg_data["brand_traits"] = get_cell_val_from_row(row_values, prev_meetings_brand_traits_col_idx)
            mtg_data["customer_needs"] = get_cell_val_from_row(row_values, prev_meetings_customer_needs_col_idx)
            mtg_data["client_pain_points"] = get_cell_val_from_row(row_values, prev_meetings_client_pain_points_col_idx)

            # Perform the name comparison to set the follow-up flag for this specific meeting
            mtg_data["is_direct_follow_up_candidate"] = False
            prev_nbh_names_str = get_cell_val_from_row(row_values, prev_meetings_nbh_participants_col_idx, "")
            prev_nbh_attendee_names_from_sheet_raw = parse_names_from_cell_helper(prev_nbh_names_str)
            prev_nbh_attendee_names_for_followup_check_sheet = { name for name in prev_nbh_attendee_names_from_sheet_raw if name not in EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP }
            
            common_nbh_attendees = find_common_attendees(
                current_nbh_attendee_names_for_followup_check,
                prev_nbh_attendee_names_for_followup_check_sheet
            )
            if common_nbh_attendees:
                mtg_data["is_direct_follow_up_candidate"] = True

        # --- STEP 2: Categorize the fully populated meetings ---
        direct_follow_up_meetings = [m for m in latest_meetings_to_analyze if m.get("is_direct_follow_up_candidate")]
        other_past_meetings = [m for m in latest_meetings_to_analyze if not m.get("is_direct_follow_up_candidate")]

        # Update the main flags based on the categorization
        is_overall_direct_follow_up = bool(direct_follow_up_meetings)
        has_other_past_interactions = bool(other_past_meetings)

        # --- STEP 3: Build the sophisticated context for the LLM brief ---
        previous_meeting_notes_for_llm_list.append(f"## Insights from Previous NBH Meetings with '{current_target_brand_name}':\n")

        if direct_follow_up_meetings:
            previous_meeting_notes_for_llm_list.append(
                "### Actionable Follow-Up Items (from meetings with attendee overlap):\n"
                "This upcoming meeting appears to be a DIRECT FOLLOW-UP to the specific meetings listed below. The brief should heavily focus on continuity, previous discussions, and action items from these engagements.\n\n"
            )
            for mtg_data_llm in direct_follow_up_meetings:
                date_llm = mtg_data_llm.get("date_obj", datetime.date.min).strftime("%Y-%m-%d")
                note_parts = [f"**Previous Meeting on {date_llm}:**\n"]
                if mtg_data_llm.get('discussion') != "N/A": note_parts.append(f"- Key Discussion: {mtg_data_llm['discussion']}\n")
                if mtg_data_llm.get('actions') != "N/A": note_parts.append(f"- Action Items: {mtg_data_llm['actions']}\n")
                if mtg_data_llm.get('key_questions') != "N/A": note_parts.append(f"- Key Client Questions: {mtg_data_llm['key_questions']}\n")
                if mtg_data_llm.get('client_pain_points') != "N/A": note_parts.append(f"- Client Pain Points Discussed: {mtg_data_llm['client_pain_points']}\n")
                note_parts.append("---\n")
                previous_meeting_notes_for_llm_list.append("".join(note_parts))

        if other_past_meetings:
            previous_meeting_notes_for_llm_list.append(
                "\n### General Context from Other Past Interactions:\n"
                "NBH has had other, separate discussions with this brand. The high-level context below is for awareness and should not be treated as direct action items for this upcoming meeting.\n\n"
            )
            for mtg_data_llm in other_past_meetings:
                date_llm = mtg_data_llm.get("date_obj", datetime.date.min).strftime("%Y-%m-%d")
                note_parts = [f"**Past Interaction on {date_llm}:**\n"]
                if mtg_data_llm.get('discussion') != "N/A": note_parts.append(f"- General Topic: {mtg_data_llm['discussion'][:200]}...\n")
                if mtg_data_llm.get('brand_traits') != "N/A": note_parts.append(f"- Observed Brand Traits: {mtg_data_llm['brand_traits']}\n")
                if mtg_data_llm.get('customer_needs') != "N/A": note_parts.append(f"- Identified Customer Needs: {mtg_data_llm['customer_needs']}\n")
                note_parts.append("---\n")
                previous_meeting_notes_for_llm_list.append("".join(note_parts))

        # --- STEP 4: Build data for the leadership alert email ---
        # The alert should contain info about the "other" meetings that are NOT direct follow-ups.
        for mtg_data_alert in other_past_meetings:
            date_display = mtg_data_alert.get("date_obj").strftime("%Y-%m-%d")
            discussion_summary = (mtg_data_alert.get('discussion', 'N/A')[:150] + "...")
            raw_nbh_str_alert = mtg_data_alert['original_row_info']['values'][prev_meetings_nbh_participants_col_idx]
            parsed_nbh_names_alert = parse_names_from_cell_helper(str(raw_nbh_str_alert))
            human_nbh_names_alert = { name for name in parsed_nbh_names_alert if name not in EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP }
            nbh_team_str = ", ".join(human_nbh_names_alert) if human_nbh_names_alert else "N/A"
            
            condensed_past_meetings_for_alert.append({
                "date": date_display,
                "discussion_summary": discussion_summary,
                "nbh_team": nbh_team_str
            })

    else:
        # This block handles the case where no previous meetings were found at all.
        previous_meeting_notes_for_llm_list.append(f"## Previous NBH Meetings:\nNo previous meeting records found for '{current_target_brand_name}'.\n")

    # --- Finalize the return dictionary ---
    final_context_parts_for_llm.append("".join(previous_meeting_notes_for_llm_list))
    llm_summary_output_str = "\n\n".join(final_context_parts_for_llm)
    final_llm_summary_for_return = f"--- Targeted Internal NBH Data & Summaries for {current_target_brand_name} ---\n\n{llm_summary_output_str}\n--- End of Targeted Internal NBH Data & Summaries ---\n"

    return {
        "llm_summary_string": final_llm_summary_for_return,
        "is_overall_direct_follow_up": is_overall_direct_follow_up,
        "has_other_past_interactions": has_other_past_interactions,
        "condensed_past_meetings_for_alert": condensed_past_meetings_for_alert
    }

# --- Calendar Processing ---
def get_upcoming_meetings(calendar_service, calendar_id='primary', time_delta_hours=96): # Process meetings in next 3 days
    now_utc = datetime.datetime.utcnow()
    time_min_str = (now_utc - datetime.timedelta(hours=24)).isoformat() + 'Z'
    time_max_str = (now_utc + datetime.timedelta(hours=time_delta_hours)).isoformat() + 'Z'
    
    print(f'Getting events between {time_min_str} and {time_max_str}')
    try:
        events_result = calendar_service.events().list(
            calendarId=calendar_id, timeMin=time_min_str, timeMax=time_max_str,
            singleEvents=True, orderBy='startTime',
            # q='brand.vmeet@nobroker.in' # This might filter too early if brandvmeet is added as resource
        ).execute()
        events = events_result.get('items', [])
        return events
    except HttpError as error:
        print(f'An error occurred fetching events: {error}')
        return []

def load_processed_event_ids():
    # ... (same as before) ...
    if not os.path.exists(PROCESSED_EVENTS_FILE): return set()
    with open(PROCESSED_EVENTS_FILE, 'r') as f: return set(line.strip() for line in f)

def save_processed_event_id(event_id):
    # ... (same as before) ...
    with open(PROCESSED_EVENTS_FILE, 'a') as f: f.write(event_id + '\n')

EVENT_TAG_PROCESSED = "[NBH_BRIEF_AGENT_PROCESSED_V1]"

def is_event_already_tagged(event_description):
    return EVENT_TAG_PROCESSED in (event_description or "")

def tag_event_as_processed(calendar_service, event_id, calendar_id='primary'):
    if not calendar_service:
        print("  Calendar service not available to tag event.")
        return
    try:
        event = calendar_service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        description = event.get('description', '')
        if not is_event_already_tagged(description):
            new_description = f"{description}\n\n{EVENT_TAG_PROCESSED}"
            updated_event_body = {'description': new_description}
            calendar_service.events().patch(
                calendarId=calendar_id, eventId=event_id, body=updated_event_body
            ).execute()
            print(f"  Tagged event {event_id} as processed in calendar.")
    except HttpError as error:
        print(f"  An error occurred tagging event {event_id}: {error}")

EVENT_REMINDER_SET_TAG = "[NBH_1HR_REMINDER_SET]"

def set_one_hour_email_reminder(calendar_service, event_id, calendar_id='primary'):
    if not calendar_service:
        print("  Calendar service not available to set reminder.")
        return False
    try:
        event = calendar_service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        
        # Check if our specific reminder tag is already there
        description = event.get('description', '')
        if EVENT_REMINDER_SET_TAG in description:
            print(f"  1-hour email reminder already marked as set for event {event_id}.")
            return True

        reminders = event.get('reminders', {})
        overrides = reminders.get('overrides', [])
        
        # Check if a 60-minute email reminder already exists
        has_one_hour_email_reminder = any(
            r.get('method') == 'email' and r.get('minutes') == 60 for r in overrides
        )

        if has_one_hour_email_reminder:
            print(f"  Event {event_id} already has a 60-minute email reminder.")
        else:
            print(f"  Adding 60-minute email reminder to event {event_id}.")
            overrides.append({'method': 'email', 'minutes': 60})
            body_update = {'reminders': {'useDefault': False, 'overrides': overrides}}
            calendar_service.events().patch(
                calendarId=calendar_id, eventId=event_id, body=body_update
            ).execute()
            print(f"  Successfully added 60-minute email reminder for event {event_id}.")

        # Add our custom tag to the description to indicate we've processed this for reminders
        if EVENT_REMINDER_SET_TAG not in description:
            new_description = f"{description}\n{EVENT_REMINDER_SET_TAG}".strip()
            body_update_desc = {'description': new_description}
            calendar_service.events().patch(
                calendarId=calendar_id, eventId=event_id, body=body_update_desc
            ).execute()
            print(f"  Tagged event {event_id} with {EVENT_REMINDER_SET_TAG}.")
        
        return True

    except HttpError as error:
        print(f"  An error occurred setting reminder for event {event_id}: {error}")
        return False



# --- Meeting Info Extraction ---

NBH_SERVICE_ACCOUNTS_TO_EXCLUDE = { # Emails to exclude from the displayed NBH attendee list
    AGENT_EMAIL.lower(),
    "pia.brand@nobroker.in",
    "pia@nobroker.in"
}


def extract_meeting_info(event, agent_email_global, nbh_service_accounts_to_exclude_global):
    """
    Extracts basic, non-inferential information from a calendar event.
    The brand name itself is NOT processed here; it is extracted by the LLM later.
    """
    event_id = event['id']
    summary = event.get('summary', 'No Title')
    start_str = event['start'].get('dateTime', event['start'].get('date'))
    start_time_obj = datetime.datetime.fromisoformat(start_str.replace('Z', '+00:00'))
    location = event.get('location', 'N/A')
    description = event.get('description', '')
    attendees = event.get('attendees', [])

    nbh_attendees = []
    brand_attendees_info = []

    is_agent_invited = any(attendee.get('email', '').lower() == agent_email_global.lower() for attendee in attendees)
    if not is_agent_invited:
        print(f"  Skipping event '{summary}': {agent_email_global} is not an attendee.")
        return None

    for attendee in attendees:
        email = attendee.get('email', '').lower()
        name = attendee.get('displayName', email.split('@')[0] if '@' in email else email)
        if '@nobroker.in' in email:
            if email not in nbh_service_accounts_to_exclude_global:
                nbh_attendees.append({'email': email, 'name': name})
        elif email:
            brand_attendees_info.append({'name': name, 'email': email})
# Removing this condition so that physical meetings do not get skipped
    # if not brand_attendees_info:
    #     print(f"  Skipping event '{summary}': No external attendees.")
    #     return "NO_EXTERNAL_ATTENDEES"

    return {
        'id': event_id,
        'title': summary, # Return the raw title
        'start_time_obj': start_time_obj,
        'start_time_str': start_time_obj.strftime("%Y-%m-%d %I:%M %p %Z (%A)"),
        'location': location,
        'description': description,
        'nbh_attendees': nbh_attendees,
        'brand_attendees_info': brand_attendees_info,
        'is_event_description_present_for_tagging': bool(event.get('description'))
    }


# --- Gemini LLM Integration ---
def configure_gemini():
    """
    Configures and returns a Gemini LLM model instance using the provided API key.
    
    Returns:
        A configured Gemini GenerativeModel object if successful, or None if configuration fails or the API key is missing.
    """
    if not GEMINI_API_KEY:
        print("GEMINI_API_KEY environment variable not set. LLM will not function.")
        return None
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        # Using a specific model version. 1.5 Flash is faster and cheaper for many tasks.
        # For higher quality, consider 'gemini-1.5-pro-latest'.
        # model = genai.GenerativeModel('gemini-2.5-flash')
        print(f"Gemini model configured successfully.")
        return client
    except Exception as e:
        print(f"Error configuring Gemini API: {e}")
        return None

YOUR_DETAILED_PROMPT_TEMPLATE_GEMINI = """
# NBH is a SAAS platform that digitizes all types of workflows within gated communities such as visitor entry approvals, helpdesk tickets, maintenance payments, communication and announcements etc.. We have 3 platforms, a guard app (used by guards at the gate), a resident app (used by residents in gated communities) and a admin dashboard. We operate in tier-1 cities in India, which means we are engaged with a customer segment which is very valuable to many brands.
# We approach various companies who want to target that segment of people we have for the ads. This might be from a large national/multinational company or a small local city brand.
# You are a world-class marketing strategist at NoBrokerHood (NBH) crafting a Pre-Meeting Brief for an upcoming client meeting for the monetization team. This function has sales team members called brand managers (BMs) that pitch digital assets (on our resident app) and physical monetization assets (society common areas that we monetize) that we have. 
# The goal is to prepare a comprehensive, data-driven brief that will help the BMs present a compelling case to the client.    
Pre-Meeting Brief
# Brand Name: {BRAND_NAME_FOR_BODY}
# Meeting Date and time: {MEETING_DATETIME}
# Meeting Title/Subject: {MEETING_TITLE} 
# Brand Attendees:
{BRAND_ATTENDEES_NAMES}
# Attendee from NoBrokerHood (NBH):
{NBH_ATTENDEES_NAMES}

NBH Internal data summary below:
# BEGIN INTERNAL NBH DATA SUMMARY
{INTERNAL_NBH_DATA_SUMMARY}
# END INTERNAL NBH DATA SUMMARY

# IMPORTANT: For every section, use markdown headers (##), sub-headers (###), and bullet/numbered lists. Never use asterisks (*) for bullets—use markdown dashes (-) or numbers.
# The total full meeting brief should not cross 6-7 pages. Remember to strictly follow the page count or line count instructions.



# --- CRITICAL INSTRUCTION BASED ON PREVIOUS MEETING DATA ---
# The "Usage of Nobroker/nobrokerhood data" section as part of the "Internal NBH Data Summary" shared above will indicate if this meeting with {BRAND_NAME_FOR_BODY}
# appears to be a DIRECT FOLLOW-UP to recent NBH interactions, if NBH has had OTHER PREVIOUS INTERACTIONS
# with the brand, or if this is likely a FIRST-TIME INTERACTION.
# ALSO, pay attention to the {MEETING_TITLE}. Keywords like "Introductory Call", "Follow-up", 
# "June discussion", "Digital discussion", "Sampling Discussion" in the title provide strong clues
# about the meeting's primary purpose and stage of engagement.

# YOUR RESPONSE MUST ADAPT BASED ON ALL THIS CONTEXT:
#   - IF A DIRECT FOLLOW-UP IS INDICATED (from previous meeting data AND/OR a title like "Follow-up with X"):
#     - GENERATE THE BRIEF USING THE "DIRECT FOLLOW-UP MEETING STRUCTURE" outlined below.
#     - Focus heavily on continuity, previous discussions, action items, and client questions from past interactions.
#     - Standard brand synopsis, generic marketing history, and general market bifurcation become secondary or very brief.
#   - IF OTHER PREVIOUS INTERACTIONS (NOT a direct follow-up, e.g., different context or attendees) ARE NOTED OR IT'S A FIRST-TIME INTERACTION:
#     - Generate the brief using the "STANDARD MEETING STRUCTURE" 
#     - Provide a comprehensive overview suitable for a new engagement or a new line of discussion.

# ==========================================================================================
# === DIRECT FOLLOW-UP MEETING STRUCTURE (Use this if a direct follow-up is identified) ===
# ==========================================================================================

## Pre-Meeting Brief for Direct Follow-Up with {BRAND_NAME_FOR_BODY}

### 1. Recap of Previous Engagement & Current Status
    - Briefly summarize the key topics, proposals, or data points from the last direct interaction.
    - Note any decisions made or understandings reached.
    - State the current status of the engagement leading into this meeting (e.g., "Client was to review X," "NBH was to provide Y").
    - *Source this heavily from "Insights from Previous NBH Meetings," specifically "Key Discussion Points," and "Action Items" from {INTERNAL_NBH_DATA_SUMMARY}.*

### 2. Key Objectives for This Follow-Up Meeting
    - List 3-4 specific, actionable objectives for THIS meeting.
    - These should aim to progress the discussion, address pending items, or move towards a decision/pilot.
    - Example: "Finalize scope for pilot campaign," "Address client's questions on Z," "Present revised proposal based on feedback."

### 3. Client's Key Questions/Concerns (from previous meeting)
    - List any significant questions or concerns raised by the client in the previous meeting that need to be addressed.
    - For each, briefly outline the NBH response or how it will be tackled in this meeting.
    - *Source this from "Key Questions Raised by Client" in {INTERNAL_NBH_DATA_SUMMARY}.*
    - If no specific questions were logged, this section can be brief or state "No outstanding questions noted from previous discussion; this meeting will focus on next steps."

### 4. Review of Action Items from Previous Meeting
    - *Source this directly from "Action Items (Relevant for Follow-up)" within the INTERNAL_NBH_DATA_SUMMARY provided above.*
    - For EACH action item from the previous meeting, clearly list the following in a bulleted format. Ensure each piece of information is clearly labeled:
        - **Action Item:** [State the full action item text]
        - **Responsible:** [Indicate Client or NBH, or specific names if available in the source data]
        - **Progress/Status by NoBrokerHood:** [Describe NBH's progress or current status on this item]
        - **Pending Input/Action from Client:** [Describe what is needed from the client for this item, if anything]
    - Present each action item as a distinct set of these four bullet points.
    - If no action items were logged from the previous meeting, state: "No specific action items were formally logged from the previous meeting."

### 5. Revised/Tailored Proposed Solution and Offering for This Follow-Up
    - Present 2-4 highly tailored solutions or next steps.
    - These MUST directly build upon the previous discussion, address client feedback, or align with the agreed-upon direction.
    - For each concept:
        - Give it a brief, compelling title (e.g., "Lenskart Clear Vision Rewards").
        - Describe the solution concisely.
        - Explicitly state its "Alignment with Client Feedback/Previous Discussion" (e.g., "Directly addresses interest in coupon distribution," "Leverages client's focus on offline experience").
    - *This section should reflect an evolution of the previous proposals, not a fresh pitch unless the context changed dramatically.*
    - *Reference relevant NBH assets (digital, physical, data) only as they support these specific, evolved solutions.*

### 6. Personal History and Persona Analysis (Key Brand Attendees)
    - For each key brand attendee *who was also present in the previous relevant meeting*:
        - **Personal History:** If available from previous brief notes or new quick search, provide a brief professional summary.
        - **Persona Analysis (from previous interaction):** Based on their contributions, questions, and concerns in the *previous meeting*, describe their likely priorities, decision-making style, or focus areas relevant to this follow-up. (e.g., "Appeared data-driven, focused on ROI," "Showed keen interest in technical feasibility").
    - If new key attendees are joining, provide standard Personal History/Persona analysis for them.

### 7. Suggested Meeting Flow for This Follow-Up (Concise)
    - Outline a very brief meeting flow (e.g., 3-4 key stages).
    - Emphasis should be on addressing action items, discussing tailored solutions, and agreeing on concrete next steps.
    - Example:
        - Welcome & Quick Recap of Today's Focus (2 mins)
        - Addressing Pending Action Items & Client Questions (10 mins)
        - Discussion of Tailored Solutions/Proposals (10 mins)
        - Agreement on Next Steps & Timelines (5 mins)
        - Closing (3 mins)

### 8. Key Talking Points/Reminders for NBH Team
    - 2-3 bullet points for the NBH team to keep in mind during this specific follow-up.
    - e.g., "Reiterate benefit X based on client's previous concern," "Be prepared to discuss technical detail Y if asked," "Aim to secure commitment for Z."

# --- END OF DIRECT FOLLOW-UP MEETING STRUCTURE ---

# =====================================================================================
# === STANDARD MEETING STRUCTURE (Use this if NOT a direct follow-up, or first time) ===
# =====================================================================================

Part 1: Concise Executive Summary (keep the concise executive summary strictly below 300 words)

Brand Name: {BRAND_NAME_FOR_BODY}

#Brand & Launch Synopsis: name, launch date/event, one-sentence description of the brand. [IF NOT A DIRECT FOLLOW-UP: name, launch date/event, one-sentence description. IF DIRECT FOLLOW-UP: Briefly acknowledge brand and move on, e.g., "Continuing discussions with Activ8 Pilates X..."]
# [IF title suggests "Introductory": provide full synopsis. IF title suggests "Follow-up" or specific topic: "Continuing our discussion on [the specified topic if mentioned in the title] with {BRAND_NAME_FOR_BODY}..."]

#Client Profiles (concise - 3 bullet points): Briefly describe their market position, key initiatives, or focus areas. [IF NOT A DIRECT FOLLOW-UP: Describe market position, key initiatives. IF DIRECT FOLLOW-UP: Very briefly summarize current understanding or state "As previously discussed..."]

## Example bullet 1: [e.g., ## Legacy Innovator: Description]
## Example bullet 2: [e.g., ## Market Driver: Description]
## Example bullet 3: [e.g., ## Growth Focus: Description]

#Recent Marketing History (3 bullet points): List 3-4 significant recent marketing campaigns or activities.

## Example bullet 1: [e.g., ## Campaign A: Brief description]
## Example bullet 2: [e.g., ## Campaign B: Brief description]

#Marketing spend bifurcation: completely search the companies financials and the internet and give me the marketing spend bifurcation between traditional media and the new digital media. If the companies bifurcation is not available, give me the industry level bifurcation. Specifically check if there is any spend on activities linked to gated communities or tier 1 cities in India.



#Suggested Meeting Flow (bullet points with timings): (#keep the total time below 30 min). Please include the proposed solution section for sure.

Please format each step as a markdown sub-header (e.g., `## Icebreaking (5 mins)`) followed by a short description or bullet points.


Example:

    ## Icebreaking (5 mins)
    - "Good evening, ..."

    ## Need Finding (7 mins)
    # - "We understand your platform excels... [IF title suggests specific topic like 'Sampling': "Given our meeting focus on 'Sampling Discussion', could you share your current objectives and challenges specifically around product sampling?"]"

    ## Deck Presentation (8 mins)
    - Present NBH's digital and physical assets ...

    #Proposed solution and offering: [Briefly describe creative co-branded solutions or offerings relevant to the client and NoBrokerHood. Try to give atleast 5 solutions. Try to check on the web on some creative marketing campaigns which can be achieved through our collaboration(keep in mind our offerings while giving suggestions.), Keep the points precise  and concise as if you are presenting to a CXO. Keep each description to a maximum of 2 lines. Also check the the canva pdf master deck. Those are the offerings in the parent company Nobroker.com. We can offer the data in those various verticals as solutioning  to the company and leverage it to get them on board to Nobrokerhood.]
    
    Example Solution 1: [Brief Title/Concept: Description]
    Example Solution 2: [Brief Title/Concept: Description]

    ## Feedback & Brief Taking (3 mins)
    - "Based on what we've shared ..."

    ## Fixing Next Steps (1 min)
    - "Would you be open to ..."

    ## Closing (1 min)
    - "Thank you for your time ..."

Part 2: Detailed report: (keep the detailed report strictly below 700 words.)

# Usage of Nobroker/nobrokerhood data: 
    # The following is a summary of relevant internal data from NoBrokerHood's Google Drive. Use this information to identify relevant data points for {BRAND_NAME_FOR_BODY}.
    # Consider the {MEETING_TITLE} alongside this data.
    # This includes:
    #   - Insights from PREVIOUS NBH MEETINGS with {BRAND_NAME_FOR_BODY} (if any). Pay close attention to:
    #       - "Key Questions Raised by Client" in past meetings.
    #       - "Observed Brand Traits", "Identified Customer Needs", "Client Pain Points Discussed".
    #     Synthesize these previous meeting insights to inform the entire brief, especially the meeting flow and proposed solutions.    
    #   - Context on our General Pitch Deck.
    #   - Text from the Case Study PDF (analyze to identify distinct case studies: Brand, Objective, Activities, Results, Relevance to {BRAND_NAME_FOR_BODY}), need to also remember the slide numbers for these
    #   - Historical Campaign data for {BRAND_NAME_FOR_BODY} or similar brands/industries.
    #   - A description of available NoBroker.com platform metrics (e.g., leads, user data).
    #

    # Your task is to synthesize this information:
    #
    # 0. Previous Meeting Insights (If data provided under "Insights from Previous NBH Meetings"):
    #    a. Summarize any "Key Questions Raised by Client", "Observed Brand Traits", "Identified Customer Needs", or "Client Pain Points Discussed"
    #       from past meetings that are relevant for the upcoming discussion.
    #    b. This understanding should PERVADE the rest of your analysis and suggestions.
    #    c. If no previous meetings are noted, assume this is a first-time interaction and skip this section. 
    # 1. Case Studies (from section "Relevant Case Study Slides"): Try to use the latest case study slides provided. The case studies should be relevant for the brand, from the same industry or for same use case.
    #    Eg: For Easemytrip (Travel Agency) the relevant case studies would be for any travel or tourism related brand (like lufthansa / Mahindra Holidays / Makemytrip)
    #    For EACH "Case Study Slide X" provided:
    #    a. Identify the **Case Study Brand Name** (often in the title or a "Brand:" field).
    #    b. Extract **Campaign Introduction/Objective/Description** (look for "Introduction:", "Description:", or similar phrasing).
    #    c. Extract **Activities/Type of Activity** (look for "Activities:", "Type of Activity:").
    #    d. Extract **Duration** (look for "Duration:").
    #    e. Extract **Key Quantitative Results**. Look for labels like:
    #       - "Revenue:"
    #       - "Leads Generated:" (and "Total Orders" if present)
    #       - "Return on Investment (ROI):"
    #       - "CTR (Click-Through Rate):"
    #       - "Impressions:"
    #       - "Reach:"
    #       - "Footfall:"
    #       (If a specific metric label isn't present but a clear number related to results is, extract it with context.)
    #    f. Extract **City/Location** (look for "City:", "Location:").
    #    g. Based on the extracted information and the current meeting brand ({BRAND_NAME_FOR_BODY}, Industry: [Inferred industry of BRAND_NAME_FOR_BODY if available]), briefly state **why this case study is relevant** (e.g., "Relevant because it's for the same brand," or "Relevant due to similar industry: [Case Study Brand's Industry]," or "Demonstrates success with [specific activity type] which could interest {BRAND_NAME_FOR_BODY}").
    #    h. Also share the case study slide number if the Brand Manager wants to refer it for creatives and other details.
    #   Present the information for each extracted case study clearly, perhaps in a structured format or bullet points under a heading like "Case Study: [Case Study Brand Name] (from Slide X)".
    #    If multiple relevant case studies are provided, select and detail the top 1-3 most impactful or directly relevant ones for {BRAND_NAME_FOR_BODY}. If only one is highly relevant, focus on that.
    #
    # 2. Historical Campaigns (from section "Historical Campaign Data"): Try to use the latest historical campaign data provided
    #    (highlight 1-2 key examples, mention type, duration, results if available for {BRAND_NAME_FOR_BODY} or also mention campaigns done by competitors in the same industry)
    #
    # 3. NoBroker.com Platform Metrics (from section "NoBroker.com Platform Metrics"):
    #    ( You have access to the nobroker data like packers and movers as well as nobrokerhood data like number of maids, pets data etc. Provide the relevant data which we can use to present for this brand {BRAND_NAME_FOR_BODY}. Eg: if its a pet brand, total pet count can be provided, pet count in various cities can be provided to show where the targetting can be done etc. If it is a brand which can be purchased when people move accomodations or cities, move-in move-out data can be cited)
    #    These need to be separate bullet points
    # Present these insights concisely as actionable data points for the sales team. Ensure to reference the slide number when discussing a specific case study.
    # If information for a specific field (e.g., ROI) is not explicitly found on a slide, state "Not specified on slide." Do not invent data.

#Brand: [IF NOT A DIRECT FOLLOW-UP: Small detailed report on the brand]

a) categories that the brand exist in

b) Target audience of the brand

c) Peak seasons for the brand where each category sales happen

d) Brand persona

e) What the brand is doing these days and the areas where they are focusing more lately

#Recent Marketing history: List all the recent marketing history of the brand be it be campaigns or activities. Analyse the recent marketing campaigns of the brand as if you are a McKinsey or a BCG consultant and give me the analysis on various points, like the cohorting they are focusing on, etc. Don't make a big report. Keep the points precise  and concise as if you are presenting to a CXO. Keep each point to a max of 2 lines.

#Personal History and Persona Analysis: The personal history and persona analysis should be done for all the attendees from the brand side.

## Personal History: You have the name of the brand as well as name of the attendee from the brand. The attendee of the brand is most likely someone from the marketing department of the brand and he can be a product manager, brand manager etc. Having these in mind, conduct a targeted search and you most likely would find the linkedin profile of the attendee. Now, give the output on the basis of the linkedin profile that is obtained. The output should be a maximum of 4 lines.

## Persona Analysis: Using the information we have till now, try to search linkedin, fb, Instagram, twitter and all the social media regarding the person and try to build a persona of the person which contains his hobbies interests, etc. ***Important note: Always give information after confirmation that the information belongs to him/her. And the output should be of maximum 3-4 lines.

Use crisp, professional language and bullet-point formatting. Cite metrics and dates precisely. Produce both the concise summary and the detailed report in one output, clearly labelled.

"""



def generate_brief_with_gemini(gemini_llm_client, meeting_data, internal_data_summary_str):
    if not gemini_llm_client:
        return "Error: Gemini model not available."

    # ... (Format placeholders as in previous thought: nbh_attendee_names_str, brand_attendees_details_str etc.)
    nbh_attendee_names_str = ", ".join([att['name'] for att in meeting_data['nbh_attendees']])
    brand_attendees_info_str = "; ".join([f"{att['name']} ({att['email']})" for att in meeting_data['brand_attendees_info']])
    brand_attendee_names_only_str = ", ".join([att['name'] for att in meeting_data['brand_attendees_info']])


    prompt_filled = YOUR_DETAILED_PROMPT_TEMPLATE_GEMINI.format(
        MEETING_DATETIME=meeting_data['start_time_str'],
        MEETING_LOCATION=meeting_data['location'],
        BRAND_NAME=meeting_data['brand_name'], # This is the overall brand name
        BRAND_ATTENDEES_NAMES=brand_attendee_names_only_str,
        NBH_ATTENDEES_NAMES=nbh_attendee_names_str,
        BRAND_NAME_FOR_BODY=meeting_data['brand_name'], # Used within the prompt body
        MEETING_TITLE=meeting_data.get('title', 'N/A'), # <<< ADDED THIS
        BRAND_ATTENDEES_FULL_DETAILS=brand_attendees_info_str,
        INTERNAL_NBH_DATA_SUMMARY=internal_data_summary_str
    )
    
    grounding_tool = types.Tool(
        google_search=types.GoogleSearch()
    )

    # Configure generation settings
    config = types.GenerateContentConfig(
    # sampling parameters (formerly generation_config dict)
    temperature=0.7,
    top_p=0.95,
    top_k=40,

    # safety filters (formerly safety_settings list of dicts)
    safety_settings=[
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_HARASSMENT,
            threshold=types.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
            threshold=types.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
            threshold=types.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
            threshold=types.HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE,
        ),
    ],

    # tools (formerly a separate config.tools or tools argument)
    tools=[grounding_tool],
)

    print(f"  Sending request to Gemini for brand: {meeting_data['brand_name']}...")
    try:
        # Use Google Search grounding by adding tools=[{"tool": "google_search"}]
        response = gemini_llm_client.models.generate_content(
            model="gemini-2.5-flash",  # Use the latest Gemini model
            contents=prompt_filled,
            config=config,
            # tools=[{"tool": "google_search"}]  # Enable Google Search grounding >> not working on API calls rights now
        )
        # print(f"Gemini raw response: {response}") # For debugging
        if response.prompt_feedback and response.prompt_feedback.block_reason:
            return f"Error: Prompt blocked by Gemini. Reason: {response.prompt_feedback.block_reason_message or response.prompt_feedback.block_reason}"

        if not response.candidates:
             return f"Error: Gemini returned no candidates. Feedback: {response.prompt_feedback}"

        # Ensure text is extracted correctly
        brief_content = ""
        for part in response.candidates[0].content.parts:
            brief_content += part.text
        
        if not brief_content.strip():
            return "Error: Gemini returned an empty brief."
            
        return brief_content
    except Exception as e:
        print(f"  Error during Gemini API call: {e}")
        # Check for specific API errors if needed, e.g. google.api_core.exceptions.ResourceExhausted
        if "429" in str(e) or "ResourceExhausted" in str(e): # Quota issue
             return "Error: Gemini API quota likely exceeded. Please check your Google Cloud/AI Studio quotas."
        return f"Error: Exception during Gemini call: {e}"

# --- Email Sending ---
def create_email_message(sender, to_emails_list, subject, message_text_html):
    # ... (Using MIMEText with 'html' for better formatting) ...
    message = MIMEText(message_text_html, 'html', 'utf-8') # Send as HTML
    message['to'] = ", ".join(to_emails_list)
    message['from'] = sender
    message['subject'] = subject
    raw_message = base64.urlsafe_b64encode(message.as_bytes())
    return {'raw': raw_message.decode()}

def send_gmail_message(gmail_service, user_id, message_body):
    # ... (same as before) ...
    """
    Sends an email message using the Gmail API.
    
    Parameters:
        user_id (str): The user's email address or the special value 'me' to indicate the authenticated user.
        message_body (dict): The message payload, typically created by create_email_message().
    
    Returns:
        dict or None: The API response containing the sent message details, or None if sending fails.
    """
    if not gmail_service:
        print("  Gmail service not available. Cannot send email.")
        return None
    try:
        message = (gmail_service.users().messages().send(userId=user_id, body=message_body).execute())
        print(f'  Message Id: {message["id"]} sent.')
        return message
    except HttpError as error:
        print(f'  An error occurred sending email: {error}')
        return None

def send_brief_email(gmail_service, meeting_data, brief_content):
    """
    Sends a pre-meeting brief email to NBH internal attendees for a given meeting.
    
    The function filters out excluded emails from the NBH attendee list, converts the brief content from markdown to HTML, formats the email body, and sends the email using the Gmail API. If no eligible recipients are found, the function exits without sending an email.
    """
    EXCLUDED_EMAILS = {AGENT_EMAIL.lower(), "pia.brand@nobroker.in","pia.hood@nobroker.in"} # Define a set of excluded emails

    nbh_recipient_emails = []
    attendees_list = meeting_data.get('nbh_attendees', []) 
    if isinstance(attendees_list, list): # Extra safety check
        for att in attendees_list:
            # Ensure 'att' is a dictionary and 'email' key exists
            if isinstance(att, dict) and 'email' in att:
                attendee_email = att.get('email')
                if attendee_email and isinstance(attendee_email, str) and attendee_email.lower() not in EXCLUDED_EMAILS:
                    nbh_recipient_emails.append(attendee_email)
    
    # For testing, override recipients:
    # nbh_recipient_emails = [ADMIN_EMAIL_FOR_NOTIFICATIONS]
    # print(f"DEBUG: Intended brief recipients: {nbh_recipient_emails}")

    if not nbh_recipient_emails:
        print(f"  No NBH recipients (other than brandvmeet) for '{meeting_data['title']}'. Brief not emailed.")
        return

    email_subject = f"Pre-Meeting Brief: {meeting_data['title']} with {meeting_data['brand_name']}"
    
    # Convert markdown-like brief (from LLM) to basic HTML for email
    html_brief_content = markdown.markdown(brief_content)
    # html_brief_content = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', html_brief_content) # Basic bold
    # Add more markdown to HTML conversions if needed (e.g., for lists, headers)  
    # For headers like # Detail Report, ## Sub-header
    # html_brief_content = html_brief_content.replace("# ", "<h3>").replace("\n", "</h3><br>") # Simple h3 for main headers
    
    email_body_html = f"""
    <html><body>
    <p>Hello Team,</p>
    <p>Please find the pre-meeting brief for your upcoming meeting:</p>
    <hr>
    {html_brief_content}
    <hr>
    <p>Best regards,<br>NBH Meeting Prep Agent</p>
    </body></html>
    """
    email_message = create_email_message(
        sender=AGENT_EMAIL,
        to_emails_list=nbh_recipient_emails,
        subject=email_subject,
        message_text_html=email_body_html
    )
    print(f"  FINAL CHECK: Sending brief for '{meeting_data['title']}' TO: {nbh_recipient_emails} FROM: {AGENT_EMAIL}")
    send_gmail_message(gmail_service, 'me', email_message)

def send_notification_email(gmail_service, subject, body_html, recipient=ADMIN_EMAIL_FOR_NOTIFICATIONS):
    if not recipient:
        print("  Admin notification email not set. Skipping notification.")
        return
    
    # Always send a copy to brandvmeet for record-keeping
    recipients = list(set([recipient, AGENT_EMAIL]))

    email_message = create_email_message(
        sender=AGENT_EMAIL,
        to_emails_list=recipients,
        subject=subject,
        message_text_html=body_html
    )
    send_gmail_message(gmail_service, 'me', email_message)

# Here are some useful functions to update meeting data in the master sheet

def read_data_from_sheets(sheet_id, sheets_service, range):

    try:
        result = (
                sheets_service.spreadsheets()
                .values()
                .get(spreadsheetId=sheet_id, range=range)
                .execute()
            )
        sheet_data = result.get("values", [])
        print(f"{len(sheet_data)} rows retrieved")
        return sheet_data
    except HttpError as error:
        print(f"An error occurred: {error}")

# Then check which events have not been updated in the google sheets
def events_to_update(meeting_ids, events):
    events_to_update = []
    for event in events:
        arr = [event["id"]]
        if arr not in meeting_ids:
            events_to_update.append(event)
    if not events_to_update:
        print("No new meetings to update")
        return None
    else:
        return events_to_update
    

def update_events_in_sheets(sheet_id, events_to_update, sheets_service, excluded_emails):
    for i, event in enumerate(events_to_update):
        id = event["id"]
        title = event["summary"]
        date = event["start"].get("date", event["start"].get("dateTime"))
        if 'T' in date:
            date = datetime.datetime.fromisoformat(date).date().isoformat()
        attendees = event.get("attendees")
        if attendees:
            emails = [attendee["email"] for attendee in attendees]
            nobroker_attendee = []
            client_attendee = []
            
            for email in emails:
                if "nobroker" in email:
                    pseudo_name=False
                    for name in excluded_emails:
                        if name in email:
                            pseudo_name=True
                            break
                    if pseudo_name==True:
                        continue
                    else:
                        nobroker_attendee.append(email)
                else:
                    client_attendee.append(email)
            row = [id, title, date, f"{nobroker_attendee}", f"{client_attendee}"]
            values = [row]
            
            body = {
                'values': values
            }
            result = sheets_service.spreadsheets().values().append(
                spreadsheetId=sheet_id,
                range="Meeting_data",
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            print(f"Appended row: {title}")
            if (i+1)%50 == 0:
                print("Sleep initiated")
                time.sleep(50)

# Function to create a google document for the brief
def create_google_doc_in_folder(drive_service, folder_id, doc_name):
    file_metadata = {
        'name': doc_name,
        'mimeType': 'application/vnd.google-apps.document',
        'parents': [folder_id]
    }
    created = drive_service.files().create(
        body=file_metadata,
        fields='id, name, parents'
    ).execute()
    print(f"Created Google Doc: {created['name']} (ID: {created['id']})")
    return created['id']

# Function to write content to a Google Doc
def write_into_doc(docs_service, doc_id, text):
    requests = [
        {
            'insertText': {
                'location': { 'index': 1 },
                'text': text
            }
        }
    ]
    
    try:
        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={'requests': requests}
        ).execute()
    except:
        print("An error occured while writing into google doc")

# --- Main Execution Logic ---
def main():
    """
    Main orchestration function for automated pre-meeting brief generation and notification.
    
    This function coordinates the end-to-end workflow for preparing and emailing pre-meeting briefs for upcoming client meetings managed by the agent account. It initializes required Google Workspace services and the Gemini LLM, fetches upcoming calendar events, and processes each event as follows:
    
    - Skips events already processed or tagged.
    - Extracts meeting details and identifies the brand and industry using the LLM.
    - Retrieves and summarizes relevant internal NBH data for the brand.
    - Determines if the meeting is a direct follow-up or involves separate historical threads, and sends leadership alert emails as needed.
    - Generates a detailed pre-meeting brief using the LLM and internal data.
    - Emails the brief to NBH attendees, tags the event as processed, sets a 1-hour reminder, and records the event as processed.
    
    Handles error conditions gracefully, including missing services, ambiguous brand extraction, and LLM failures, with appropriate notifications and fallback logic.
    """
    print(f"Script started at {datetime.datetime.now()}")
    print(f"Using NBH GDrive Folder ID: {NBH_GDRIVE_FOLDER_ID}")

    calendar_token = os.getenv("CALENDAR_TOKEN")
    gmail_token = os.getenv("GMAIL_TOKEN")
    drive_token = os.getenv("DRIVE_TOKEN")
    sheets_token = os.getenv("SHEET_TOKEN")
    docs_token = os.getenv("DOCS_TOKEN")  # Token for Google Docs API

    # Initialize Google Services
    # Use a combined token file strategy or separate ones. Separate is fine.
    calendar_service = get_google_service('calendar', 'v3', SCOPES, calendar_token)
    gmail_service = get_google_service('gmail', 'v1', SCOPES, gmail_token)
    drive_service = get_google_service('drive', 'v3', SCOPES, drive_token)
    sheets_service = get_google_service('sheets', 'v4', SCOPES, sheets_token)
    docs_service = get_google_service('docs', 'v1', SCOPES, docs_token)  # Docs service for creating briefs
    gemini_llm_client = configure_gemini()


    if not calendar_service: # Critical service
        print("Exiting: Calendar service failed to initialize.")
        return

    upcoming_events = get_upcoming_meetings(calendar_service)
    if not upcoming_events:
        print('No upcoming events found for agent email that need processing.')
        return

    # Updating events in the master sheet

    master_sheet_id = "1xtB1KUAXJ6IKMQab0Sb0NJfQppCKLkUERZ4PMZlNfOw"
    meeting_ids = read_data_from_sheets(master_sheet_id, sheets_service, "Meeting_data!A2:A")

    events_to_update_list = events_to_update(meeting_ids, upcoming_events)

    if not events_to_update_list:
        print("No new meetings to update in master sheet.")
    else:
        print(f"{len(events_to_update_list)} new meetings found")
        update_events_in_sheets(master_sheet_id, events_to_update_list, sheets_service, EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP)

    updated_meeting_ids = read_data_from_sheets(master_sheet_id, sheets_service, "Meeting_data!A2:A")

    processed_ids_local_file = load_processed_event_ids()

    for event_payload in upcoming_events:
        event_id = event_payload['id']
        event_summary = event_payload.get('summary', 'No Title')
        event_description_for_tag_check = event_payload.get('description')

        print(f"\nProcessing event: '{event_summary}' (ID: {event_id})")

        # Step 1: Check if the event has already been processed
        if is_event_already_tagged(event_description_for_tag_check):
            print(f"  Skipping event '{event_summary}': Already tagged as processed.")
            continue
        
        if event_id in processed_ids_local_file:
            print(f"  Skipping event '{event_summary}': Found in local processed file.")
            continue

        # Step 2: Extract basic meeting info (attendees, raw title, etc.)
        meeting_data_result = extract_meeting_info(event_payload, AGENT_EMAIL, NBH_SERVICE_ACCOUNTS_TO_EXCLUDE)

        # Step 3: Handle the possible "skip" results from the extraction
        if meeting_data_result is None: # Case where agent is not an attendee
            print(f"  Skipping event '{event_summary}': Agent is not an attendee.")
            save_processed_event_id(event_id)
            tag_event_as_processed(calendar_service, event_id)
            continue
        # Skipping this condition so that physical meetings can also be processed
        # if meeting_data_result == "NO_EXTERNAL_ATTENDEES":
        #     print(f"  Event '{event_summary}': No external attendees. No brief needed.")
        #     save_processed_event_id(event_id)
        #     tag_event_as_processed(calendar_service, event_id)
        #     continue

        # Step 4: If we are here, extraction was successful. Assign the result to meeting_data.
        # This is the key fix: assign the dictionary before trying to use it.
        meeting_data = meeting_data_result

        # Step 5: Use the LLM to get the brand name and industry from the raw title
        print(f"  Using LLM to extract brand details from title: '{meeting_data['title']}'")
        brand_details = get_brand_details_from_title_with_llm(gemini_llm_client, meeting_data['title'])

        # Step 6: Handle ambiguous result from the LLM
        if brand_details['brand_name'] == 'Unknown Brand' or brand_details['brand_name'].lower() == 'unknown':
            print(f"  Event '{meeting_data['title']}': Title is ambiguous for brand extraction by LLM.")
            # Your notification logic for ambiguous titles can go here if needed.
            # Example:
            # ambiguous_body_html = f"..."
            # send_notification_email(...)
            save_processed_event_id(event_id)
            tag_event_as_processed(calendar_service, event_id) # Tag it so we don't retry
            # Updating the unknown brand name and industry in the master sheet
            index_of_event = updated_meeting_ids.index([event_id]) + 2 # +2 because A1 is header and A2 is first data row
            print(f"  Updating master sheet for event ID '{event_id}' at row {index_of_event} with brand 'Unknown")
            update_values = [[brand_details['brand_name'], brand_details['industry']]]
            body = {
            'values': update_values
            }
            try:
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=master_sheet_id,
                    range=f"Meeting_data!F{index_of_event}:G{index_of_event}",
                    valueInputOption='RAW',
                    body=body
                ).execute()
                print(f"  Master sheet updated successfully for event ID '{event_id}'.")
            except HttpError as error:
                print(f"  Error updating master sheet for event ID '{event_id}': {error}")
            continue

        # Step 7: Merge the successful LLM results into the main meeting_data dictionary
        meeting_data.update(brand_details)

        current_brand_name_for_meeting = meeting_data['brand_name']
        target_brand_industry = meeting_data['industry']

        # Updating the brand name and industry in the master sheet
        index_of_event = updated_meeting_ids.index([event_id]) + 2 # +2 because A1 is header and A2 is first data row
        print(f"  Updating master sheet for event ID '{event_id}' at row {index_of_event} with brand '{current_brand_name_for_meeting}' and industry '{target_brand_industry}'")
        update_values = [[current_brand_name_for_meeting, target_brand_industry]]
        body = {
            'values': update_values
        }
        try:
            sheets_service.spreadsheets().values().update(
                spreadsheetId=master_sheet_id,
                range=f"Meeting_data!F{index_of_event}:G{index_of_event}",
                valueInputOption='RAW',
                body=body
            ).execute()
            print(f"  Master sheet updated successfully for event ID '{event_id}'.")
        except HttpError as error:
            print(f"  Error updating master sheet for event ID '{event_id}': {error}")
        
        print(f"  LLM identified Brand: '{current_brand_name_for_meeting}', Industry: '{target_brand_industry}'")

        # --- THIS IS THE CORRECTED AND SIMPLIFIED BLOCK ---
        
        # Step 1: Check if necessary services are available.
        if drive_service and sheets_service:
            # If services are available, call the function to get the real data.
            internal_data_result = get_internal_nbh_data_for_brand(
                drive_service=drive_service,
                sheets_service=sheets_service,
                gemini_llm_client=gemini_llm_client,
                current_target_brand_name=current_brand_name_for_meeting,
                target_brand_industry=target_brand_industry,  # <-- Pass the extracted industry
                current_meeting_data=meeting_data,
                EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP=EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP,
                AGENT_EMAIL=AGENT_EMAIL
            )
        else:
            # If services are NOT available, create the default/fallback structure.
            print(f"  Drive/Sheets service not available. Skipping internal data fetch for '{current_brand_name_for_meeting}'.")
            internal_data_result = {
                "llm_summary_string": "Internal NBH Data: Not fetched due to Drive/Sheets service issues.",
                "is_overall_direct_follow_up": False,
                "has_previous_interactions": False,
                "condensed_past_meetings_for_alert": []
            }
        
        # Step 2: Extract the summary string for the LLM brief from the result (either real or default).
        internal_nbh_data_for_brand_str = internal_data_result["llm_summary_string"]
        
        # Step 3: Now, use the result for the leadership alert logic.
        has_prev_interactions_in_main = internal_data_result.get("has_previous_interactions", False)
        is_overall_follow_up_in_main = internal_data_result.get("is_overall_direct_follow_up", False)
        
        
        
                
        # --- >>> LEADERSHIP ALERT LOGIC (FINAL, CORRECTED VERSION) <<< ---

        # Step 1: Extract the flags and data we need from the internal data check.
        is_direct_follow_up = internal_data_result.get("is_overall_direct_follow_up", False)
        has_other_interactions = internal_data_result.get("has_other_past_interactions", False)
        condensed_meetings_for_alert = internal_data_result.get("condensed_past_meetings_for_alert", [])

        # Helper variables for the email body
        upcoming_meeting_title = meeting_data.get('title', 'N/A')
        upcoming_nbh_attendees_list = [att['name'] for att in meeting_data.get('nbh_attendees', [])]
        upcoming_nbh_attendees_str = ", ".join(upcoming_nbh_attendees_list) if upcoming_nbh_attendees_list else "N/A"


        # SCENARIO 1: "Hybrid" Engagement - A follow-up, but other separate threads also exist.
        if is_direct_follow_up and has_other_interactions:
            print("DEBUG: HYBRID SCENARIO DETECTED. Sending nuanced leadership alert.")
            
            alert_subject = f"FYI: Complex Engagement with {current_brand_name_for_meeting} (Follow-up & Separate Threads)"
            
            alert_body_html = f"""
            <html><head><style> body {{ font-family: Arial, sans-serif; }} li {{ margin-bottom: 8px; }} </style></head>
            <body>
                <p>Hello Leadership Team,</p>
                <p>A new meeting has been scheduled with <b>{current_brand_name_for_meeting}</b>. This engagement is complex and requires coordination:</p>
                <ul style="list-style-type:square;">
                    <li>It appears to be a <b>direct follow-up</b> to some recent discussions.</li>
                    <li>However, there are also <b>other, separate historical interactions</b> with this brand.</li>
                </ul>
                <p><b>Upcoming Meeting Details:</b></p>
                <ul>
                    <li><b>Title:</b> {upcoming_meeting_title}</li>
                    <li><b>NBH Attendees:</b> {upcoming_nbh_attendees_str}</li>
                </ul>
                <p>This highlights a need for internal coordination. Context on the separate past interactions is below for awareness:</p>
                <ul>
            """
            if condensed_meetings_for_alert:
                for past_mtg in condensed_meetings_for_alert:
                    alert_body_html += f"<li><b>{past_mtg['date']}:</b> {past_mtg['discussion_summary']} (NBH Team: {past_mtg['nbh_team']})</li>"
            alert_body_html += "</ul><p>Best regards,<br>NBH Meeting Prep Agent</p></body></html>"
            
            # --- CORRECT EMAIL SENDING LOGIC ---
            if gmail_service and leadership_emails:
                email_message = create_email_message(
                    sender=AGENT_EMAIL,
                    to_emails_list=leadership_emails,
                    subject=alert_subject,
                    message_text_html=alert_body_html
                )
                send_gmail_message(gmail_service, 'me', email_message)
                print(f"    Leadership alert for HYBRID scenario with {current_brand_name_for_meeting} sent.")
            else:
                print(f"    WARNING: Leadership alert for {current_brand_name_for_meeting} NOT sent (Gmail service or recipient list unavailable).")


        # SCENARIO 2: "Purely Separate" Engagement - Not a follow-up, but other past interactions exist.
        elif has_other_interactions and not is_direct_follow_up:
            print("DEBUG: PURELY SEPARATE THREAD DETECTED. Sending standard leadership alert.")
            
            alert_subject = f"FYI: New Meeting Scheduled with Existing Brand - {current_brand_name_for_meeting}"
            
            alert_body_html = f"""
            <html><head><style> body {{ font-family: Arial, sans-serif; }} li {{ margin-bottom: 8px; }} </style></head>
            <body>
                <p>Hello Leadership Team,</p>
                <p>A new meeting has been scheduled with <b>{current_brand_name_for_meeting}</b>. This meeting does <b>NOT</b> appear to be a direct follow-up to recent discussions.</p>
                <p>This could indicate a new opportunity or a new NBH team engaging with the client.</p>
                <p><b>Upcoming Meeting Details:</b></p>
                <ul>
                    <li><b>Title:</b> {upcoming_meeting_title}</li>
                    <li><b>NBH Attendees:</b> {upcoming_nbh_attendees_str}</li>
                </ul>
                <p><b>Summary of Past Interactions (for context):</b></p>
                <ul>
            """
            if condensed_meetings_for_alert:
                for past_mtg in condensed_meetings_for_alert:
                    alert_body_html += f"<li><b>{past_mtg['date']}:</b> {past_mtg['discussion_summary']} (NBH Team: {past_mtg['nbh_team']})</li>"
            alert_body_html += "</ul><p>Best regards,<br>NBH Meeting Prep Agent</p></body></html>"
            
            # --- CORRECT EMAIL SENDING LOGIC ---
            if gmail_service and leadership_emails:
                email_message = create_email_message(
                    sender=AGENT_EMAIL,
                    to_emails_list=leadership_emails,
                    subject=alert_subject,
                    message_text_html=alert_body_html
                )
                send_gmail_message(gmail_service, 'me', email_message)
                print(f"    Leadership alert for SEPARATE THREAD with {current_brand_name_for_meeting} sent.")
            else:
                print(f"    WARNING: Leadership alert for {current_brand_name_for_meeting} NOT sent (Gmail service or recipient list unavailable).")

        else:
            # This covers the "clean" cases: a brand-new meeting or a simple follow-up with no other threads.
            print("DEBUG: No leadership alert needed (Clean new meeting or simple follow-up).")
    


        if not gemini_llm_client:
            print(f"  Skipping brief generation for '{meeting_data['title']}': Gemini LLM not available.")
            # Don't mark as processed yet, maybe LLM will be available next run
            continue
        
        if not meeting_data.get('nbh_attendees'): # Check if any NBH humans are there
            print(f"  Event '{meeting_data['title']}': No NBH attendees (other than brandvmeet) to send brief to.")
            save_processed_event_id(event_id)
            tag_event_as_processed(calendar_service, event_id)
            continue


        print(f"  Proceeding with brief generation for: {meeting_data['brand_name']}")
        generated_brief = generate_brief_with_gemini(gemini_llm_client, meeting_data, internal_nbh_data_for_brand_str)

        if "Error:" in generated_brief or not generated_brief.strip(): # Check for errors from LLM
            print(f"  Failed to generate brief for '{meeting_data['title']}': {generated_brief}")
            error_body_html = f"""
            <html><body><p>The pre-meeting brief agent encountered an error while generating the brief for:</p>
            <p><b>Event:</b> {meeting_data['title']}<br>
            <b>Brand:</b> {meeting_data['brand_name']}<br>
            <b>Scheduled:</b> {meeting_data['start_time_str']}</p>
            <p><b>Error details:</b> {generated_brief}</p></body></html>"""
            send_notification_email(gmail_service,
                                    f"Error Generating Brief: {meeting_data['title']}",
                                    error_body_html)
            # Don't tag as fully processed if LLM fails, maybe it's temporary.
            # Or use a different tag like [NBH_BRIEF_AGENT_ERROR_V1]
        else:
            print(f"  Successfully generated brief for '{meeting_data['title']}'.")
            send_brief_email(gmail_service, meeting_data, generated_brief)
            tag_event_as_processed(calendar_service, event_id) # Tag only on full success
            set_one_hour_email_reminder(calendar_service, event_id) # Add this call
            save_processed_event_id(event_id)
            # --- Create Google Doc for the brief ---
            BRIEF_FOLDER_ID = "1RhhsFq5NGC2QtHPj8FQaU5BfhxJR5R6I"
            doc_id = create_google_doc_in_folder(
                drive_service,
                BRIEF_FOLDER_ID,
                f"Pre-Meeting Brief - {meeting_data['brand_name']} - {meeting_data['title']}"
            )
            if doc_id:
                write_into_doc(docs_service, doc_id=doc_id, text=generated_brief)
                # Updating doc link in master sheet
                index_of_event = updated_meeting_ids.index([event_id]) + 2 # +2 because A1 is header and A2 is first data row
                update_values = [[f"https://docs.google.com/document/d/{doc_id}"]]
                body = {
                    'values': update_values
                }
                try:
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=master_sheet_id,
                        range=f"Meeting_data!H{index_of_event}:H{index_of_event}",
                        valueInputOption='RAW',
                        body=body
                    ).execute()
                    print(f"  Master sheet updated with Google Doc link for event ID '{event_id}'.")
                except HttpError as error:
                    print(f"  Error updating master sheet with Google Doc link for event ID '{event_id}': {error}")
                print(f"  Google Doc created and content written for '{meeting_data['title']}'.")
            
        

    #print(f"Script finished at {datetime.datetime.now()}")

if __name__ == '__main__':
    main()
