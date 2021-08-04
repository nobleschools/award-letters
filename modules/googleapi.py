#!python3

"""Module file for all interaction with Google API"""
import os
import pickle
import socket
import gspread

from googleapiclient import errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth.transport.requests import AuthorizedSession


CREDENTIAL_STORE_DIR = ".credentials"
CREDENTIAL_STORE_FILE = "award-letters.json"
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]
CLIENT_SECRET_FILE = "client_secret.json"
DEFAULT_TIMEOUT = 300.0  # 5 minutes total timeout
APPLICATION_NAME = "Award Letter Trackers"
SCRIPT_ID = "Mnmyh2DYQEzLuWOvbDD0zJZ76E4tkxNYa"
# SCRIPT_ID = 'M3ZRRi0AvnjoCeQzL3JszW3d8W73qGbVI'
SCRIPT_V = "v1"
DRIVE_V = "v3"


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        credentials, the obtained credential.
    """
    if not os.path.exists(CREDENTIAL_STORE_DIR):
        os.makedirs(CREDENTIAL_STORE_DIR)
    credential_path = os.path.join(CREDENTIAL_STORE_DIR, CREDENTIAL_STORE_FILE)
    credentials = None
    if os.path.exists(credential_path):
        with open(credential_path, "rb") as token:
            credentials = pickle.load(token)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            secret_path = os.path.join(CREDENTIAL_STORE_DIR, CLIENT_SECRET_FILE)
            flow = InstalledAppFlow.from_client_secrets_file(secret_path, SCOPES)
            credentials = flow.run_local_server()
        with open(credential_path, "wb") as token:
            pickle.dump(credentials, token)

    return credentials


def get_drive_service(credentials=None):
    """
    Returns a drive service, optionally taking supplied credentials
    """
    if not credentials:
        credentials = get_credentials()
    try:
        return build("drive", DRIVE_V, credentials=credentials)
    except AttributeError as e:
        print(f"Credentials attribute error {credentials.items()}")
        raise e


def gspread_client(credentials):
    """
    Returns a gspread client object.
    Google has deprecated Oauth2, but the gspread library still uses the creds
    from that system, so this function bypasses the regular approach and creates
    and authorizes the client here instead.
    Code copied from answer here: https://github.com/burnash/gspread/issues/472
    """
    gc = gspread.Client(auth=credentials)
    gc.session = AuthorizedSession(credentials)
    return gc


def move_spreadsheet_and_share(s_id, folder, credentials=None):
    """
    Moves a file to the folder by adding that as a parent
    Also sets permissions for access with anyone with the link
    """
    # Switch parents
    service = get_drive_service(credentials)
    file = service.files().get(fileId=s_id, fields="parents").execute()
    previous_parents = ",".join(file.get("parents"))
    file = (
        service.files()
        .update(
            fileId=s_id,
            addParents=folder,
            removeParents=previous_parents,
            fields="id, parents",
        )
        .execute()
    )

    # Fix permissions
    file_permission = {"role": "writer", "type": "anyone", "withLink": True}
    service.permissions().create(
        fileId=s_id, body=file_permission, fields="id"
    ).execute()


def call_script_service(request, credentials=None, service=None):
    """
    Handles calls to script service if provided a request dict
    Credentials and/or service can be passed

    Handles errors in the function, but returns the request response or
    a None response if not available
    """
    socket.setdefaulttimeout(DEFAULT_TIMEOUT)
    if not service:
        if not credentials:
            credentials = get_credentials()
        service = build("script", SCRIPT_V, credentials=credentials)

    try:
        request["devMode"] = "true"  # runs last save instead of last deployed
        response = service.scripts().run(body=request, scriptId=SCRIPT_ID).execute()

        if "error" in response:
            # The API executes, but the script returned an error.

            # Extract the first (and only) set of error details. The values of
            # this object are the script's 'errorMessage' and 'errorType', and
            # and list of stack trace elements.
            error = response["error"]["details"][0]
            print("Script error message: {}".format(error["errorMessage"]))

            if "scriptStackTraceElements" in error:
                # There may not be a stacktrace if the script didn't start
                # executing.
                print("Script error stacktrace:")
                for trace in error["scriptStackTraceElements"]:
                    print("\t{1}: {0}".format(trace["function"], trace["lineNumber"]))
            return None
        else:
            # return the response:
            return response["response"].get("result", {})

    except errors.HttpError as e:
        # The API encountered a problem before the script started executing.
        print(e.content)
