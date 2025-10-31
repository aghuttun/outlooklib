"""
Provide generic access to Outlook functionalities via the Microsoft Graph API.

This module defines an Outlook client for interacting with Outlook mail folders, messages, and attachments.
It supports authentication, folder and message management, and downloading attachments, using a generic interface that
is not tailored to specific project requirements.

Classes
-------
Outlook
    Client for interacting with Outlook via the Microsoft Graph API.

Notes
-----
- Do not modify this library for project-specific needs.
- All features are designed to remain generic and reusable.
"""

import base64
import dataclasses
import json
import logging
import os
import re
from typing import Any, Type

# TypeAdapter v2 vs parse_obj_as v1
from pydantic import BaseModel, parse_obj_as  # pylint: disable=no-name-in-module
import requests
from .models import (
    ListFolders,
    ListMessages,
    MoveMessage,
)

# Creates a logger for this module
logger = logging.getLogger(__name__)


class Outlook(object):
    """
    Provide a client for interacting with Outlook via the Microsoft Graph API.

    Use this class to authenticate, manage mail folders, retrieve and move messages, and download attachments in a
    generic and reusable manner. Do not modify this class for project-specific needs.

    Attributes
    ----------
    _logger : logging.Logger
        Logger instance for the Outlook client.
    _session : requests.Session
        Session object for making HTTP requests.
    _configuration : Configuration
        Configuration dataclass containing credentials and settings.

    Methods
    -------
    auth()
        Authenticate and obtain an access token using client credentials.
    change_client_email(email)
        Change the current client email to be accessed.
    change_folder(id)
        Change the current mail folder to be accessed.
    list_folders(save_as=None)
        Retrieve a list of mail folders for the authenticated user.
    list_messages(filter, save_as=None)
        Retrieve the top 100 messages from the specified folder, filtered by a given condition.
    move_message(id, to, save_as=None)
        Move a message from the current folder to another specified folder.
    delete_message(id)
        Delete a message from the current folder.
    download_message_attachment(id, path, index=False)
        Download attachments from an email message.

    Notes
    -----
    - All features are designed to remain generic and reusable.
    - Refer to Microsoft Graph API documentation for valid folder names, IDs, and filter syntax.
    """

    @dataclasses.dataclass
    class Configuration(object):
        """
        Store configuration settings for the Outlook client.

        Use this dataclass to hold API domains, credentials, tokens, and user context for Microsoft Graph API access.

        Attributes
        ----------
        api_domain : str or None
            The domain for the Microsoft Graph API (e.g., "graph.microsoft.com").
        api_version : str or None
            The version of the Microsoft Graph API to use (e.g., "v1.0").
        sp_domain : str or None
            The SharePoint domain associated with the client.
        client_id : str or None
            The Azure client ID for authentication.
        tenant_id : str or None
            The Azure tenant ID for authentication.
        client_secret : str or None
            The secret key for the Azure client.
        token : str or None
            The OAuth2 access token for API requests.
        client_email : str or None
            The email address of the client user.
        client_folder : str
            The mail folder to access (default is "Inbox").
        """

        api_domain: str | None = None
        api_version: str | None = None
        sp_domain: str | None = None
        client_id: str | None = None
        tenant_id: str | None = None
        client_secret: str | None = None
        token: str | None = None
        client_email: str | None = None
        client_folder: str = "Inbox"

    @dataclasses.dataclass
    class Response:
        """
        Represent the response from Outlook client methods.

        Parameters
        ----------
        status_code : int
            The HTTP status code returned by the API request.
        content : Any, optional
            The content of the response, such as deserialized data or None if not applicable.

        Attributes
        ----------
        status_code : int
            The HTTP status code of the response.
        content : Any
            The content returned by the API, if any.
        """

        status_code: int
        content: Any = None

    def __init__(
        self,
        client_id: str,
        tenant_id: str,
        client_secret: str,
        sp_domain: str,
        client_email: str,
        client_folder: str = "Inbox",
        custom_logger: logging.Logger | None = None,
    ) -> None:
        """
        Initialize the Outlook client with the provided credentials and configuration.

        Parameters
        ----------
        client_id : str
            Specify the Azure client ID used for authentication.
        tenant_id : str
            Specify the Azure tenant ID associated with the client.
        client_secret : str
            Specify the secret key for the Azure client.
        sp_domain : str
            Specify the SharePoint domain.
        client_email : str
            Specify the client email account.
        client_folder : str, optional
            Specify the client folder. Defaults to "Inbox".
        custom_logger : logging.Logger, optional
            Provide a logger instance to use. If None, create a default logger.

        Notes
        -----
        Set up logging, initialize the HTTP session, store configuration, select the mail folder, and authenticate the
        client.
        """
        # Init logging
        # Use provided logger or create a default one
        self._logger = custom_logger or logging.getLogger(name=__name__)

        # Init variables
        self._session: requests.Session = requests.Session()
        api_domain = "graph.microsoft.com"
        api_version = "v1.0"

        # Credentials/Configuration
        self._configuration = self.Configuration(
            api_domain=api_domain,
            api_version=api_version,
            sp_domain=sp_domain,
            client_id=client_id,
            tenant_id=tenant_id,
            client_secret=client_secret,
            token=None,
            client_email=client_email,
            client_folder=client_folder,
        )

        # Handle folder
        self.change_folder(id=client_folder)

        # Authenticate
        self.auth()

    def __del__(self) -> None:
        """
        Clean up resources when the Outlook client is deleted.

        Close the HTTP session and log the cleanup action. This method is called automatically when the
        Outlook object is about to be destroyed.

        Notes
        -----
        This method ensures that the HTTP session is properly closed to free up system resources.
        """
        self._logger.info(msg="Cleaning up resources and closing the HTTP session upon exit.")
        self._session.close()

    def auth(self) -> None:
        """
        Authenticate and obtain an access token using the client credentials flow.

        Log the authentication attempt, construct the request to the Microsoft identity platform, and store the
        resulting access token in the configuration for subsequent API requests.

        Notes
        -----
        This method uses the OAuth2 client credentials flow to authenticate with Microsoft Graph API.
        The access token is required for all subsequent API calls.

        Returns
        -------
        None
            This method does not return a value. The access token is stored in the configuration.
        """
        self._logger.info(msg="Authenticating with Microsoft Graph API using client credentials flow.")

        # Request headers
        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        # Authorization URL
        url_auth = f"https://login.microsoftonline.com/{self._configuration.tenant_id}/oauth2/v2.0/token"

        # Request body
        body = {
            "grant_type": "client_credentials",
            "client_id": self._configuration.client_id,
            "client_secret": self._configuration.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }

        # Request
        response = self._session.post(url=url_auth, data=body, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Return valid response
        if response.status_code == 200:
            self._configuration.token = json.loads(response.content.decode("utf-8"))["access_token"]

    def _export_to_json(self, content: bytes, save_as: str | None) -> None:
        """
        Export response content to a JSON file.

        Save the given content to a specified file in JSON format if a file path is provided.

        Parameters
        ----------
        content : bytes
            The content to export, typically the response content from an API call.
        save_as : str or None
            The file path where the JSON content will be saved. If None, do not save the content.

        Returns
        -------
        None
        """
        if save_as is not None:
            self._logger.info(msg="Exports response to JSON file.")
            with open(file=save_as, mode="wb") as file:
                file.write(content)

    def _handle_response(
        self, response: requests.Response, model: Type[BaseModel], rtype: str = "scalar"
    ) -> dict | list[dict]:
        """
        Deserialize and validate the JSON content from an API response.

        Process the response from an API request and convert the JSON content into a Pydantic BaseModel or a list of
        BaseModel instances, depending on the specified response type.

        Parameters
        ----------
        response : requests.Response
            The response object returned by the API request.
        model : Type[BaseModel]
            The Pydantic BaseModel class to use for deserialization and validation.
        rtype : str, default="scalar"
            Specify the type of response to handle. Use "scalar" for a single record and "list" for a list of records.

        Returns
        -------
        dict or list of dict
            The deserialized content as a dictionary (for scalar) or a list of dictionaries (for list).
        """
        if rtype.lower() == "scalar":
            # Deserialize json (scalar values)
            content_raw = response.json()
            # Pydantic v1 validation
            validated = model(**content_raw)
            # Convert to dict
            return validated.dict()

        # Deserialize json
        content_raw = response.json()["value"]
        # Pydantic v1 validation
        validated_list = parse_obj_as(list[model], content_raw)
        # return [dict(data) for data in parse_obj_as(list[model], content_raw)]
        # Convert to a list of dicts
        return [item.dict() for item in validated_list]

    def change_client_email(self, email: str) -> None:
        """
        Change the current client email address.

        Update the email address of the client that the Outlook instance will interact with for subsequent operations.

        Parameters
        ----------
        email : str
            The new client email address to set.

        Returns
        -------
        None
        """
        self._logger.info(msg="Changing the current client email address for subsequent Outlook API operations.")
        self._logger.info(msg=email)

        self._configuration.client_email = email

    def change_folder(self, id: str) -> None:
        """
        Change the current mail folder to be accessed.

        Update the mail folder that the Outlook client will interact with for subsequent operations.
        By default, the folder is set to "Inbox".

        Parameters
        ----------
        id : str
            The name or ID of the folder to access. Refer to Microsoft Graph API documentation for
            valid folder names and IDs (e.g., "Inbox", "SentItems", "Drafts", or a folder ID).

        Returns
        -------
        None
            This method does not return a value. The selected folder is updated in the client configuration.

        Notes
        -----
        See Also
        --------
        https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0
        """
        self._logger.info(msg="Changing the current mail folder being accessed by the Outlook client.")
        self._logger.info(msg=id)

        self._configuration.client_folder = id

    def list_folders(self, save_as: str | None = None) -> Response:
        """
        Retrieve a list of mail folders for the authenticated user.

        Log the retrieval action, construct the request to the Microsoft Graph API, and return the list of mail folders.
        Optionally, save the response to a JSON file if a file path is provided.

        Parameters
        ----------
        save_as : str, optional
            File path to save the JSON response. If None, do not save the response.

        Returns
        -------
        Response
            A Response dataclass instance containing the status code and the list of folders.
            - status_code (int): The HTTP status code of the request.
            - content (list[BaseModel] or None): A list of deserialized folder objects if the request is successful,
              otherwise None.

        Notes
        -----
        Refer to Microsoft Graph API documentation for valid folder names and IDs.
        """
        self._logger.info(msg="Retrieving the list of mail folders for the authenticated user.")

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        client_email = self._configuration.client_email
        client_folder = self._configuration.client_folder

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json"}

        # Request query
        client_folder = "" if client_folder.lower() == "root" else f"{client_folder}/childFolders"
        url_query = fr"https://{api_domain}/{api_version}/users/{client_email}/mailFolders/{client_folder}"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in ListFolders.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list), "$top": 100, "includeHiddenFolders": True}

        # Request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json
            content = self._handle_response(response=response, model=ListFolders, rtype="list")

        return self.Response(status_code=response.status_code, content=content)

    def list_messages(self, filter: str, save_as: str | None = None) -> Response:
        """
        Retrieve the top 100 messages from the specified folder, filtered by a given condition.

        Apply the provided filter to fetch up to 100 email messages from the current mail folder.
        Optionally, save the response to a JSON file.

        Parameters
        ----------
        filter : str
            Filter string to apply to the messages. Examples:
                - "isRead ne true" (unread messages)
                - "subject eq 'Meeting Reminder'" (messages with a specific subject)
                - "startswith(subject, 'Invoice')" (subjects starting with 'Invoice')
                - "from/emailAddress/address eq 'no-reply@example.com'" (from a specific email address)
                - "receivedDateTime le 2025-04-01T00:00:00Z" (received before a specific date)
        save_as : str, optional
            File path to save the JSON response. If None, do not save the response.

        Returns
        -------
        Response
            Response dataclass containing the status code and the list of messages.
            - status_code (int): HTTP status code of the request.
            - content (list[BaseModel] or None): List of deserialized message objects if successful, otherwise None.

        Notes
        -----
        Refer to Microsoft Graph API documentation for valid filter syntax and folder names.
        """
        self._logger.info(msg="Retrieving the top 100 messages from the specified folder using the provided filter criteria.")
        self._logger.info(msg=filter)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        client_email = self._configuration.client_email
        client_folder = self._configuration.client_folder

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/users/{client_email}/mailFolders/{client_folder}/messages"

        # Query parameters
        # filter examples
        #   new emails: ?$filter=isRead ne true
        #   with subject AccReview_Processor_ServiceHub_NAM: ?$filter=subject eq 'AccReview_Processor_ServiceHub_NAM'
        #   subject starts with AccReview_: "?$filter=startswith(subject, 'AccReview_')""
        #   from address: "from/emailAddress/address eq 'no-reply@eusmtp.ariba.com'"
        #   from name: "from/emailAddress/name eq 'ASPEN Notification'"
        #   date less or equal: "receivedDateTime le 2021-12-01T00:00:00Z"
        #   extra: # &$count=true
        # Pydantic v1
        alias_list = [field.alias for field in ListMessages.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list), "$filter": filter, "$top": 100}

        # Request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json
            content = self._handle_response(response=response, model=ListMessages, rtype="list")

        return self.Response(status_code=response.status_code, content=content)

    def move_message(self, id: str, to: str, save_as: str | None = None) -> Response:
        """
        Move a message from the current folder to another specified folder.

        Move an email message identified by its unique ID from the current folder to a destination folder specified by
        its ID.

        Parameters
        ----------
        id : str
            Specify the unique identifier of the message to move.
        to : str
            Specify the unique identifier of the destination folder (RestID).
        save_as : str, optional
            Specify the file path to save the JSON response. If None, do not save the response.

        Returns
        -------
        Response
            Return a Response dataclass instance containing the status code and any relevant content.
            - status_code (int): The HTTP status code of the move request.
            - content (BaseModel or None): The deserialized content of the response if the move is successful,
              otherwise None.

        Notes
        -----
        Refer to Microsoft Graph API documentation for valid folder and message IDs.
        """
        self._logger.info(msg="Moving a message from the current folder to the specified destination folder.")
        self._logger.info(msg=id)
        self._logger.info(msg=to)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        client_email = self._configuration.client_email
        client_folder = self._configuration.client_folder

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json"}

        # Request query
        url_query = rf"https://{api_domain}/{api_version}/users/{client_email}/mailFolders/{client_folder}/messages/{id}/move"

        # Pydantic v1
        alias_list = [field.alias for field in MoveMessage.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Body
        body = {"DestinationId": to}

        # Request query
        response = self._session.post(url=url_query, headers=headers, params=params, json=body, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 201:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=MoveMessage, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def delete_message(self, id: str) -> Response:
        """
        Delete a message from the current folder.

        Remove the specified email message from the currently selected folder.

        Parameters
        ----------
        id : str
            Unique identifier of the message to delete.

        Returns
        -------
        Response
            Response dataclass instance containing the status code and any relevant content.
            - status_code (int): HTTP status code of the deletion request.
            - content (None): Content is None when the deletion is successful.

        Notes
        -----
        This method does not return the deleted message content.
        """
        self._logger.info(msg="Deleting a message from the current folder in progress.")
        self._logger.info(msg=id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        client_email = self._configuration.client_email
        client_folder = self._configuration.client_folder

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/users/{client_email}/mailFolders/{client_folder}/messages/{id}"

        # Request query
        response = self._session.delete(url=url_query, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 204:
            self._logger.info(msg="Request successful")

        return self.Response(status_code=response.status_code, content=content)

    def download_message_attachment(self, id: str, path: str, index: bool = False) -> Response:
        """
        Download attachments from an email message and save them to a local directory.

        Retrieve all attachments from the specified email message and write them to the given directory path.
        If `index` is True, append an index to each file name to avoid overwriting files with the same name.

        Parameters
        ----------
        id : str
            Unique identifier of the email message from which to download attachments.
        path : str
            Local directory path where the attachments will be saved.
        index : bool, optional
            If True, append an index to each file name to prevent overwriting (default is False).

        Returns
        -------
        Response
            Response dataclass instance containing the HTTP status code and any relevant content.

        Notes
        -----
        Invalid characters in file names are replaced with underscores. Only attachments with a "contentBytes"
        field are downloaded; others are skipped with an error logged.
        """
        self._logger.info(msg="Initiating the process of downloading attachments from the specified email message.")
        self._logger.info(msg=id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        client_email = self._configuration.client_email
        client_folder = self._configuration.client_folder

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/users/{client_email}/mailFolders/{client_folder}/messages/{id}/attachments"

        # Request query
        response = self._session.get(url=url_query, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        counter = 0
        if response.status_code == 200:
            for row in response.json()["value"]:
                if "contentBytes" in row:
                    file_content = base64.b64decode(row["contentBytes"])

                    if index:
                        counter = counter + 1
                        filename_lst = row["name"].rsplit(".", 1)
                        filename_ext = "." + filename_lst[1] if len(filename_lst) > 1 else None
                        filename = f"{filename_lst[0]}_{str(counter)}{filename_ext}"
                    else:
                        filename = row["name"]

                    # Create file while removing invalid characters
                    filename = re.sub(pattern=r"[^a-zA-Z0-9.]+", repl="_", string=filename)
                    self._logger.info(msg=filename)
                    open(file=os.path.join(path, filename), mode="wb").write(file_content)
                else:
                    self._logger.error(msg="Invalid attachment found")

        return self.Response(status_code=response.status_code, content=content)

# eof
