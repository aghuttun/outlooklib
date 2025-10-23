# Databricks notebook source
"""
This library provides access to Outlook functionalities.
Its features are designed to remain generic and should not be modified to meet the specific needs of individual
projects.
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
    Outlook client to interact with Outlook via Microsoft Graph API.
    """

    @dataclasses.dataclass
    class Configuration(object):
        """Configuration dataclass for Outlook client."""

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
        """Response dataclass for Outlook client methods."""

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

        Args:
            client_id (str): The Azure client ID used for authentication.
            tenant_id (str): The Azure tenant ID associated with the client.
            client_secret (str): The secret key for the Azure client.
            sp_domain (str): The SharePoint domain.
            client_email (str): Client email account.
            client_folder (str): Client folder. Defaults to "Inbox".
            custom_logger (logging.Logger, optional): Logger instance to use. If None, a default logger is created.
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
        Clean the house at the exit.
        """
        self._logger.info(msg="Cleans the house at the exit")
        self._session.close()

    def auth(self) -> None:
        """
        Authentication.

        This method performs the authentication process to obtain an access token using the client credentials flow. The
        token is stored in the Configuration dataclass for subsequent API requests.
        """
        self._logger.info(msg="Authentication")

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

        This method takes the content to be exported and saves it to a specified file in JSON format.
        If the `save_as` parameter is provided, the content will be written to that file.

        Args:
            content (bytes): The content to be exported, typically the response content from an API call.
            save_as (str): The file path where the JSON content will be saved. If None, the content will not be saved.
        """
        if save_as is not None:
            self._logger.info(msg="Exports response to JSON file.")
            with open(file=save_as, mode="wb") as file:
                file.write(content)

    def _handle_response(
        self, response: requests.Response, model: Type[BaseModel], rtype: str = "scalar"
    ) -> dict | list[dict]:
        """
        Handle and deserializes the JSON content from an API response.

        This method processes the response from an API request and deserializes the JSON content into a Pydantic
        BaseModel or a list of BaseModel instances, depending on the response type.

        Args:
            response (requests.Response): The response object from the API request.
            model (Type[BaseModel]): The Pydantic BaseModel class to use for deserialization and validation.
            rtype (str, optional): The type of response to handle. Use "scalar" for a single record and "list" for a
              list of records. Defaults to "scalar".

        Returns:
            dict or list[dict]: The deserialized content as a dictionary (scalar) or a list of dictionaries (list).
        """
        if rtype.lower() == "scalar":
            # Deserialize json (scalar values)
            content_raw = response.json()
            # Pydantic v1 validation
            validated = model(**content_raw)
            # Convert to dict
            return validated.dict()
        else:
            # Deserialize json
            content_raw = response.json()["value"]
            # Pydantic v1 validation
            validated_list = parse_obj_as(list[model], content_raw)
            # return [dict(data) for data in parse_obj_as(list[model], content_raw)]
            # Convert to a list of dicts
            return [item.dict() for item in validated_list]

    def change_client_email(self, email: str) -> None:
        """
        Change the current client email to be accessed.

        This method updates the email address of the client that the Outlook instance will interact with for subsequent
        operations.

        Args:
            email (str): The new client email address to be set.
        """
        self._logger.info(msg="Change the current client email to be accessed")
        self._logger.info(msg=email)

        self._configuration.client_email = email

    def change_folder(self, id: str) -> None:
        """
        Change the current folder to be accessed. By default, it is set to "Inbox".
        This method updates the folder that the client will interact with for subsequent operations.

        Note: https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0

        Args:
            id (str): The name or ID of the folder to be accessed. Refer to the Microsoft Graph API documentation for
              valid folder names and IDs.
              Example folder names: "Inbox", "SentItems", "Drafts", etc.
              Example folder ID: "AAMkAGI2T..."
        """
        self._logger.info(msg="Change the current folder to be accessed")
        self._logger.info(msg=id)

        self._configuration.client_folder = id

    def list_folders(self, save_as: str | None = None) -> Response:
        """
        Retrieve a list of mail folders for the authenticated user.

        This method fetches the mail folders for the authenticated user's email account. The results can be saved to a
        JSON file if the `save_as` parameter is provided.

        Args:
            save_as (str, optional): The file path where the JSON response will be saved. If None, the response will not
              be saved.

        Returns:
            Response: A Response dataclass instance containing the status code and the list of folders.
                      - status_code (int): The HTTP status code of the request.
                      - content (list[BaseModel] | None): A list of deserialized folder objects if the request is
                        successful, otherwise None.
        """
        self._logger.info(msg="Get list of folders")

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

        This method fetches up to 100 email messages from the current folder, applying an filter to narrow down the
        results. The results can be saved to a JSON file if the `save_as` parameter is provided.

        Args:
            filter (str): A filter string to apply to the messages.
              Examples:
              - "isRead ne true" (unread messages)
              - "subject eq 'Meeting Reminder'" (messages with a specific subject)
              - "startswith(subject, 'Invoice')" (messages with subjects starting with 'Invoice')
              - "from/emailAddress/address eq 'no-reply@example.com'" (messages from a specific email address)
              - "receivedDateTime le 2025-04-01T00:00:00Z" (messages received before a specific date)
            save_as (str, optional): The file path where the JSON response will be saved. If None, the response will not
              be saved.

        Returns:
            Response: A Response dataclass instance containing the status code and the list of messages.
              - status_code (int): The HTTP status code of the request.
              - content (list[BaseModel] | None): A list of deserialized message objects if the request is successful,
                otherwise None.
        """
        self._logger.info(msg="Returns the (top 100) messages found in the folder")
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

        This method relocates an email message identified by its unique ID from the current folder to a destination
        folder specified by its ID.

        Args:
            id (str): The unique identifier of the message to be moved.
            to (str): The unique identifier of the destination folder (RestID).
            save_as (str, optional): The file path where the JSON response will be saved. If None, the response will
              not be saved.

        Returns:
            Response: A Response dataclass instance containing the status code and any relevant content.
              - status_code (int): The HTTP status code of the move request.
              - content (BaseModel | None): The deserialized content of the response if the move is successful,
                otherwise None.
        """
        self._logger.info(msg="Moves a message from the folder to another folder")
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

        This method removes a specified email message from the currently selected folder.

        Args:
            id (str): The unique identifier of the message to be deleted.

        Returns:
            Response: A Response dataclass instance containing the status code and any relevant content.
              - status_code (int): The HTTP status code of the deletion request.
              - content (None): Content is None when the deletion is successful.
        """
        self._logger.info(msg="Deletes a message from the current folder")
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
        Download attachments from an email message.

        This method retrieves and saves attachments from the specified email message to a local file path.
        If the `index` parameter is set to True, each saved file will have an appended index in its name.

        Args:
            id (str): The unique identifier of the email message from which attachments will be downloaded.
            path (str): The local directory path where the downloaded attachments will be saved.
            index (bool, optional): If True, an index is added to the file name to prevent overwriting. Defaults to
              False.

        Returns:
            Response: A Response dataclass instance containing the status code and any relevant content.
        """
        self._logger.info(msg="Downloads attachments from an email message")
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

# eom
