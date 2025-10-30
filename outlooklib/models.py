"""
Define Pydantic data structures for Outlook entities.

This module provides Pydantic models for representing Outlook folders and messages,
enabling structured data validation and serialization for use with Outlook-related APIs.

Classes
-------
ListFolders
    Represent an Outlook folder with its identifier and display name.

ListMessages
    Represent an Outlook message with sender, subject, status, and metadata fields.

MoveMessage
    Represent the result of moving an Outlook message, including its new identifier and change key.
"""

import datetime
from pydantic import BaseModel, Field


class ListFolders(BaseModel):
    """
    Represent an Outlook folder with its identifier and display name.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the Outlook folder.
    display_name : str
        Specify the display name of the Outlook folder.

    Attributes
    ----------
    id : str
        The unique identifier of the folder.
    display_name : str
        The display name of the folder.
    """

    id: str = Field(alias="id")
    display_name: str = Field(alias="displayName")


class ListMessages(BaseModel):
    """
    Represent an Outlook message with sender, subject, status, and metadata fields.

    Define the structure for an Outlook message, including sender information,
    subject, read status, attachment presence, importance, flag status, and web link.

    Attributes
    ----------
    id : str
        Specify the unique identifier of the Outlook message.
    sender : dict
        Provide the sender information as a dictionary.
    received_date_time : datetime.datetime
        Indicate the date and time the message was received.
    subject : str
        Specify the subject of the message.
    is_read : bool
        Indicate whether the message has been read.
    has_attachments : bool
        Indicate whether the message contains attachments.
    importance : str
        Specify the importance level of the message.
    flag : dict
        Provide the flag status as a dictionary.
    web_link : str
        Specify the web link to access the message.
    """

    id: str = Field(alias="id")
    sender: dict = Field(alias="sender")
    # cc_recipients: list | dict = Field(alias="ccRecipients")
    # bcc_recipients: list | dict = Field(alias="bccRecipients")
    received_date_time: datetime.datetime = Field(alias="receivedDateTime")
    subject: str = Field(alias="subject")
    is_read: bool = Field(alias="isRead")
    has_attachments: bool = Field(alias="hasAttachments")
    importance: str = Field(alias="importance")
    flag: dict = Field(alias="flag")
    web_link: str = Field(alias="webLink")


class MoveMessage(BaseModel):
    """
    Represent the result of moving an Outlook message.

    Capture the new identifier and change key for a message after it has been moved.

    Parameters
    ----------
    id : str
        Set the unique identifier of the moved message.
    change_key : str
        Set the change key representing the current version of the message.

    Attributes
    ----------
    id : str
        The unique identifier of the moved message.
    change_key : str
        The change key for the current version of the message.
    """

    id: str = Field(alias="id")
    change_key: str = Field(alias="changeKey")


# eom
