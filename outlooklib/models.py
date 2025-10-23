"""
Pydantic data structures for Outlook entities.
"""

import datetime
from pydantic import BaseModel, Field


class ListFolders(BaseModel):
    """Pydantic data structure for move_message()."""

    id: str = Field(alias="id")
    display_name: str = Field(alias="displayName")


class ListMessages(BaseModel):
    """Pydantic data structure for list_messages()."""

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
    """Pydantic data structure for move_message()."""

    id: str = Field(alias="id")
    change_key: str = Field(alias="changeKey")


# eom
