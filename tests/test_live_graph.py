"""Optional live integration tests against Microsoft Graph.

Tests in this module are skipped unless Outlook Graph credentials are available
through environment variables.
"""

import os
import unittest

import outlooklib


class LiveGraphTests(unittest.TestCase):
    """Run end-to-end checks for the Outlook client against a live mailbox."""

    @staticmethod
    def _get_env(var_name):
        """Read one environment variable used by integration tests."""
        return os.getenv(var_name)

    def test_outlook_live_smoke(self):
        """Validate list and folder-switch flows against Microsoft Graph."""
        client_id = self._get_env("OUTLOOKLIB_CLIENT_ID")
        tenant_id = self._get_env("OUTLOOKLIB_TENANT_ID")
        client_secret = self._get_env("OUTLOOKLIB_CLIENT_SECRET")
        client_email = self._get_env("OUTLOOKLIB_CLIENT_EMAIL")

        if not all([client_id, tenant_id, client_secret, client_email]):
            self.skipTest("Set OUTLOOKLIB_CLIENT_ID, OUTLOOKLIB_TENANT_ID, OUTLOOKLIB_CLIENT_SECRET, OUTLOOKLIB_CLIENT_EMAIL")

        outlook = outlooklib.Outlook(
            client_id=client_id,
            tenant_id=tenant_id,
            client_secret=client_secret,
            client_email=client_email,
            client_folder="Inbox",
        )

        folders_response = outlook.list_folders()
        self.assertEqual(folders_response.status_code, 200)
        self.assertIsInstance(folders_response.content, list)

        messages_response = outlook.list_messages(filter="isRead ne true")
        self.assertEqual(messages_response.status_code, 200)
        self.assertIsInstance(messages_response.content, list)

        # Validate folder switching path used by existing Python API.
        outlook.change_folder(id="root")
        root_folders_response = outlook.list_folders()
        self.assertEqual(root_folders_response.status_code, 200)
        self.assertIsInstance(root_folders_response.content, list)


if __name__ == "__main__":
    unittest.main()
