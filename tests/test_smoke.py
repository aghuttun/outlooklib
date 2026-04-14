"""Smoke tests for the Python bindings public surface.

These tests verify that the main symbols exposed by the extension module are
available and importable in a local environment.
"""

import unittest

import outlooklib


class SmokeTests(unittest.TestCase):
    """Validate minimal public API symbols exposed by the package."""

    def test_public_api_symbols_exist(self):
        """Ensure Python consumers can access core `Outlook` and `Response` classes."""
        self.assertTrue(hasattr(outlooklib, "Outlook"))
        self.assertTrue(hasattr(outlooklib, "Response"))


if __name__ == "__main__":
    unittest.main()
