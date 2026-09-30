"""Keep the Calendar GPT listing consistent with its published API provider."""
import json
from pathlib import Path
import re
import unittest
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]


class CalendarDocumentation(unittest.TestCase):
    def test_calendar_description_matches_graph_provider(self):
        api = json.loads((ROOT / "openapi-json/ms-graph.json").read_text())
        self.assertTrue(api["openapi"].startswith("3."))
        self.assertIn("/calendars", api["paths"])
        self.assertEqual(
            {urlparse(server["url"]).hostname for server in api["servers"]},
            {"graph.microsoft.com"},
        )
        readme = (ROOT / "README.md").read_text()
        listing = re.search(
            r"\*\*GPT Calendar and Email Scheduler\*\*(.*?)(?=\n\d+\. |\Z)",
            readme,
            re.DOTALL,
        )
        self.assertIsNotNone(listing, "Calendar GPT must be documented")
        self.assertIn("Microsoft", listing.group(1))
        self.assertNotIn("Google Calendar", listing.group(1))
        self.assertIn("openapi-json/ms-graph.json", listing.group(1))

    def test_calendar_oauth_uses_same_provider(self):
        api = json.loads((ROOT / "openapi-json/ms-graph.json").read_text())
        schemes = api["components"]["securitySchemes"]
        flow = schemes["OAuth2"]["flows"]["authorizationCode"]
        for field in ("authorizationUrl", "tokenUrl"):
            url = urlparse(flow[field])
            self.assertEqual(url.scheme, "https")
            self.assertEqual(url.hostname, "login.microsoftonline.com")
        for scope in flow["scopes"]:
            self.assertEqual(urlparse(scope).hostname, "graph.microsoft.com")


if __name__ == "__main__":
    unittest.main()
