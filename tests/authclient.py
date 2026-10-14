"""Test helper: a Flask test client that signs in and adds the CSRF token to every POST."""
import re
from flask.testing import FlaskClient


class AuthClient(FlaskClient):
    def open(self, *args, **kwargs):
        if not getattr(self, "_signed_in", False):
            self._signed_in = True
            page = super().open("/login").data.decode()
            tok = re.search(r'name="csrf_token" value="([^"]+)"', page).group(1)
            super().open("/login", method="POST", data={"username": "admin", "password": "admin", "csrf_token": tok})
        if kwargs.get("method", "GET").upper() == "POST" or (args and False):
            tok = self._csrf()
            if kwargs.get("json") is not None:
                kwargs.setdefault("headers", {})["X-CSRF-Token"] = tok
            else:
                data = dict(kwargs.get("data") or {})
                data["csrf_token"] = tok
                kwargs["data"] = data
        return super().open(*args, **kwargs)

    def post(self, *a, **k):
        k["method"] = "POST"
        return self.open(*a, **k)

    def _csrf(self):
        page = super().open("/help").data.decode()
        return re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
