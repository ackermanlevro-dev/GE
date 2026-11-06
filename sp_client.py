"""Thin wrapper around the SharePoint classic REST API for list CRUD."""

import json
import time
import urllib.parse

import requests

from auth import get_access_token

DIGEST_LIFETIME_SECONDS = 25 * 60  # SharePoint digests are valid ~30 min; refresh a bit early


class ConflictError(Exception):
    """Raised when an item changed on the server since it was last read
    (i.e. someone else edited the record concurrently)."""


class SharePointClient:
    def __init__(self, site_url, list_title):
        self.site_url = site_url.rstrip("/")
        self.list_title = list_title
        self._token = get_access_token()
        self._entity_type = None
        self._digest = None
        self._digest_fetched_at = 0
        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {self._token}",
                "Accept": "application/json;odata=verbose",
            }
        )

    def _url(self, path):
        return f"{self.site_url}/_api/{path}"

    def _request(self, method, path, **kwargs):
        url = self._url(path)
        resp = None
        for attempt in range(5):
            resp = self._session.request(method, url, **kwargs)
            if resp.status_code in (429, 503):
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                time.sleep(wait)
                continue
            return resp
        return resp

    def get_fields(self):
        """Returns a list of dicts: Title, InternalName, TypeAsString for
        the list's non-hidden fields."""
        list_part = urllib.parse.quote(self.list_title, safe="")
        resp = self._request(
            "GET",
            f"web/lists/getbytitle('{list_part}')/fields"
            "?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField"
            "&$filter=Hidden eq false",
        )
        resp.raise_for_status()
        results = resp.json()["d"]["results"]
        return [
            {
                "Title": f["Title"],
                "InternalName": f["InternalName"],
                "TypeAsString": f["TypeAsString"],
                "ReadOnlyField": f["ReadOnlyField"],
            }
            for f in results
        ]

    def _get_entity_type(self):
        if self._entity_type is None:
            list_part = urllib.parse.quote(self.list_title, safe="")
            resp = self._request(
                "GET",
                f"web/lists/getbytitle('{list_part}')?$select=ListItemEntityTypeFullName",
            )
            resp.raise_for_status()
            self._entity_type = resp.json()["d"]["ListItemEntityTypeFullName"]
        return self._entity_type

    def _get_digest(self):
        now = time.time()
        if self._digest is None or (now - self._digest_fetched_at) > DIGEST_LIFETIME_SECONDS:
            resp = self._request("POST", "contextinfo")
            resp.raise_for_status()
            self._digest = resp.json()["d"]["GetContextWebInformation"]["FormDigestValue"]
            self._digest_fetched_at = now
        return self._digest

    def find_item_by_key(self, key_internal_name, key_value, select_internal_names=()):
        """Looks up a single item by an exact match on key_internal_name.

        Returns a dict with at least {"Id": ..., "__etag": ...} plus the
        current value of every field in select_internal_names (so callers
        can both validate uniqueness and capture a pre-change snapshot in
        one round trip), or None if there's no match. Raises ValueError if
        more than one record matches, since the key should be unique.
        """
        list_part = urllib.parse.quote(self.list_title, safe="")
        escaped_value = str(key_value).replace("'", "''")
        filter_expr = f"{key_internal_name} eq '{escaped_value}'"
        select_fields = ["Id"] + list(select_internal_names)
        resp = self._request(
            "GET",
            f"web/lists/getbytitle('{list_part}')/items"
            f"?$select={','.join(select_fields)}"
            f"&$filter={urllib.parse.quote(filter_expr, safe='')}",
        )
        resp.raise_for_status()
        results = resp.json()["d"]["results"]
        if len(results) == 0:
            return None
        if len(results) > 1:
            raise ValueError(
                f"Key value {key_value!r} matched {len(results)} items; expected a unique match."
            )
        item = results[0]
        item["__etag"] = item.get("__metadata", {}).get("etag")
        return item

    def update_item(self, item_id, fields, etag=None):
        """fields: dict of {InternalName: value} to update on the item.

        If etag is given, the update uses it as an optimistic-concurrency
        check: if the record was modified on the server since etag was
        read, this raises ConflictError instead of overwriting the change.
        """
        list_part = urllib.parse.quote(self.list_title, safe="")
        entity_type = self._get_entity_type()
        digest = self._get_digest()
        body = {"__metadata": {"type": entity_type}}
        body.update(fields)
        resp = self._request(
            "POST",
            f"web/lists/getbytitle('{list_part}')/items({item_id})",
            headers={
                "X-RequestDigest": digest,
                "X-HTTP-Method": "MERGE",
                "IF-MATCH": etag or "*",
                "Content-Type": "application/json;odata=verbose",
            },
            data=json.dumps(body),
        )
        if resp.status_code == 412:
            raise ConflictError(
                "Record was modified by someone else since it was read; update aborted "
                "to avoid overwriting their change. Re-run this row after reviewing it."
            )
        if resp.status_code not in (200, 204):
            raise RuntimeError(f"Update failed ({resp.status_code}): {resp.text[:500]}")
