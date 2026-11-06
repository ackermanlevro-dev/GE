"""
Device-code sign-in to SharePoint Online.

Uses a Microsoft first-party public client app ID (the same one PnP
PowerShell's `-DeviceLogin` uses under the hood), so no Azure AD app
registration is needed. You sign in once with your GE work account in a
browser; after that a cached refresh token keeps you signed in until it
expires or is revoked.

If device-code sign-in fails with an error mentioning admin consent
(AADSTS65001 / AADSTS70011), your tenant may not have that app
pre-approved. Try the alternate client_id in config.json:
    "31359c7f-bd7e-475c-86db-fdb8c937548e"   (PnP Management Shell)
If neither works, ask your Azure AD admin to register an app with
delegated Sites.Selected (or SharePoint "AllSites.Write") permission
scoped to this one list/site.
"""

import json
import os

import msal

CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".msal_cache.json")


def _load_config():
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            f.write(cache.serialize())
        try:
            os.chmod(CACHE_FILE, 0o600)
        except OSError:
            pass


def get_access_token():
    """Returns a bearer token for the SharePoint site's REST API, using a
    cached refresh token when possible and falling back to an interactive
    device-code prompt otherwise."""
    config = _load_config()
    site_url = config["site_url"]
    resource = site_url.split("/sites/")[0].split("/teams/")[0]
    scopes = [f"{resource}/.default"]

    cache = _load_cache()
    app = msal.PublicClientApplication(
        client_id=config["client_id"],
        authority=config["authority"],
        token_cache=cache,
    )

    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to start device-code flow: {flow}")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(
            f"Sign-in failed: {result.get('error')} - {result.get('error_description')}"
        )
    return result["access_token"]
