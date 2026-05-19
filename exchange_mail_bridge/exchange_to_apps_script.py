import argparse
import json
import os
import re
from datetime import timedelta
from html import unescape
from typing import Any, Dict, Iterable, List

import requests
from exchangelib import (
    BASIC,
    DELEGATE,
    NTLM,
    Account,
    Configuration,
    Credentials,
    EWSDateTime,
    FaultTolerance,
    UTC,
)
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter


DEFAULT_EWS_URL = "https://mail.wackerhagengruppe.de/EWS/Exchange.asmx"


def env(name: str, default: str = "") -> str:
    value = os.environ.get(name)
    if value is None or not value.strip():
        return default
    return value.strip()


def required_env(name: str) -> str:
    value = env(name)
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


def parse_bool(value: str, default: bool = True) -> bool:
    if not value:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def auth_type_from_env() -> str:
    auth_type = env("EWS_AUTH_TYPE", "NTLM").upper()
    if auth_type == "BASIC":
        return BASIC
    if auth_type == "NTLM":
        return NTLM
    raise SystemExit("EWS_AUTH_TYPE must be NTLM or BASIC")


def connect_account() -> Account:
    if not parse_bool(env("EWS_VERIFY_TLS"), default=True):
        BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

    credentials = Credentials(
        username=required_env("EWS_USERNAME"),
        password=required_env("EWS_PASSWORD"),
    )
    config = Configuration(
        service_endpoint=env("EWS_URL", DEFAULT_EWS_URL),
        credentials=credentials,
        auth_type=auth_type_from_env(),
        retry_policy=FaultTolerance(max_wait=60),
    )

    return Account(
        primary_smtp_address=required_env("EWS_EMAIL"),
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )


def fetch_messages(account: Account, lookback_minutes: int, top: int) -> List[Dict[str, Any]]:
    cutoff = EWSDateTime.now(tz=UTC) - timedelta(minutes=lookback_minutes)
    items = (
        account.inbox.filter(datetime_received__gte=cutoff)
        .order_by("-datetime_received")[:top]
    )
    return [serialize_item(item) for item in items]


def serialize_item(item: Any) -> Dict[str, Any]:
    sender = getattr(item, "sender", None) or getattr(item, "author", None)
    body_preview = extract_body_preview(item)

    return {
        "id": str(getattr(item, "id", "") or ""),
        "receivedTime": iso_or_empty(getattr(item, "datetime_received", None)),
        "fromName": str(getattr(sender, "name", "") or ""),
        "fromEmail": str(getattr(sender, "email_address", "") or ""),
        "subject": str(getattr(item, "subject", "") or ""),
        "bodyPreview": body_preview,
        "hasAttachments": bool(getattr(item, "has_attachments", False)),
        "conversationId": str(getattr(item, "conversation_id", "") or ""),
        "webLink": "",
    }


def extract_body_preview(item: Any, limit: int = 1000) -> str:
    text_body = getattr(item, "text_body", None)
    if text_body:
        return normalize_text(str(text_body))[:limit]

    body = getattr(item, "body", None)
    if not body:
        return ""

    return html_to_text(str(body))[:limit]


def html_to_text(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?i)<br\s*/?>", "\n", value)
    value = re.sub(r"(?i)</p\s*>", "\n", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    return normalize_text(unescape(value))


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\x00", "")).strip()


def iso_or_empty(value: Any) -> str:
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def post_messages(messages: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    payload = {
        "token": required_env("BRIDGE_TOKEN"),
        "messages": list(messages),
    }
    response = requests.post(
        required_env("APPS_SCRIPT_WEBAPP_URL"),
        json=payload,
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Poll Exchange EWS and push mail rows to Apps Script.")
    parser.add_argument("--dry-run", action="store_true", help="Print messages instead of posting to Apps Script.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    lookback_minutes = int(env("MAIL_LOOKBACK_MINUTES", "60"))
    top = int(env("MAIL_TOP", "25"))

    account = connect_account()
    messages = fetch_messages(account, lookback_minutes=lookback_minutes, top=top)

    if args.dry_run:
        print(json.dumps({"messages": messages}, ensure_ascii=False, indent=2))
        return

    result = post_messages(messages)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
