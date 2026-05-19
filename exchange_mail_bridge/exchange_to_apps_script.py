import argparse
import base64
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
    FileAttachment,
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
    include_attachments = parse_bool(env("INCLUDE_ATTACHMENTS"), default=False)
    include_body_html = parse_bool(env("INCLUDE_BODY_HTML"), default=True)
    max_attachment_bytes = int(env("MAX_ATTACHMENT_BYTES", "2500000"))
    max_body_html_chars = int(env("MAX_BODY_HTML_CHARS", "200000"))
    items = (
        account.inbox.filter(datetime_received__gte=cutoff)
        .order_by("-datetime_received")[:top]
    )
    return [
        serialize_item(
            item,
            include_attachments=include_attachments,
            include_body_html=include_body_html,
            max_attachment_bytes=max_attachment_bytes,
            max_body_html_chars=max_body_html_chars,
        )
        for item in items
    ]


def serialize_item(
    item: Any,
    include_attachments: bool,
    include_body_html: bool,
    max_attachment_bytes: int,
    max_body_html_chars: int,
) -> Dict[str, Any]:
    sender = getattr(item, "sender", None) or getattr(item, "author", None)
    body_preview = extract_body_preview(item)
    body_html = extract_body_html(item, max_body_html_chars) if include_body_html else ""
    body_text_full = extract_body_text(item, limit=50000)
    body_links = extract_body_links(item)
    attachments = serialize_attachments(item, include_attachments, max_attachment_bytes)
    body_image_urls = extract_body_image_urls(item)
    attachment_count = count_attachments(item)

    return {
        "id": str(getattr(item, "id", "") or ""),
        "receivedTime": iso_or_empty(getattr(item, "datetime_received", None)),
        "fromName": str(getattr(sender, "name", "") or ""),
        "fromEmail": str(getattr(sender, "email_address", "") or ""),
        "subject": str(getattr(item, "subject", "") or ""),
        "bodyPreview": body_preview,
        "bodyTextFull": body_text_full,
        "bodyHtml": body_html,
        "bodyLinks": body_links,
        "hasAttachments": bool(getattr(item, "has_attachments", False) or attachments),
        "attachmentCount": attachment_count,
        "attachments": attachments,
        "bodyImageUrls": body_image_urls,
        "conversationId": str(getattr(item, "conversation_id", "") or ""),
        "webLink": "",
    }


def extract_body_preview(item: Any, limit: int = 1000) -> str:
    return extract_body_text(item, limit=limit)


def extract_body_text(item: Any, limit: int = 50000) -> str:
    text_body = getattr(item, "text_body", None)
    if text_body:
        return normalize_text(str(text_body))[:limit]

    body = getattr(item, "body", None)
    if not body:
        return ""

    return html_to_text(str(body))[:limit]


def extract_body_html(item: Any, limit: int) -> str:
    body = str(getattr(item, "body", "") or "")
    if not body:
        return ""
    if "<" not in body and ">" not in body:
        return ""
    return body[:limit]


def html_to_text(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?i)<br\s*/?>", "\n", value)
    value = re.sub(r"(?i)</?(?:p|div|tr|table|tbody|thead|tfoot|ul|ol|li|h[1-6])\b[^>]*>", "\n", value)
    value = re.sub(r"(?i)</?(?:td|th)\b[^>]*>", "\n", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    return normalize_text_lines(unescape(value))


def extract_body_links(item: Any, limit: int = 100) -> List[str]:
    body = str(getattr(item, "body", "") or "")
    if not body:
        return []

    links: List[str] = []
    for match in re.finditer(r"""(?is)<a\b[^>]*\bhref=["']([^"']+)["'][^>]*>(.*?)</a>""", body):
        href = normalize_text(unescape(match.group(1)))
        label = html_to_text(match.group(2))[:200]
        if not href:
            continue
        value = f"{label} | {href}" if label else href
        if value not in links:
            links.append(value[:4000])
        if len(links) >= limit:
            return links
    return links


def extract_body_image_urls(item: Any, limit: int = 20) -> List[str]:
    body = str(getattr(item, "body", "") or "")
    if not body:
        return []

    urls: List[str] = []
    patterns = [
        r"""(?is)<img\b[^>]*\bsrc=["']([^"']+)["']""",
        r"""(?is)\b(?:src|href|background)=["']([^"']+\.(?:png|jpe?g|gif|webp|bmp|svg)(?:\?[^"']*)?)["']""",
        r"""(?is)url\(["']?([^)"']+\.(?:png|jpe?g|gif|webp|bmp|svg)(?:\?[^)"']*)?)["']?\)""",
        r"""(?is)\b(cid:[^"' <>)]+)""",
        r"""(?is)\b(data:image/[^"' <>)]+)""",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, body):
            src = normalize_text(unescape(match.group(1)))
            if src and src not in urls:
                urls.append(src[:2000])
            if len(urls) >= limit:
                return urls
    return urls


def count_attachments(item: Any) -> int:
    attachments = getattr(item, "attachments", None) or []
    try:
        return len(attachments)
    except TypeError:
        return 0


def serialize_attachments(item: Any, include_attachments: bool, max_attachment_bytes: int) -> List[Dict[str, Any]]:
    attachments = getattr(item, "attachments", None) or []
    if not include_attachments:
        return []

    serialized = []
    for attachment in attachments:
        if not isinstance(attachment, FileAttachment):
            continue

        name = str(getattr(attachment, "name", "") or "attachment")
        content_type = str(getattr(attachment, "content_type", "") or "application/octet-stream")
        is_inline = bool(getattr(attachment, "is_inline", False))
        content_id = str(getattr(attachment, "content_id", "") or "")

        if not should_download_attachment(name, content_type, is_inline):
            serialized.append(attachment_metadata(attachment, skipped_reason="unsupported_type"))
            continue

        content = getattr(attachment, "content", b"") or b""
        size = len(content)
        if size <= 0:
            serialized.append(attachment_metadata(attachment, skipped_reason="empty"))
            continue
        if size > max_attachment_bytes:
            serialized.append(attachment_metadata(attachment, skipped_reason="too_large"))
            continue

        serialized.append(
            {
                "name": name,
                "contentType": content_type,
                "size": size,
                "isInline": is_inline,
                "contentId": content_id,
                "base64": base64.b64encode(content).decode("ascii"),
            }
        )
    return serialized


def attachment_metadata(attachment: Any, skipped_reason: str) -> Dict[str, Any]:
    return {
        "name": str(getattr(attachment, "name", "") or "attachment"),
        "contentType": str(getattr(attachment, "content_type", "") or ""),
        "size": int(getattr(attachment, "size", 0) or 0),
        "isInline": bool(getattr(attachment, "is_inline", False)),
        "contentId": str(getattr(attachment, "content_id", "") or ""),
        "skippedReason": skipped_reason,
    }


def should_download_attachment(name: str, content_type: str, is_inline: bool) -> bool:
    allowed_prefixes = [
        prefix.strip().lower()
        for prefix in env("ATTACHMENT_MIME_PREFIXES", "image/").split(",")
        if prefix.strip()
    ]
    lowered_content_type = content_type.lower()
    if any(lowered_content_type.startswith(prefix) for prefix in allowed_prefixes):
        return True

    lowered_name = name.lower()
    image_extensions = (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff")
    return is_inline and lowered_name.endswith(image_extensions)


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\x00", "")).strip()


def normalize_text_lines(value: str) -> str:
    lines = [
        re.sub(r"[ \t\f\v]+", " ", line).strip()
        for line in value.replace("\x00", "").splitlines()
    ]
    return "\n".join(line for line in lines if line)


def iso_or_empty(value: Any) -> str:
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def post_messages(messages: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    message_list = list(messages)
    batch_size = max(1, int(env("POST_BATCH_SIZE", "10")))
    result: Dict[str, Any] = {"ok": True, "appended": 0, "updated": 0, "skipped": 0}

    for index in range(0, len(message_list), batch_size):
        payload = {
            "token": required_env("BRIDGE_TOKEN"),
            "messages": message_list[index:index + batch_size],
        }
        response = requests.post(
            required_env("APPS_SCRIPT_WEBAPP_URL"),
            json=payload,
            timeout=int(env("APPS_SCRIPT_TIMEOUT", "60")),
        )
        response.raise_for_status()
        batch_result = response.json()
        if not batch_result.get("ok"):
            raise SystemExit(f"Apps Script error: {batch_result}")

        for key in ("appended", "updated", "skipped"):
            result[key] += int(batch_result.get(key, 0) or 0)

    return result


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
        print(json.dumps({"messages": redact_attachment_payloads(messages)}, ensure_ascii=False, indent=2))
        return

    result = post_messages(messages)
    print(json.dumps(result, ensure_ascii=False))


def redact_attachment_payloads(messages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    redacted = json.loads(json.dumps(messages))
    for message in redacted:
        for attachment in message.get("attachments", []):
            if "base64" in attachment:
                attachment["base64"] = f"<{len(attachment['base64'])} base64 chars>"
    return redacted


if __name__ == "__main__":
    main()
