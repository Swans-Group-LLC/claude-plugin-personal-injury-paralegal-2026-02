#!/usr/bin/env python3
"""Upload a document to Clio Manage via Make.com webhook.

Outputs JSON to stdout on success, errors to stderr with exit code 1.

Usage:
    python3 clio-manage-upload-document.py --matter-id <ID> --file-path <PATH> --document-name <NAME>
"""

import argparse
import json
import mimetypes
import ssl
import sys
import uuid
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

# macOS Python often lacks SSL certs; use certifi if available, else unverified
try:
    import certifi
    _ssl_ctx = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    _ssl_ctx = ssl.create_default_context()
    _ssl_ctx.check_hostname = False
    _ssl_ctx.verify_mode = ssl.CERT_NONE

WEBHOOK_URL = "XXX"


def error(msg):
    print(json.dumps({"status": "error", "message": msg}), file=sys.stderr)
    sys.exit(1)


def build_multipart(fields, files):
    """Build multipart/form-data body using only stdlib."""
    boundary = uuid.uuid4().hex
    parts = []

    for key, value in fields.items():
        parts.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{key}\"\r\n\r\n{value}".encode())

    for key, (filename, data, content_type) in files.items():
        parts.append(
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"{key}\"; filename=\"{filename}\"\r\n"
            f"Content-Type: {content_type}\r\n\r\n".encode() + data
        )

    parts.append(f"--{boundary}--\r\n".encode())
    body = b"\r\n".join(parts)
    return body, f"multipart/form-data; boundary={boundary}"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--matter-id", required=True)
    parser.add_argument("--file-path", required=True)
    parser.add_argument("--document-name", required=True)
    args = parser.parse_args()

    file_path = Path(args.file_path)
    if not file_path.is_file():
        error(f"File not found: {file_path}")

    content_type = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
    file_data = file_path.read_bytes()

    body, content_type_header = build_multipart(
        fields={"matter_id": args.matter_id, "document_name": args.document_name},
        files={"file": (file_path.name, file_data, content_type)},
    )

    req = Request(WEBHOOK_URL, data=body, headers={"Content-Type": content_type_header}, method="POST")

    try:
        with urlopen(req, context=_ssl_ctx) as resp:
            response_body = resp.read().decode()
    except HTTPError as e:
        error(f"HTTP {e.code} from webhook")
    except URLError as e:
        error(f"Request failed: {e.reason}")

    # Pass through the response from Make.com (expected to contain document_id)
    print(response_body)


if __name__ == "__main__":
    main()
