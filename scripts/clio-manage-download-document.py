#!/usr/bin/env python3
"""Download a document from Clio Manage via Make.com webhook.

Outputs JSON to stdout on success, errors to stderr with exit code 1.

Usage:
    python3 clio-manage-download-document.py --document-id <ID> --output-path <PATH>
"""

import argparse
import json
import ssl
import sys
from pathlib import Path
from urllib.parse import urlencode
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--document-id", required=True)
    parser.add_argument("--output-path", required=True)
    args = parser.parse_args()

    output_path = Path(args.output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    url = f"{WEBHOOK_URL}?{urlencode({'document_id': args.document_id})}"
    req = Request(url, data=b"", method="POST")

    try:
        with urlopen(req, context=_ssl_ctx) as resp:
            data = resp.read()
    except HTTPError as e:
        error(f"HTTP {e.code} from webhook")
    except URLError as e:
        error(f"Request failed: {e.reason}")

    if not data:
        error("Downloaded file is empty")

    output_path.write_bytes(data)

    print(json.dumps({
        "status": "ok",
        "path": str(output_path),
        "size_bytes": len(data),
    }))


if __name__ == "__main__":
    main()
