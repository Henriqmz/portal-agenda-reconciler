from __future__ import annotations

import base64
import email.utils
import hashlib
import hmac
import struct
import time
import urllib.request


def measure_clock_offset(reference_urls: tuple[str, ...] = ("https://www.google.com",), timeout: float = 3.0) -> float:
    """Measures server_utc - local_utc in seconds to compensate local clock drift."""
    for url in reference_urls:
        try:
            t0 = time.time()
            req = urllib.request.Request(url, method="HEAD")
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                header = resp.headers.get("Date")
            t1 = time.time()
            if header:
                server_ts = email.utils.parsedate_to_datetime(header).timestamp()
                return server_ts - ((t0 + t1) / 2.0)
        except Exception:
            continue
    return 0.0


def generate_totp_compensated(secret_b32: str, clock_offset_sec: float = 0.0, now_ts: float | None = None) -> tuple[str, int]:
    """Generates a 6-digit TOTP code applying clock_offset_sec and returns (code, remaining_seconds)."""
    clean = secret_b32.strip().replace(" ", "").upper()
    key = base64.b32decode(clean + "=" * (-len(clean) % 8))
    corrected = (now_ts if now_ts is not None else time.time()) + clock_offset_sec
    counter = int(corrected // 30)
    digest = hmac.new(key, struct.pack(">Q", counter), hashlib.sha1).digest()
    offset = digest[-1] & 0x0F
    code = (struct.unpack(">I", digest[offset : offset + 4])[0] & 0x7FFFFFFF) % 1_000_000
    remaining = 30 - int(corrected % 30)
    return str(code).zfill(6), remaining
