import ipaddress
import os
from urllib.parse import urlparse

import requests


def _is_private_host(hostname: str) -> bool:
    try:
        ip = ipaddress.ip_address(hostname)
        return ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved
    except ValueError:
        name = hostname.lower()
        return (
            name in {"localhost"}
            or name.endswith(".internal")
            or name.endswith(".local")
        )


def safe_fetch(url: str, **kwargs):
    """requests.get with SSRF guard + default timeout.

    - https only (http allowed for explicit localhost when ALLOW_LOCAL_FETCH=1)
    - private/link-local/loopback hosts blocked unless allowed
    - optional comma-separated host allowlist via ALLOWED_FETCH_HOSTS
    """
    parsed = urlparse(url)
    allow_local = os.getenv("ALLOW_LOCAL_FETCH", "").lower() in {"1", "true", "yes"}

    if parsed.scheme == "https":
        pass
    elif parsed.scheme == "http" and allow_local:
        pass
    else:
        raise ValueError(f"Blocked non-https fetch: {url}")

    host = parsed.hostname or ""
    if not allow_local and _is_private_host(host):
        raise ValueError(f"Blocked fetch to private host: {host}")

    allowlist = [
        h.strip().lower()
        for h in os.getenv("ALLOWED_FETCH_HOSTS", "").split(",")
        if h.strip()
    ]
    if allowlist and host.lower() not in allowlist:
        raise ValueError(f"Host not in ALLOWED_FETCH_HOSTS: {host}")

    kwargs.setdefault("timeout", (5, 30))
    return requests.get(url, **kwargs)
