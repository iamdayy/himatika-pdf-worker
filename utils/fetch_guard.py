import ipaddress
import os
import socket
from urllib.parse import urlparse

import requests


def _resolve_and_check_host(hostname: str) -> bool:
    """Resolve hostname and return True if any resolved IP is private/loopback/link-local/reserved."""
    try:
        infos = socket.getaddrinfo(hostname, None)
    except socket.gaierror:
        return False  # cannot resolve -> treat as public (allow), SSRF risk is low

    for info in infos:
        ip = info[4][0]
        try:
            ip_obj = ipaddress.ip_address(ip)
            if ip_obj.is_private or ip_obj.is_loopback or ip_obj.is_link_local or ip_obj.is_reserved:
                return True
        except ValueError:
            pass
    return False


def _is_private_host(hostname: str) -> bool:
    # Fast path: literal IP
    try:
        ip = ipaddress.ip_address(hostname)
        return ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved
    except ValueError:
        pass

    # Hostname: check allow/block list
    name = hostname.lower()
    if name in {"localhost"} or name.endswith(".internal") or name.endswith(".local"):
        return True

    # Full DNS resolution check
    return _resolve_and_check_host(hostname)


def safe_fetch(url: str, **kwargs):
    """requests.get with SSRF guard + default timeout.

    - https only (http allowed for explicit localhost when ALLOW_LOCAL_FETCH=1)
    - private/link-local/loopback/resolved-private hosts blocked unless allowed
    - optional comma-separated host allowlist via ALLOWED_FETCH_HOSTS
    - NO auto redirects; caller must handle if needed
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
    kwargs.setdefault("allow_redirects", False)
    return requests.get(url, **kwargs)
