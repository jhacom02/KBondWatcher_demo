from __future__ import annotations

import base64
from typing import Literal, Protocol

ProtectionMethod = Literal["dpapi", "tpm", "plaintext_dev"]


class SecretProtector(Protocol):
    method: ProtectionMethod

    def protect(self, secret: bytes) -> bytes: ...

    def unprotect(self, blob: bytes) -> bytes: ...


class TpmProtector:
    """Stub for future TPM/NCrypt backing. Not used in Pilot v1."""

    method: ProtectionMethod = "tpm"

    def protect(self, secret: bytes) -> bytes:
        raise NotImplementedError("TPM protector not implemented; use DPAPI")

    def unprotect(self, blob: bytes) -> bytes:
        raise NotImplementedError("TPM protector not implemented; use DPAPI")


class DpapiProtector:
    method: ProtectionMethod = "dpapi"

    def protect(self, secret: bytes) -> bytes:
        import win32crypt

        encrypted = win32crypt.CryptProtectData(secret, "KBondWatcher", None, None, None, 0)
        return bytes(encrypted)

    def unprotect(self, blob: bytes) -> bytes:
        import win32crypt

        _desc, raw = win32crypt.CryptUnprotectData(blob, None, None, None, 0)
        return bytes(raw)


def get_protector(*, prefer_tpm: bool = False) -> SecretProtector:
    if prefer_tpm:
        # Interface reserved; fall through until implemented.
        pass
    return DpapiProtector()


def protect_secret(secret: bytes, *, prefer_tpm: bool = False) -> tuple[str, ProtectionMethod]:
    protector = get_protector(prefer_tpm=prefer_tpm)
    blob = protector.protect(secret)
    return base64.urlsafe_b64encode(blob).decode("ascii"), protector.method


def unprotect_secret(blob_b64: str, method: ProtectionMethod = "dpapi") -> bytes:
    if method == "tpm":
        raise NotImplementedError("TPM unprotect not implemented")
    if method == "plaintext_dev":
        return base64.urlsafe_b64decode(blob_b64 + "==")
    protector = DpapiProtector()
    return protector.unprotect(base64.urlsafe_b64decode(blob_b64 + "=="))
