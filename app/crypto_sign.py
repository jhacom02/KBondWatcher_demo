from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any, Optional

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import (
    Ed25519PrivateKey,
    Ed25519PublicKey,
)
from cryptography.hazmat.primitives.serialization import (
    Encoding,
    NoEncryption,
    PrivateFormat,
    PublicFormat,
)

from .paths import data_dir


class CryptoError(RuntimeError):
    pass


def _canonical(payload: dict[str, Any]) -> bytes:
    return json.dumps(
        payload, sort_keys=True, separators=(",", ":"), ensure_ascii=False
    ).encode("utf-8")


def generate_keypair() -> tuple[bytes, bytes]:
    private = Ed25519PrivateKey.generate()
    priv_bytes = private.private_bytes(Encoding.Raw, PrivateFormat.Raw, NoEncryption())
    pub_bytes = private.public_key().public_bytes(Encoding.Raw, PublicFormat.Raw)
    return priv_bytes, pub_bytes


def load_or_create_admin_private_key() -> Ed25519PrivateKey:
    """Admin-only. Never ship this with Trader binary."""
    env = (os.environ.get("KBOND_SIGNING_PRIVATE_KEY") or "").strip()
    if env:
        raw = base64.urlsafe_b64decode(env + "==")
        return Ed25519PrivateKey.from_private_bytes(raw[:32])
    path = data_dir() / "admin_signing_private.key"
    if path.is_file():
        raw = base64.urlsafe_b64decode(path.read_text(encoding="ascii").strip() + "==")
        return Ed25519PrivateKey.from_private_bytes(raw[:32])
    priv, pub = generate_keypair()
    path.write_text(base64.urlsafe_b64encode(priv).decode("ascii"), encoding="ascii")
    pub_path = data_dir() / "admin_signing_public.key"
    pub_path.write_text(base64.urlsafe_b64encode(pub).decode("ascii"), encoding="ascii")
    return Ed25519PrivateKey.from_private_bytes(priv)


def _public_key_bytes() -> bytes:
    env = (os.environ.get("KBOND_SIGNING_PUBLIC_KEY") or "").strip()
    if env:
        return base64.urlsafe_b64decode(env + "==")[:32]
    # Prefer packaged/public key file next to data or repo
    for candidate in (
        data_dir() / "admin_signing_public.key",
        Path(__file__).resolve().parents[1] / "keys" / "admin_signing_public.key",
    ):
        if candidate.is_file():
            return base64.urlsafe_b64decode(
                candidate.read_text(encoding="ascii").strip() + "=="
            )[:32]
    # Dev fallback: derive from admin private if present
    priv_path = data_dir() / "admin_signing_private.key"
    if priv_path.is_file():
        priv = load_or_create_admin_private_key()
        return priv.public_key().public_bytes(Encoding.Raw, PublicFormat.Raw)
    raise CryptoError(
        "no KBOND_SIGNING_PUBLIC_KEY configured; cannot verify signatures"
    )


def admin_sign_payload(payload: dict[str, Any]) -> str:
    private = load_or_create_admin_private_key()
    sig = private.sign(_canonical(payload))
    return base64.urlsafe_b64encode(sig).decode("ascii")


def verify_admin_signature(payload: dict[str, Any], signature: str) -> bool:
    if not signature:
        return False
    try:
        pub = Ed25519PublicKey.from_public_bytes(_public_key_bytes())
        pub.verify(base64.urlsafe_b64decode(signature + "=="), _canonical(payload))
        return True
    except (InvalidSignature, ValueError, CryptoError, OSError):
        return False


def export_public_key_b64() -> str:
    return base64.urlsafe_b64encode(_public_key_bytes()).decode("ascii")
