"""Băm và kiểm tra mật khẩu (PBKDF2-HMAC-SHA256)."""
from __future__ import annotations

import hashlib
import os
import secrets

_ITERATIONS = 200_000


def hash_password(plain: str) -> tuple[str, str]:
    """Trả về (salt_hex, hash_hex)."""
    salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", plain.encode("utf-8"), salt, _ITERATIONS)
    return salt.hex(), dk.hex()


def verify_password(plain: str, salt_hex: str, hash_hex: str) -> bool:
    try:
        salt = bytes.fromhex(str(salt_hex).strip())
        expected = bytes.fromhex(str(hash_hex).strip())
    except ValueError:
        return False
    dk = hashlib.pbkdf2_hmac("sha256", plain.encode("utf-8"), salt, _ITERATIONS)
    return secrets.compare_digest(dk, expected)
