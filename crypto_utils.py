"""Encrypt/decrypt refresh tokens at rest using Fernet (symmetric, authenticated).

ENCRYPTION_KEY must be a URL-safe base64-encoded 32-byte key, e.g. generated with:
    python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
"""

import os

from cryptography.fernet import Fernet
from dotenv import load_dotenv

load_dotenv()

_raw_key = os.environ.get("ENCRYPTION_KEY")
if not _raw_key or _raw_key.startswith("replace-with"):
    raise RuntimeError(
        "ENCRYPTION_KEY is not set. Generate one with: "
        "python -c \"from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())\" "
        "and put it in your .env file."
    )

_fernet = Fernet(_raw_key.encode())


def encrypt_token(plaintext: str) -> str:
    """Encrypt a plaintext token, returning a base64 ciphertext string safe for DB storage."""
    return _fernet.encrypt(plaintext.encode("utf-8")).decode("utf-8")


def decrypt_token(ciphertext: str) -> str:
    """Decrypt a ciphertext previously produced by encrypt_token."""
    return _fernet.decrypt(ciphertext.encode("utf-8")).decode("utf-8")
