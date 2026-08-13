import base64
import hashlib
import unittest

from cryptography.fernet import Fernet, InvalidToken

import general_server


class EncryptionSecurityTests(unittest.TestCase):
    def test_encryption_uses_versioned_randomized_container(self):
        first = general_server._encrypt_payload(b"confidential", "correct horse")
        second = general_server._encrypt_payload(b"confidential", "correct horse")

        self.assertTrue(first.startswith(general_server.ENCRYPTION_MAGIC))
        self.assertNotEqual(first, second)
        plaintext, legacy = general_server._decrypt_payload(first, "correct horse")
        self.assertEqual(plaintext, b"confidential")
        self.assertFalse(legacy)

    def test_wrong_password_is_rejected(self):
        encrypted = general_server._encrypt_payload(b"confidential", "right")
        with self.assertRaises(InvalidToken):
            general_server._decrypt_payload(encrypted, "wrong")

    def test_legacy_payload_remains_readable(self):
        password = "legacy-password"
        legacy_key = base64.urlsafe_b64encode(
            hashlib.sha256(password.encode("utf-8")).digest()
        )
        legacy_payload = Fernet(legacy_key).encrypt(b"legacy data")

        plaintext, legacy = general_server._decrypt_payload(legacy_payload, password)
        self.assertEqual(plaintext, b"legacy data")
        self.assertTrue(legacy)

    def test_empty_password_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "密码不能为空"):
            general_server._encrypt_payload(b"data", "")


if __name__ == "__main__":
    unittest.main()
