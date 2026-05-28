from __future__ import annotations

import getpass
import pathlib
import sys

from auth import _hash

USERS_FILE = pathlib.Path(__file__).parent / "users.txt"


def main() -> None:
    print("Set login credentials (writes users.txt)\n")
    username = input("Username: ").strip()
    if not username:
        print("Username cannot be empty.")
        return

    password = getpass.getpass("Password: ")
    confirm = getpass.getpass("Confirm password: ")
    if password != confirm:
        print("Passwords do not match. Nothing written.")
        return
    if not password:
        print("Password cannot be empty.")
        return

    content = (
        f"{username}:{_hash(password)}\n"
    )
    USERS_FILE.write_text(content, encoding="utf-8")
    print(f"\nSaved {USERS_FILE.name} for user '{username}'.")


def hash_only() -> None:
    """Print the hash for a password — paste it into APP_PASSWORD_HASH (e.g. Railway)."""
    password = getpass.getpass("Password: ")
    confirm = getpass.getpass("Confirm password: ")
    if password != confirm:
        print("Passwords do not match.")
        return
    if not password:
        print("Password cannot be empty.")
        return
    print("\nAPP_PASSWORD_HASH=" + _hash(password))


if __name__ == "__main__":
    if "--hash" in sys.argv:
        hash_only()
    else:
        main()
