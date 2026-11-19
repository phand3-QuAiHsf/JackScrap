"""
Run this ONCE (or any time you need to change the username/password) to
store your Sustainalytics login for the scraper to use automatically.

There is NO master password - anyone who can run this script can set or
update the stored username/password. The credentials are saved to
credentials.enc in this folder, lightly encoded so they aren't sitting
around as obvious plain text if someone opens the file directly, but this
is NOT strong encryption/security - treat credentials.enc like you would
a saved browser password: don't upload it anywhere public or share the
folder with people who shouldn't have this login.

WHAT YOU NEED TO DO:
1. Run:  python setup_credentials.py
2. Enter your Sustainalytics username and password.

Re-run this script any time you want to update the stored username or
password - it overwrites the old credentials.enc automatically, no
password or confirmation needed.
"""

import base64
import getpass
import json

CRED_FILE = "credentials.enc"


def main():
    print("=== Sustainalytics Credential Setup ===\n")
    username = input("Sustainalytics username/email: ").strip()
    password = getpass.getpass("Sustainalytics password: ")

    payload = json.dumps({"username": username, "password": password}).encode()
    encoded = base64.urlsafe_b64encode(payload)

    with open(CRED_FILE, "wb") as fh:
        fh.write(encoded)

    print(f"\nSaved credentials to {CRED_FILE}")
    print("The scraper will read this file automatically - no password needed to run it.")
    print("Run this script again any time to update the username/password.")


if __name__ == "__main__":
    main()
