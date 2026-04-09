"""Standalone script to send bandwidth allocation emails to all recipients."""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), "../.env"))

from services import sheets, email as email_svc

TEST_MODE = os.getenv("TEST_MODE", "true").lower() == "true"
TEST_EMAIL = os.getenv("TEST_EMAIL", "")
QUARTER_LABEL = os.getenv("QUARTER_LABEL", "Q1'26")


def main():
    print("\n=========================================")
    print("  Bandwidth Bot — Email Sender")
    print(f"  Quarter   : {QUARTER_LABEL}")
    print(f"  Test mode : {TEST_MODE}" + (f" (→ {TEST_EMAIL})" if TEST_MODE else ""))
    print("=========================================\n")

    print("[1/4] Reading capitalised rows from sheet…")
    rows = sheets.get_capitalised_rows()
    print(f"      Found {len(rows)} capitalised ticket(s).")

    if not rows:
        print("\nNo capitalised rows found. Exiting.")
        return

    print("[2/4] Building people map…")
    people_map = sheets.build_people_map()
    print(f"      Found {len(people_map)} people in People<>Soln.")

    print("[3/4] Determining recipients…")
    recipients = sheets.get_all_recipients(rows, people_map)
    print(f"      {len(recipients)} recipient(s) to email.\n")

    if not recipients:
        print("No valid recipients. Check name matching and email validity.")
        return

    print("Recipients:")
    for r in recipients:
        print(f"  - {r['name']} <{r['email']}>")
    print()

    print("[4/4] Sending emails…\n")
    results = email_svc.send_all_emails(recipients)

    print("\n=========================================")
    print("  SUMMARY")
    print("=========================================")
    print(f"  Sent  : {len(results['sent'])}")
    print(f"  Failed: {len(results['failed'])}")
    if results["failed"]:
        print("\nFailed:")
        for f in results["failed"]:
            print(f"  - {f['name']}: {f.get('error')}")
    print("\nDone.\n")


if __name__ == "__main__":
    main()
