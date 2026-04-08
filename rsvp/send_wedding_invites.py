"""
Wedding WhatsApp Invite Sender
-------------------------------
Reads guest names and phone numbers from an Excel file,
then sends each guest a personalised WhatsApp message.

Requirements:
    pip install pywhatkit openpyxl

Before running:
    1. Fill in wedding_guests.xlsx with your guest list.
    2. Open WhatsApp Web (https://web.whatsapp.com) and stay logged in.
    3. Run this script — keep your computer on and Chrome open.
"""

import pywhatkit as kit
import openpyxl
import time

# ── Config ────────────────────────────────────────────────────────────────────
EXCEL_FILE = "wedding_guests.xlsx"   # Must be in the same folder as this script
WAIT_TIME  = 20  # Seconds to wait for WhatsApp Web to load (increase if slow)
DELAY      = 10    # Seconds between each message (be gentle on WhatsApp)
# ─────────────────────────────────────────────────────────────────────────────


def build_message(first_name: str) -> str:
    return (
        f"Dear {first_name},\n\n"
        "Brehmer and I are getting married, and we'd love for you to be there to celebrate with us!\n\n"
        "📅 Date & Time: 8th August 2026, 12pm\n"
        "📍 Venue: Parkroyal Collection at Marina Bay\n\n"
        "For full details including the schedule, venue info, and any questions, do check out our wedding website: "
        "https://withjoy.com/brehmerjiaxin\n\n"
        "We'd love your RSVP through the website by Friday, 8th May.\n\n"
        "Can't wait to celebrate with you! 🥂"
    )


def load_guests(filepath: str) -> list[tuple[str, str]]:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    guests = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name, phone = row[0], row[1]
        # Skip empty rows or the placeholder note row
        if name and phone and str(phone).startswith("+"):
            guests.append((str(name).strip(), str(phone).strip()))
    return guests


def send_invites():
    guests = load_guests(EXCEL_FILE)

    if not guests:
        print("No guests found. Check your Excel file.")
        return

    print(f"Found {len(guests)} guest(s). Starting in 5 seconds...")
    print("Make sure WhatsApp Web is open and logged in!\n")
    time.sleep(5)

    success, failed = [], []

    for i, (name, phone) in enumerate(guests, 1):
        message = build_message(name)
        print(f"[{i}/{len(guests)}] Sending to {name} ({phone})...")
        try:
            kit.sendwhatmsg_instantly(
                phone_no=phone,
                message=message,
                wait_time=WAIT_TIME,
                tab_close=True,
                close_time=3,
            )
            print(f"  ✓ Sent to {name}")
            success.append(name)
        except Exception as e:
            print(f"  ✗ Failed for {name}: {e}")
            failed.append((name, phone))

        if i < len(guests):
            print(f"  Waiting {DELAY}s before next message...\n")
            time.sleep(DELAY)

    # Summary
    print("\n─── Summary ───────────────────────────")
    print(f"  Sent:   {len(success)}")
    print(f"  Failed: {len(failed)}")
    if failed:
        print("\n  Failed guests (resend manually):")
        for name, phone in failed:
            print(f"    • {name}  {phone}")
    print("────────────────────────────────────────")


if __name__ == "__main__":
    send_invites()
