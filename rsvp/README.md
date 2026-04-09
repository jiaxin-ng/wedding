# Wedding WhatsApp Invite Sender

This script sends personalized wedding invites through WhatsApp Web using contacts in an Excel file.

Script: `send_wedding_invites.py`

## What It Does

- Reads guest names and phone numbers from `wedding_guests.xlsx`
- Builds a custom message for each guest
- Sends each message via WhatsApp Web
- Prints a success/failure summary at the end

## Requirements

- Python 3.9+ recommended
- Google Chrome installed
- WhatsApp Web session already logged in
- Python packages:

```bash
pip install pywhatkit openpyxl
```

## Files in This Folder

- `send_wedding_invites.py` - main script
- `wedding_guests.xlsx` - guest list input

## Excel Format

Use `wedding_guests.xlsx` with:

- Row 1 as headers (for example: `name`, `phone`)
- Data from row 2 onward
- Phone numbers in full international format, starting with `+`
  - Example: `+6591234567`

The script currently reads:

- Column A -> guest name
- Column B -> phone number

Rows are skipped if:

- Name is empty
- Phone is empty
- Phone does not start with `+`

## Configuration

In `send_wedding_invites.py`:

- `EXCEL_FILE` - input file name
- `WAIT_TIME` - seconds to wait for WhatsApp Web to load per message send
- `DELAY` - seconds to pause between guests

Increase `WAIT_TIME` if your browser/network is slow.

## How to Run

1. Open WhatsApp Web and make sure you are logged in.
2. Keep Chrome open.
3. In terminal, go to this folder:

```bash
cd "/Users/user/Documents/github/wedding/rsvp"
```

4. Run:

```bash
python3 send_wedding_invites.py
```

5. Do not use the keyboard/mouse much while messages are being sent.

## Safety Notes

- Start with a small test batch first (for example 1-3 contacts).
- Keep delays reasonable to avoid spam-like behavior.
- Review message content in `build_message()` before full send.
- Respect recipient consent and local messaging rules.

## Troubleshooting

- **"No guests found"**
  - Check that `wedding_guests.xlsx` exists in this same folder.
  - Confirm rows from row 2 contain name + `+` phone number.

- **Messages not sending / wrong timing**
  - Increase `WAIT_TIME` and `DELAY`.
  - Ensure WhatsApp Web is open and logged in before running.

- **Import errors**
  - Reinstall dependencies:
    - `pip install pywhatkit openpyxl`

## Suggested Next Improvements

- Add a dry-run mode (preview recipients without sending)
- Add logging to CSV for sent/failed entries
- Add duplicate phone number checks before sending

