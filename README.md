# whatsapp-bot

Send personalized messages via WhatsApp Web using contacts from an Excel file, event details from JSON, and a message template.

## Files

- [main.py](main.py) — core program with functions: `read_contacts`, `read_event`, `read_template`, `setup_driver`, `wait_for_logged_in`, `send_message`, `send_messages`
- [template.txt](template.txt) — message template with placeholders
- [requirements.txt](requirements.txt) — Python dependencies
- [.gitignore](.gitignore) — ignored files/patterns

### Files you need to create (not in repo)

- `contacts.xlsx` — Excel file with contact information (see format below)
- `event.json` — JSON file with event details (see format below)

### Example files provided

- [contacts.example.xlsx](contacts.example.xlsx) — example contact format
- [event.example.json](event.example.json) — example event format

## Description

This tool reads:

- Contact info (Name, Nickname, Number) from `contacts.xlsx`
- Event details from `event.json`
- Message template from `template.txt`

It then opens WhatsApp Web via Selenium and sends personalized messages to each unique phone number.

## File Formats

### contacts.xlsx

Excel file with these columns (case-sensitive):

- `Name` — full name
- `Nickname` — personalized greeting name
- `Number` — phone number (without country code)

Example:

```text
Name          | Nickname | Number
John Doe      | John     | 791234567
Jane Smith    | Jane     | 799876543
```

### event.json

JSON file with event details. Keys must match placeholders in template:

```json
{
  "groom": "Groom Name",
  "bride": "Bride Name",
  "date": "DD-MM-YYYY",
  "day": "Day Name",
  "time": "Event Time",
  "gather_time": "Gathering Time",
  "destination": "Event Location",
  "source": "Meeting Point",
  "map1": "https://maps.google.com/link1",
  "map2": "https://maps.google.com/link2"
}
```

### template.txt

Message template with placeholders in `{placeholder}` format:

- `{nickname}` — replaced with contact's nickname from Excel
- Other placeholders — replaced with values from event.json

Example:

```text
السلام عليكم {nickname}
اتشرف بدعوتكم لجاهة خطوبة ابني:
{groom} و {bride}
- يوم {day} الموافق {date} الساعة ({time}) {destination}
```

## Setup

1. Create and activate a virtual environment:

   ```bash
   # Windows (PowerShell)
   python -m venv myenv
   .\myenv\Scripts\Activate.ps1
   
   # Windows (cmd)
   python -m venv myenv
   .\myenv\Scripts\activate
   
   # macOS / Linux
   python3 -m venv myenv
   source myenv/bin/activate
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Create required files:
   - Copy `event.example.json` to `event.json` and fill in your event details
   - Create `contacts.xlsx` with your contacts (use `contacts.example.xlsx` as reference)

## Usage

1. Ensure Chrome is installed (Selenium will auto-download ChromeDriver)
2. Run:

   ```bash
   python main.py
   ```

The script will:

- Open WhatsApp Web and wait for login (scan QR code on first run)
- For each contact:
  - Open chat with their number
  - Insert personalized message (nickname from Excel + event details from JSON)
  - Send the message
- Log successes and failures

## Notes & Troubleshooting

- Scan QR code on first run; session persists in `chrome-data/` folder
- Phone numbers should be plain digits without country code (default: `962` for Jordan)
- To change country code, edit `country_code='962'` in `send_message()` function
- Duplicate numbers are automatically removed
- Script logs progress to console
- If ChromeDriver download hangs, run `pip install --upgrade selenium` and try again
- Automation may violate WhatsApp terms; use responsibly

## Reusability

The bot is designed to be reusable:

- Keep the same `contacts.xlsx` for different events
- Just update `event.json` with new event details
- Modify `template.txt` to change message format
- Add new placeholders: add keys to `event.json` and use `{key}` in template

## License

Personal project — use responsibly
