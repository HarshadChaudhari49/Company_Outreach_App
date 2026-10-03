# Company Outreach App

A local Flask application for organizing lawful, permission-based business outreach from an Excel contact list. It supports contact review, reusable message campaigns, personalized previews, controlled SMTP batches, and delivery tracking.

Source repository: [HarshadChaudhari49/Company_Outreach_App](https://github.com/HarshadChaudhari49/Company_Outreach_App)

## Features

- Password-protected workspace
- Import `.xlsx` and `.xlsm` contact lists with email validation and duplicate checks
- Recognize common headings such as `emailid1`, `email`, `url`, `clientname`, `company`, `contact_name`, `firstname`, `bussinesstype`, `country`, and `phoneormobile`
- Search, filter, and paginate contacts; edit contact details without changing their email or send history
- Suppression list to keep opted-out addresses out of future sends
- Create, edit, preview, and duplicate reusable campaigns
- Personalize subjects and messages with contact fields
- Send rate-limited SMTP batches
- Track pending, sent, failed, and suppressed contacts; inspect failures and selectively queue retries
- Export all contacts or the current filtered contact view as CSV

## Requirements

- Python 3.11 or newer
- SMTP account credentials for the mail provider you use

## Setup

From the project directory, create and activate a virtual environment, then install the dependencies:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Copy `.env.example` to `.env`, then set a strong `SECRET_KEY`, a private `ADMIN_PASSWORD`, and your SMTP settings. Keep `.env` private; it contains credentials and is excluded from Git.

For Gmail, use an App Password on an account with 2-Step Verification enabled. The app supports port `587` with STARTTLS and port `465` with implicit SSL. Never use your regular Google account password as the SMTP password.

Start the local development server:

```powershell
python run.py
```

Open <http://127.0.0.1:5000> and sign in using `ADMIN_USERNAME` and `ADMIN_PASSWORD` from `.env`.

## Configuration

| Setting | Purpose | Default |
| --- | --- | --- |
| `SECRET_KEY` | Flask session signing key | Must be changed |
| `ADMIN_USERNAME` | Initial administrator name | `admin` |
| `ADMIN_PASSWORD` | Initial administrator password | Must be changed |
| `SMTP_HOST` | SMTP server hostname | `smtp.gmail.com` |
| `SMTP_PORT` | SMTP port (`587` for STARTTLS or `465` for SSL) | `587` |
| `SMTP_USERNAME` | SMTP login and From address | — |
| `SMTP_PASSWORD` | SMTP credential or provider app password | — |
| `SMTP_USE_TLS` | Use STARTTLS on ports other than `465` | `true` |
| `SEND_INTERVAL_SECONDS` | Pause between messages in a batch | `10` |
| `MAX_SENDS_PER_RUN` | Maximum contacts processed per batch (capped at 100) | `20` |

The initial administrator is created when the database is first initialized. Changing the administrator values in `.env` later does not update an existing user's password.

## Typical workflow

1. Import an Excel workbook and review the imported contact fields.
2. Create or edit a campaign using variables such as `{{company}}`, `{{contact_name}}`, `{{email}}`, `{{website}}`, `{{business_type}}`, `{{country}}`, and `{{phone}}`.
3. Preview the rendered message and correct any missing or incorrect contact data.
4. Send the next controlled batch. The app processes pending contacts up to `MAX_SENDS_PER_RUN`, pausing between messages.
5. Review delivery statuses in Contacts. Failed contacts are not retried automatically; queue only the contacts you intend to retry, review the preview, then start a new batch.

The `Sent` status means the SMTP provider accepted the message for sending. It does not confirm delivery to the recipient's inbox.

## Data and deployment

The app stores its SQLite database at `instance/outreach.db`. The local database and the `instance/` directory are excluded from Git. Back up this directory securely if you need to preserve local contacts and campaign history.

`python run.py` starts Flask's development server. Do not expose that server directly to the public internet or use it as a production deployment; use a production WSGI server and appropriate HTTPS, access controls, and operational safeguards for a hosted installation.

Before contacting real recipients, confirm you have a lawful basis to contact them, follow applicable privacy and anti-spam requirements, honor opt-outs, and keep the suppression list current.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for the full text.
