# Bulk Email Script

This Python script allows you to send bulk emails with personalized content to multiple recipients by using an email list and a predefined message template. It also supports sending an attached file (e.g., a resume) with each email.

## Features

- Send bulk emails using an Excel (`.xlsx`) or CSV file containing the list of recipients.
- Personalize email content using placeholders that are replaced with recipient-specific data from the spreadsheet.
- Attach files (such as a PDF resume) to each email.
- Configurable SMTP server settings for different email providers.

## Requirements

- Python 3.x
- Required Python libraries:
  - `smtplib`
  - `email`
  - `pandas`
  - `argparse`
  - `re`
  - `os`

Install the required libraries by running:

```
pip install -r requirements.txt

```

## Usage

```bash
python bulk-email.py -r <path_to_resume> -l <path_to_list> -m <path_to_message> -u <your_email> -p <your_password> --subject <email_subject> [-s <smtp_server>] [-po <smtp_port>]
```

### Command-Line Arguments

- `-r, --resume` (required): Path to the file you want to attach (e.g., a resume).
- `-l, --list` (required): Path to the Excel (`.xlsx`) or CSV (`.csv`) file containing the recipient information, including their email addresses.
- `-m, --message` (required): Path to the message template file. The message can include placeholders using column names from the Excel/CSV file in the format `{{column_name}}`.
- `-u, --username` (required): Your email address.
- `-p, --password` (required): Your email account password.
- `--subject`: The subject line of the email.
- `-s, --server`: SMTP server to use for sending the emails (default: `smtp.gmail.com`).
- `-po, --port`: Port number of the SMTP server (default: `587`).

### Example

```bash
python bulk-email.py -r ./resume.pdf -l ./email_list.xlsx -m ./message.txt -u example@gmail.com -p mypassword --subject "Job Application" -s smtp.gmail.com -po 587
```

In this example:

- A resume (`resume.pdf`) will be attached to each email.
- The list of recipients is loaded from `email_list.xlsx`.
- The email content will be loaded from `message.txt` and personalized for each recipient using the columns from the spreadsheet.
- The script will use Gmail's SMTP server to send the emails.

### Message Template

You can personalize the email by using placeholders in the message template file. For example:

```
Dear {{name}},

I hope this email finds you well. I am interested in the {{position}} at {{company}}.

Best regards,
Your Name
```

In the above template, `{{name}}`, `{{position}}`, and `{{company}}` will be replaced with the corresponding values from the spreadsheet.

### Spreadsheet Format

The Excel or CSV file should contain a column named `email` for the recipient's email address, and any other columns needed for placeholders in the message template. Example format:

| name | position  | company   | email            |
| ---- | --------- | --------- | ---------------- |
| John | Developer | Google    | john@example.com |
| Jane | Manager   | Microsoft | jane@example.com |

#### Note

> The column names will be used as the variables that need to be replaced in the message, so keep in mind to use the same variables in both Excel sheet and Message

## Error Handling

The script validates:

- The file paths for the resume, list, and message.
- That the email list contains a valid `email` column.
- That the sender's email is valid.

If any issues are found, the script will raise an appropriate error.

## SMTP Server Configuration

By default, the script is configured to use Gmail's SMTP server (`smtp.gmail.com`) on port `587`. You can specify a different SMTP server and port using the `-s` and `-po` arguments if necessary.

## License

This project is licensed under the MIT License.
