💻 Business Continuity Plan Verifier
This script automates the process of verifying the bcp file is up to date and sends a reminder email to the helpdesk team. It helps ensure that in the event of an outage, we still have a business continuity plan to keep operations running.

-- THIS IS ON MY LAPTOP CURRENTLY. IT WILL NOT WORK ON A VM --

📦 Features

finds the xlsx file on the file server and checks its creation date
if the creation date is today then we're fine, if not it raises a ticket to the helpdesk
Then if there's an issue you speak with the data team to identify what the issue is and how we resolve it

🧰 Requirements

Python 3.8+
pywin32 (for Outlook integration)

Install dependencies:
run uv sync to get dependencies from the toml file

📁 File Structure
project/
│
├── main.py # Main script
└── README.md # Documentation

⚙️ Configuration
"FilePathHere" - File Server file being targetted

🚀 Usage
Set up a batch file with the below, put your user path in and then specify on your c drive where you cloned the repo. This is just where most of my stuff went:
C:\Users\YOURUSERPATHHERE\AppData\Local\Programs\Python\Python313\python.exe C:\Utilities\Python\BCP_CHECKER\main.py
Then set up a schedule task to run once a week on your chosen day and time that targets this batch file and runs it. i've set it to run at 9.45am every morning

This will:

Read the file on the file server.
Check the date to ensure it's today
Send an email to the helpdesk with the status of either everything's fine or we need to contact the data team to identify the issue

✉️ Email Setup
The script uses win32com.client to send emails via Outlook. The recipient is currently set to:
mail.To = "EmailAddressHere"

You can change this to any valid email address or distribution list.

🧪 Testing
No testing required unless tampered with. Works as intended

📌 Notes
run a uv sync in the terminal to pull the dependencies from the toml
if you need uv run pip install uv in the terminal to add to your global scope