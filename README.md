# Outlook Automation PowerShell Script

This PowerShell script is designed to automate the login process for Outlook using a Windows username-based email. The script checks for matching credentials in a CSV file and simulates the login process for Outlook. Ideal for automating the setup of multiple user accounts in corporate environments.

## Features
- Automatically retrieves the Windows username and generates an email.
- Matches the generated email with credentials stored in a CSV file.
- Launches Outlook and simulates typing the username and password for automated login.
- Uses network connectivity checks to verify Outlook's readiness before entering the password.

## Usage
1. Modify the path to the CSV file and Outlook executable in the script.
2. Ensure the CSV file contains columns for `email` and `password`.
3. Run the script using PowerShell with administrator privileges.
4. The script will automatically detect the username, match it with an email in the CSV, and proceed with the login.

### Requirements
- Windows OS with PowerShell support.
- Outlook installed on the system.
- A properly formatted CSV file with user credentials.

## License
This script is provided as-is. Use at your own risk. Please ensure that the script is adapted to your own environment and security policies.
