# Windows PC Renaming & Email Lookup Script

This PowerShell script is used to:

- Automatically detect the current user's Windows username.
- Check if the username contains a period (`.`) or dash (`-`).
- If the username contains a dash, the script replaces it with a period to match a specific email format.
- It will then match the email with a CSV file containing user data, including email and password.
- Launch Outlook, simulate user input, and login using the matched credentials.

## How to Use

This PowerShell script was designed to automate Microsoft Outlook login in a corporate environment. It relies on the Windows username (obtained via WindowsIdentity) to automatically generate the email address according to the company’s format (for example, firstname.lastname@example.com).

1. **Download the Script**: Clone or download this repository.

2. **Prepare the CSV File**:
   - The CSV file should contain at least two columns: `email` and `password`.
   - The CSV should be delimited by `;` (semicolon).

3. **Configure the Path to the CSV File**:
   - Update the `$csvFile` variable with the correct path to your CSV file in the script.
   - Example: `$csvFile = "C:\Path\To\Your\CSV\File.csv"`

4. **Run the Script**:
   - Execute the script in PowerShell.
   - The script will automatically fetch the current user, match the email in the CSV, and log in to Outlook.

5. **Requirements**:
   Windows PowerShell 5.1 or higher.

   Outlook installed on the machine.

   Each machine must have a properly configured Windows username (Windows identity) so that the email address is automatically generated.

   A CSV file containing the credentials, with semicolon (;) as the delimiter and with the columns email and password.
   Example CSV format :csv
            " email;password
            " firstname.lastname@example.com;password1
            " john.doe@example.com;password2

## Troubleshooting

- If the script doesn't find a matching email in the CSV file, it will prompt you to input a new PC name using `-` instead of `.` (e.g., `john-doe`,the new PC name will be updated after your restart your device ).
- Make sure that the Outlook path in the script matches your system's installation.

## License

This script is provided as-is. Use it at your own risk.
