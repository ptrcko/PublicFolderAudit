# Public Folders Item Count and Last Received Date Exporter

This VBA script retrieves public folder item counts and the date of the most recently received email, exporting the information into a CSV file. It organizes the data hierarchically to indicate the folder structure.

This is really helpful if you want to cleanup a legacy set of public folders and saves a lot of time from manually checking each folder.

## Prerequisites

- Microsoft Outlook (desktop version)
- Basic familiarity with Outlook's VBA Editor
- Permissions to access the desired public folders in Outlook

---

## Instructions

### 1. Setting Up the Script in Outlook

1. **Open the VBA Editor in Outlook:**
   - Press `ALT + F11` to open the Outlook VBA Editor.
  
2. **Create a New Module:**
   - In the VBA Editor, click `Insert` > `Module` to create a new module.
  
3. **Paste the Script:**
   - Copy and paste the provided VBA script into the new module.
  
4. **Adjust the Script for Your Setup:**
   - **Select the Correct Root Folder:**
     - The script uses `Set olRootFolder = olNamespace.Folders.Item(1).Folders.Item(2)` to access the root public folder.
     - Adjust `Item(1)` and `Item(2)` as needed to point to your specific public folder. You may need to explore the folder structure in Outlook to determine the appropriate values.
   - **Set the Output File Path:**
     - Update `outputFilePath = "C:\outputpath.csv"` to your desired file path for saving the CSV file.

### 2. Running the Script

1. **Enable Macros in Outlook:**
   - To run macros in Outlook, you may need to enable macro settings. See the "Enabling Macros in Outlook" section below for more details.
2. **Run the Macro:**
   - In the VBA Editor, press `F5` or click `Run` to execute the script.
   - A message box will inform you when the data export is complete, and the CSV file will be saved at the specified location.

---

## Enabling Macros in Outlook

Macros are disabled by default in Outlook to improve security. You must enable them to run this script.

### Steps to Enable Macros:

1. Open Outlook and go to `File` > `Options`.
2. Click on `Trust Center` > `Trust Center Settings`.
3. Go to `Macro Settings` and select **"Notifications for all macros"** or **"Enable all macros"** (not recommended, use with caution).
4. Click `OK` to save your settings.

For more information on enabling and running macros in Outlook, refer to the official Microsoft documentation:
- [Enable or disable macros in Office files](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)
- [Outlook VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook)

---

## Notes

- **Security Warning**: Be cautious when enabling macros and only run VBA scripts from trusted sources.
- **Testing**: It's recommended to test the script in a controlled environment before using it with important data.

---

Feel free to customize the script as needed for your specific requirements. If you encounter any issues or have questions, consult the links above or seek guidance from your IT administrator.