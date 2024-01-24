# GUI_Get_MessageTrackingLog
# PowerShell Script Documentation: Top Mail Senders Report on Edge Server

# Overview
This PowerShell script is designed to run on an Edge Server with Exchange Management Shell and admin privileges. It generates reports on the top mail senders within a specified time range. The script provides a user interface to select the start and end dates and times, and options to generate a short list of top 200 senders or a full list of senders and recipients.

# Prerequisites
- PowerShell 5.1 or later.
- Exchange Management Shell.
- Admin privileges on the Edge Server.
- Ensure the path D:\Script_output_senders\ exists or modify the script to use an existing path.

# Usage
- Open Exchange Management Shell as an administrator.
- Navigate to the directory containing the script.
- Run the script.
- Follow the on-screen prompts to select the desired date range and report type.
- Check the specified output directory for the generated CSV files.

![image](https://github.com/Nivellem/GUI_Get_MessageTrackingLog/assets/84031994/910826cc-59ff-4908-b2fc-f5a35edc7e3f)


# Notes
Ensure the output directory has appropriate write permissions.
The script may require modifications to match the specific Exchange and server environment.
This script is intended for use by experienced Exchange administrators.
# Script Breakdown

# Data Collection

1. Server Information Retrieval: Fetches the name and domain of the server.

```Powershell
$Server_data = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object Name,domain
$servers = $Server_data.name+"."+$Server_data.domain
```

2. Timestamp Generation: Creates a timestamp for file naming.

```Powershell
$timestamp = Get-Date -Format "MMddT_HH_mm_ss"
```

# File Paths
3. Output File Paths: Defines the paths for the output CSV files.

```Powershell
$OutFile_Table = "D:\Script_output_senders\List_of_top_sender$timestamp.csv"
$OutFile = "D:\Script_output_senders\All_senders_and_recipients_$timestamp.csv"
```
4. Form Creation: A Windows Form is created for user interaction.
```Powershell
$form = New-Object System.Windows.Forms.Form
...

```
5. Date and Time Selection: Users can select start and end dates and times using calendar and numeric up/down controls.
```Powershell
$calendar = New-Object System.Windows.Forms.MonthCalendar
$hourPicker = New-Object System.Windows.Forms.NumericUpDown
...

```
6. Checkboxes for Report Type: Users can select whether they want the top 200 senders or a full list of senders and recipients.
```Powershell
$calendar = New-Object System.Windows.Forms.MonthCalendar
$hourPicker = New-Object System.Windows.Forms.NumericUpDown
...
```
7. Message Tracking Log Retrieval: Retrieves message tracking logs from the server for the specified date range.
```Powershell
foreach ($s in $servers) {
    $info += Get-MessageTrackingLog -Start (Get-Date $data_start) -End (Get-Date $data_end) -ResultSize Unlimited 
}

```
8. Top 200 Senders Report: If selected, generates a report of the top 200 mail senders.
```Powershell
$list_of_senders = $info | Group-Object -Property Sender | ...

```
9. Full List Report: If selected, generates a detailed report including all senders and recipients.
```Powershell
$info | Select-Object ... | Export-Csv $outfile -NoTypeInformation -Delimiter ";"

```

# Execution Flow
The script displays the form, allowing the user to select options and generate reports. It provides feedback on the console about the actions being performed and the status of the report generation.

















