$Server_data = Get-CimInstance -ClassName Win32_ComputerSystem |Select-Object Name,domain
$servers = $Server_data.name+"."+$Server_data.domain
$timestamp= Get-Date -Format "MMddT_HH_mm_ss"

$OutFile_Table="D:\Script_output_senders\List_of_top_sender$timestamp.csv"
$OutFile="D:\Script_output_senders\All_senders_and_recipients_$timestamp.csv"

# Load necessary assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create a new form (window)
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Date and Time'
$form.Size = New-Object System.Drawing.Size(700,500) # Adjusted size for better layout
$form.StartPosition = 'CenterScreen'

# Calculate date and time minus 24 hours
$defaultDateTime = (Get-Date).AddHours(-24)
$defaultDateTime_Today_end = (Get-Date)

# Initialize variables
$top_200_senders = $true
$full_list_senders = $false

# Create CheckBox for Top 200 short list
$top200CheckBox = New-Object System.Windows.Forms.CheckBox
$top200CheckBox.Location = New-Object System.Drawing.Point(300,340)
$top200CheckBox.Size = New-Object System.Drawing.Size(160,20)
$top200CheckBox.Text = 'Top Sender Short List'
$top200CheckBox.Checked = $true
$form.Controls.Add($top200CheckBox)

# Create CheckBox for Full List of Senders
$fullListCheckBox = New-Object System.Windows.Forms.CheckBox
$fullListCheckBox.Location = New-Object System.Drawing.Point(300,370)
$fullListCheckBox.Size = New-Object System.Drawing.Size(160,20)
$fullListCheckBox.Text = 'Full List of Senders'
$fullListCheckBox.Checked = $false
$form.Controls.Add($fullListCheckBox)

# Add checkbox changed event handlers
$top200CheckBox.add_CheckedChanged({ Update-Preview })
$fullListCheckBox.add_CheckedChanged({ Update-Preview })

# Function to update the preview
function Update-Preview {
    $selectedDate = $calendar.SelectionStart.Date
    $selectedHour = [int]$hourPicker.Value
    $selectedMinute = [int]$minutePicker.Value
    $previewLabel.Text = "Start Date: $($selectedDate.AddHours($selectedHour).AddMinutes($selectedMinute).ToString('g'))"
}

function Update-Preview_end {
    $selectedDate_end = $calendar_END.SelectionStart.Date
    $selectedHour_end = [int]$hourPicker_end.Value
    $selectedMinute_end = [int]$minutePicker_end.Value
    $previewLabel_end.Text = "End Date: $($selectedDate_end.AddHours($selectedHour_end).AddMinutes($selectedMinute_end).ToString('g'))"
}

# Create a label for the calendar
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(230,20)
$label.Text = 'Start Date'
$form.Controls.Add($label)

# Create a label end  for the calendar
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(310,10)
$label.Size = New-Object System.Drawing.Size(230,20)
$label.Text = 'End Date'
$form.Controls.Add($label)

# Create a calendar control
$calendar = New-Object System.Windows.Forms.MonthCalendar
$calendar.Location = New-Object System.Drawing.Point(10,40)
$calendar.ShowTodayCircle = $false
$calendar.MaxSelectionCount = 1
$calendar.MaxDate = [System.DateTime]::Today # Set maximum date to today
$calendar.SelectionStart = $defaultDateTime.Date # Set default selection to calculated date
$calendar.add_DateSelected({ Update-Preview })
$form.Controls.Add($calendar)

# Create a calendar END control
$calendar_END = New-Object System.Windows.Forms.MonthCalendar
$calendar_END.Location = New-Object System.Drawing.Point(310,40)
$calendar_END.ShowTodayCircle = $false
$calendar_END.MaxSelectionCount = 1
$calendar_END.MaxDate = [System.DateTime]::Today # Set maximum date to today
$calendar_END.SelectionStart = $defaultDateTime_Today_end.Date # Set default selection to calculated date
$calendar_END.add_DateSelected({ Update-Preview_end })
$form.Controls.Add($calendar_END)

# Create Hour and Minute labels
$hourLabel = New-Object System.Windows.Forms.Label
$hourLabel.Location = New-Object System.Drawing.Point(10,220)
$hourLabel.Size = New-Object System.Drawing.Size(50,20)
$hourLabel.Text = 'Hour'
$form.Controls.Add($hourLabel)

$minuteLabel = New-Object System.Windows.Forms.Label
$minuteLabel.Location = New-Object System.Drawing.Point(110,220)
$minuteLabel.Size = New-Object System.Drawing.Size(60,20)
$minuteLabel.Text = 'Minute'
$form.Controls.Add($minuteLabel)

# END Create Hour and Minute labels
$hourLabel_end = New-Object System.Windows.Forms.Label
$hourLabel_end.Location = New-Object System.Drawing.Point(310,220)
$hourLabel_end.Size = New-Object System.Drawing.Size(50,20)
$hourLabel_end.Text = 'Hour'
$form.Controls.Add($hourLabel_end)

$minuteLabel_end = New-Object System.Windows.Forms.Label
$minuteLabel_end.Location = New-Object System.Drawing.Point(410,220)
$minuteLabel_end.Size = New-Object System.Drawing.Size(60,20)
$minuteLabel_end.Text = 'Minute'
$form.Controls.Add($minuteLabel_end)

# Create NumericUpDown controls for Hour and Minute with default values
$hourPicker = New-Object System.Windows.Forms.NumericUpDown
$hourPicker.Location = New-Object System.Drawing.Point(10,240)
$hourPicker.Size = New-Object System.Drawing.Size(50,20)
$hourPicker.Minimum = 0
$hourPicker.Maximum = 23
$hourPicker.Value = $defaultDateTime.Hour # Set default hour
$form.Controls.Add($hourPicker)

$minutePicker = New-Object System.Windows.Forms.NumericUpDown
$minutePicker.Location = New-Object System.Drawing.Point(110,240)
$minutePicker.Size = New-Object System.Drawing.Size(50,20)
$minutePicker.Minimum = 0
$minutePicker.Maximum = 59
$minutePicker.Value = 0 # Set default minute
$form.Controls.Add($minutePicker)

#END  Create NumericUpDown controls for Hour and Minute with default values
$hourPicker_end = New-Object System.Windows.Forms.NumericUpDown
$hourPicker_end.Location = New-Object System.Drawing.Point(310,240)
$hourPicker_end.Size = New-Object System.Drawing.Size(50,20)
$hourPicker_end.Minimum = 0
$hourPicker_end.Maximum = 23
$hourPicker_end.Value = $defaultDateTime_Today_end.Hour # Set default hour
$form.Controls.Add($hourPicker_end)

$minutePicker_end = New-Object System.Windows.Forms.NumericUpDown
$minutePicker_end.Location = New-Object System.Drawing.Point(410,240)
$minutePicker_end.Size = New-Object System.Drawing.Size(50,20)
$minutePicker_end.Minimum = 0
$minutePicker_end.Maximum = 59
$minutePicker_end.Value = 0 # Set default minute
$form.Controls.Add($minutePicker_end)

# Add event handlers for hour and minute picker value changed
$hourPicker.add_ValueChanged({ Update-Preview })
$minutePicker.add_ValueChanged({ Update-Preview })

$hourPicker_end.add_ValueChanged({ Update-Preview_end })
$minutePicker_end.add_ValueChanged({ Update-Preview_end })

# Create a preview label and initialize it with the default date and time
$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Location = New-Object System.Drawing.Point(10,290)
$previewLabel.Size = New-Object System.Drawing.Size(280,20)
$previewLabel.Text = "Start Date: $($defaultDateTime.ToString('g'))" # Initialize with default date and time
$form.Controls.Add($previewLabel)

# END Create a preview label and initialize it with the default date and time
$previewLabel_end = New-Object System.Windows.Forms.Label
$previewLabel_end.Location = New-Object System.Drawing.Point(310,290)
$previewLabel_end.Size = New-Object System.Drawing.Size(280,20)
$previewLabel_end.Text = "End Date: $($defaultDateTime_Today_end.ToString('g'))" # Initialize with default date and time
$form.Controls.Add($previewLabel_end)

# Add an OK button to confirm the date and time selection
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(300,410) # Adjusted position for better layout
$okButton.Size = New-Object System.Drawing.Size(50,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)



# Update the preview immediately after setting default values
Update-Preview
Update-Preview_end
# Show the form and handle the response

$response = $form.ShowDialog()
if ($response -eq [System.Windows.Forms.DialogResult]::OK)
{
    $selectedDate = $calendar.SelectionStart.Date
    $selectedHour = [int]$hourPicker.Value
    $selectedMinute = [int]$minutePicker.Value
    $data_start = $selectedDate.AddHours($selectedHour).AddMinutes($selectedMinute)

    $selectedDate_end = $calendar_END.SelectionStart.Date
    $selectedHour_end = [int]$hourPicker_end.Value
    $selectedMinute_end = [int]$minutePicker_end.Value
    $data_end = $selectedDate_end.AddHours($selectedHour_end).AddMinutes($selectedMinute_end)


    $top_200_senders = $top200CheckBox.Checked
    $full_list_senders = $fullListCheckBox.Checked

    # Display the selected date and time
    Write-Host "Start Date and Time: $data_start"
    Write-Host "End Date and Time: $data_end"
    Write-Host "Top 200 Short List: $top_200_senders"
    Write-Host "Full List of Senders: $full_list_senders"

        $info=$null
        foreach ($s in $servers)
            {
            $info += Get-MessageTrackingLog -Start (Get-Date $data_start) -End (Get-Date $data_end) -ResultSize Unlimited 
 } 

if($top_200_senders -eq $true){
    Write-Host ""
    Write-Host "exporting the list of top senders to a file $OutFile_Table"

    $list_of_senders=$info| Group-Object -Property Sender | %{ New-Object psobject -Property @{Sender=$_.Name;Recipients=($_.Group | Measure-Object RecipientCount -Sum).Sum}} | Where-Object {$_.Recipients -gt 200} | Sort-Object -Descending Recipients  |Export-Csv $OutFile_Table -NoTypeInformation -Delimiter ";"
    
    $list_of_senders |Select-Object -first 3
    Write-Host ""
    Write-Host "export successful $OutFile_Table"
}

if($full_list_senders -eq $true){

    Write-Host ""
    Write-Host "exporting the list of top senders to a file $OutFile"

    $info | Select-Object @{Name='Recipients'; Expression={$_.Recipients -join ','}},`
    @{Name='RecipientStatus'; Expression={$_.RecipientStatus -join ','}},`
    @{Name='Sender'; Expression={$_.Sender -join ','}}, `
    @{Name='EventData'; Expression={($_.EventData | ForEach-Object {"$($_.Key):$($_.Value)"}) -join ','}},`
    * -ExcludeProperty Recipients,Sender,RecipientStatus,EventData | Export-Csv $outfile -NoTypeInformation -Delimiter ";"

    Write-Host ""
    Write-Host "export of all senders successful $OutFile"
}
}
else
{
    Write-Host "Date and Time selection was canceled."
}