Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = "$env:TEMP\ADUserUpdateLog_$timestamp.csv"
"d,Timestamp,User,Field,OldValue,NewValue,DryRun" | Out-File -FilePath $logPath -Encoding UTF8

function Log-Update {
    param (
        [string]$User,
        [hashtable]$Changes,
        [bool]$DryRun
    )
    $now = Get-Date -Format 's'
    if ($Changes.Count -eq 0) {
        "$now,$User,None,None,None,None,$DryRun" | Out-File -FilePath $logPath -Append
    } else {
        foreach ($key in $Changes.Keys) {
            $old = $Changes[$key].Old -replace ",", ";"
            $new = $Changes[$key].New -replace ",", ";"
            "$now,$User,$key,$old,$new,$DryRun" | Out-File -FilePath $logPath -Append
        }
    }
}

function Ask-DryRun {
    $result = [System.Windows.Forms.MessageBox]::Show("Run in dry-run mode (no actual AD updates)?", "Dry Run?", "YesNo", "Question")
    return ($result -eq 'Yes')
}

function Show-DetailsForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enter Shared User Details"
    $form.Size = '300,480'
    $form.StartPosition = "CenterScreen"

    $labels = @(
        "Address:", "Phone Number:", "Post Code:",
        "State:", "Suburb:", "Department:", "Job Title:"
    )
    $textboxes = @()

    for ($i = 0; $i -lt $labels.Length; $i++) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labels[$i]
        $label.Location = "10,$(20 + ($i * 40))"
        $label.Size = '120,20'
        $form.Controls.Add($label)

        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = "140,$(20 + ($i * 40))"
        $textbox.Size = '120,20'
        $form.Controls.Add($textbox)

        $textboxes += $textbox
    }

    $submitButton = New-Object System.Windows.Forms.Button
    $submitButton.Text = "Next"
    $submitButton.Location = '90,400'
    $submitButton.Add_Click({
        $form.Tag = $true
        $form.Close()
    })
    $form.Controls.Add($submitButton)

    $form.ShowDialog() | Out-Null
    return if ($form.Tag) { $textboxes | ForEach-Object { $_.Text } } else { $null }
}

function Show-UserSearchForm {
    $input = [System.Windows.Forms.Interaction]::InputBox("Enter name (wildcards like *smith*)", "Search AD Users", "*")
    if (-not $input) { return $null }

    try {
        $users = Get-ADUser -Filter "Name -like '$input'" -Properties SamAccountName, Name, StreetAddress, TelephoneNumber, PostalCode, State, City, Department, Title
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to query AD: $_")
        return $null
    }

    if (-not $users) {
        [System.Windows.Forms.MessageBox]::Show("No users found.")
        return $null
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Users"
    $form.Size = '650,500'
    $form.StartPosition = "CenterScreen"

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = '10,10'
    $panel.Size = '610,400'
    $panel.AutoScroll = $true
    $form.Controls.Add($panel)

    $checkboxes = @()
    $y = 0
    foreach ($user in $users) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = "$($user.Name) [$($user.SamAccountName)]"
        $cb.Tag = $user
        $cb.Location = "10,$y"
        $cb.Width = 580
        $panel.Controls.Add($cb)
        $checkboxes += $cb
        $y += 25
    }

    $nextButton = New-Object System.Windows.Forms.Button
    $nextButton.Text = "Preview Changes"
    $nextButton.Location = '250,420'
    $nextButton.Add_Click({
        $selected = $checkboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag }
        if (-not $selected) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one user.")
            return
        }
        $form.Tag = $selected
        $form.Close()
    })
    $form.Controls.Add($nextButton)

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Show-PreviewForm {
    param ($users, $newProps)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Preview Updates"
    $form.Size = '800,500'
    $form.StartPosition = "CenterScreen"

    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.Location = '10,10'
    $list.Size = '760,400'

    $columns = @("User", "Field", "Old Value", "New Value")
    foreach ($col in $columns) {
        $list.Columns.Add($col, 180)
    }

    foreach ($user in $users) {
        foreach ($key in $newProps.Keys) {
            $newVal = $newProps[$key]
            if (![string]::IsNullOrWhiteSpace($newVal)) {
                $oldVal = $user.$key
                if ($oldVal -ne $newVal) {
                    $item = New-Object System.Windows.Forms.ListViewItem("$($user.SamAccountName)")
                    $item.SubItems.Add($key)
                    $item.SubItems.Add($oldVal)
                    $item.SubItems.Add($newVal)
                    $list.Items.Add($item)
                }
            }
        }
    }

    $form.Controls.Add($list)

    $applyButton = New-Object System.Windows.Forms.Button
    $applyButton.Text = "Apply Changes"
    $applyButton.Location = '320,420'
    $applyButton.Add_Click({
        $form.Tag = $true
        $form.Close()
    })
    $form.Controls.Add($applyButton)

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

# MAIN EXECUTION
$dryRun = Ask-DryRun
$sharedProps = Show-DetailsForm
if (-not $sharedProps) { return }

$props = @{
    StreetAddress   = $sharedProps[0]
    TelephoneNumber = $sharedProps[1]
    PostalCode      = $sharedProps[2]
    State           = $sharedProps[3]
    City            = $sharedProps[4]
    Department      = $sharedProps[5]
    Title           = $sharedProps[6]
}

$selectedUsers = Show-UserSearchForm
if (-not $selectedUsers) { return }

$confirmed = Show-PreviewForm -users $selectedUsers -newProps $props
if (-not $confirmed) { return }

foreach ($user in $selectedUsers) {
    $changes = @{}
    try {
        foreach ($key in $props.Keys) {
            $newVal = $props[$key]
            if (![string]::IsNullOrWhiteSpace($newVal)) {
                $oldVal = $user.$key
                if ($oldVal -ne $newVal) {
                    if (-not $dryRun) {
                        Set-ADUser -Identity $user.SamAccountName -Replace @{ $key = $newVal }
                    }
                    $changes[$key] = @{ Old = $oldVal; New = $newVal }
                }
            }
        }
        Log-Update -User $user.SamAccountName -Changes $changes -DryRun:$dryRun
    } catch {
        $now = Get-Date -Format 's'
        "$now,$($user.SamAccountName),ERROR,$_,$dryRun" | Out-File -FilePath $logPath -Append
    }
}

[System.Windows.Forms.MessageBox]::Show("Update process completed.`nOpening log file...", "Done", 'OK', 'Information')
Start-Process notepad.exe $logPath
