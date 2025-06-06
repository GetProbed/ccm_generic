Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory

# ===== CONFIG =====
$dryRun = $true
$logPath = "$env:TEMP\ADUserUpdateLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# ===== FUNCTIONS =====
function Select-DomainController {
    try {
        $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $dcList = $forest.Domains | ForEach-Object { $_.DomainControllers } | Select-Object -ExpandProperty Name
        if (-not $dcList) { throw "No domain controllers found" }

        $form = New-Object Windows.Forms.Form
        $form.Text = "Select Domain Controller"
        $form.Size = '400,200'
        $combo = New-Object Windows.Forms.ComboBox
        $combo.DropDownStyle = 'DropDownList'
        $combo.Location = '30,40'
        $combo.Size = '320,30'
        $combo.Items.AddRange($dcList)
        $combo.SelectedIndex = 0
        $form.Controls.Add($combo)
        $okButton = New-Object Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = '150,100'
        $okButton.Add_Click({ $form.Tag = $combo.SelectedItem; $form.Close() })
        $form.Controls.Add($okButton)
        $form.ShowDialog() | Out-Null
        return $form.Tag
    } catch {
        [Windows.Forms.MessageBox]::Show("Error: $_")
        return $null
    }
}

function Log-Change {
    param($User, $Field, $Old, $New, $Status)
    "$User,$Field,$Old,$New,$Status" | Out-File -FilePath $logPath -Append -Encoding UTF8
}

# ===== GUI FORM =====
function Show-UpdateForm {
    param ($dc)

    $form = New-Object Windows.Forms.Form
    $form.Text = "AD User Updater"
    $form.Size = '800,500'
    $form.StartPosition = "CenterScreen"

    # Left Panel - User List
    $userList = New-Object Windows.Forms.CheckedListBox
    $userList.Location = '10,10'
    $userList.Size = '300,400'
    $form.Controls.Add($userList)

    # Search box
    $searchBox = New-Object Windows.Forms.TextBox
    $searchBox.Location = '10,420'
    $searchBox.Size = '200,25'
    $form.Controls.Add($searchBox)

    $searchBtn = New-Object Windows.Forms.Button
    $searchBtn.Text = "Search"
    $searchBtn.Location = '220,420'
    $searchBtn.Add_Click({
        $userList.Items.Clear()
        $users = Get-ADUser -Filter "Name -like '*$($searchBox.Text)*'" -Properties * -Server $dc
        $global:selectedObjects = @()
        foreach ($user in $users) {
            $userList.Items.Add($user, $false)
        }
    })
    $form.Controls.Add($searchBtn)

    # Right Panel - Fields
    $labels = "Address","Phone","Post Code","State","Suburb","Department","Job Title"
    $props  = "StreetAddress","TelephoneNumber","PostalCode","State","City","Department","Title"
    $textBoxes = @{}
    for ($i = 0; $i -lt $labels.Length; $i++) {
        $lbl = New-Object Windows.Forms.Label
        $lbl.Text = $labels[$i]
        $lbl.Location = "330,$(20 + ($i * 35))"
        $lbl.Size = '100,25'
        $form.Controls.Add($lbl)

        $txt = New-Object Windows.Forms.TextBox
        $txt.Location = "450,$(20 + ($i * 35))"
        $txt.Size = '300,25'
        $form.Controls.Add($txt)

        $textBoxes[$props[$i]] = $txt
    }

    # Apply Button
    $applyBtn = New-Object Windows.Forms.Button
    $applyBtn.Text = "Apply"
    $applyBtn.Location = '600,400'
    $applyBtn.Add_Click({
        $sharedProps = @{}
        foreach ($key in $textBoxes.Keys) {
            if ($textBoxes[$key].Text -ne '') {
                $sharedProps[$key] = $textBoxes[$key].Text
            }
        }

        $targets = @()
        foreach ($item in $userList.CheckedItems) {
            $targets += $item
        }

        if (-not $targets) {
            [Windows.Forms.MessageBox]::Show("No users selected.")
            return
        }

        foreach ($user in $targets) {
            $original = Get-ADUser -Identity $user.SamAccountName -Properties * -Server $dc
            foreach ($key in $sharedProps.Keys) {
                $old = $original.$key
                $new = $sharedProps[$key]
                if ($old -ne $new) {
                    try {
                        if (-not $dryRun) {
                            Set-ADUser -Identity $original -Replace @{ $key = $new } -Server $dc
                        }
                        Log-Change -User $user.SamAccountName -Field $key -Old $old -New $new -Status "Success"
                    } catch {
                        Log-Change -User $user.SamAccountName -Field $key -Old $old -New $new -Status "Failed: $_"
                    }
                }
            }
        }

        [Windows.Forms.MessageBox]::Show("Completed. Log: $logPath")
        Start-Process notepad.exe $logPath
        $form.Close()
    })
    $form.Controls.Add($applyBtn)

    $form.ShowDialog()
}

# ===== RUN SCRIPT =====
$dc = Select-DomainController
if (-not $dc) { return }

"User,Field,Old,New,Status" | Out-File -FilePath $logPath -Encoding UTF8

Show-UpdateForm -dc $dc
