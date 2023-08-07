# Define SCCM Site Code and Distribution Point Server
$SiteCode = "CM1"               # Replace with your SCCM Site Code

# Prompt for Distribution Point server name
$DistributionPoint = Read-Host "Enter the Distribution Point server name"

# Function to remove content from the distribution point
function Remove-SCCMContent {
    param (
        [string]$PackageID,
        [string]$ContentType,
        [string]$PackageState,
        [string]$Action,
        [array]$FailedContent
    )
    try {
        if ($Action -eq "Remove") {
            # Remove the content from the distribution point based on the content type
            switch ($ContentType) {
                "Package" {
                    Invoke-WmiMethod -Namespace "root\SMS\site_$SiteCode" -Class "SMS_Package" -Name "RemoveDP" -ArgumentList "$DistributionPoint", $PackageID
                }
                "Application" {
                    Invoke-WmiMethod -Namespace "root\SMS\site_$SiteCode" -Class "SMS_Application" -Name "RemoveDP" -ArgumentList "$DistributionPoint", $PackageID
                }
                "Image" {
                    Invoke-WmiMethod -Namespace "root\SMS\site_$SiteCode" -Class "SMS_Image" -Name "RemoveDP" -ArgumentList "$DistributionPoint", $PackageID
                }
                "BootImage" {
                    Invoke-WmiMethod -Namespace "root\SMS\site_$SiteCode" -Class "SMS_BootImagePackage" -Name "RemoveDP" -ArgumentList "$DistributionPoint", $PackageID
                }
                "DriverPackage" {
                    Invoke-WmiMethod -Namespace "root\SMS\site_$SiteCode" -Class "SMS_DriverPackage" -Name "RemoveDP" -ArgumentList "$DistributionPoint", $PackageID
                }
                # Add other content types here if needed
                default {
                    Write-Host "Unknown content type: $ContentType. Skipping removal for Package ID: $PackageID" -ForegroundColor Yellow
                    return
                }
            }
            Write-Host "Content with Package ID $PackageID, Type $ContentType, and State $PackageState has been removed successfully from $DistributionPoint."
        } elseif ($Action -eq "Redistribute") {
            # Redistribute the content using WMI method
            $failedPackageIDs = $FailedContent | Where-Object { $_.PackageID -eq $PackageID } | Select-Object -ExpandProperty PackageID
            foreach ($failedPackageID in $failedPackageIDs) {
                # Get the WMI object for the specific package and set .RefreshNow = $true
                $package = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" -Class "SMS_Package" -Filter "PackageID = '$failedPackageID'"
                if ($package) {
                    $package.RefreshNow = $true
                    $package.Put_()
                    Write-Host "Content with Package ID $failedPackageID, Type $ContentType, and State $PackageState has been redistributed successfully on $DistributionPoint."
                }
            }
        } else {
            Write-Host "Invalid action: $Action. Skipping action for Package ID: $PackageID" -ForegroundColor Yellow
            return
        }
    } catch {
        Write-Host "Failed to perform the action for content with Package ID $PackageID, Type $ContentType, and State $PackageState on $DistributionPoint." -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
    }
}

# Retrieve content information from WMI for anything that is not in a success state
$wmiQuery = "SELECT * FROM SMS_PackageStatusDistPointsSummarizer WHERE StateType != 1 AND SiteCode = '$SiteCode' AND ServerNALPath LIKE '%$DistributionPoint%'"
$failedContent = Get-WmiObject -Namespace "root\SMS\site_$SiteCode" -Query $wmiQuery

# Check if there's any content to remove or redistribute
if ($failedContent.Count -eq 0) {
    Write-Host "No content found on $DistributionPoint that is not in a success state."
    return
}

# Create a PowerShell array to hold the content information
$contentList = @()

# Populate the contentList array with content information
foreach ($content in $failedContent) {
    $state = @{
        1 = "Success"
        2 = "In Progress"
        3 = "Waiting"
        4 = "Failed"
        # Add other state mappings here if needed
    }

    $contentInfo = [PSCustomObject]@{
        PackageID = $content.PackageID
        Name = $content.PackageName
        Type = $content.PackageType
        State = $state[$content.StateType]
        SourcePath = $content.SourceNALPath
    }
    $contentList += $contentInfo
}

# Display the content list in a GUI using Out-GridView
$selectedContent = $contentList | Out-GridView -Title "Select Content to Perform Action" -OutputMode Multiple -PassThru -Columns PackageID, Name, Type, State, SourcePath

# Confirm the action for selected content
if ($selectedContent) {
    $actionOptions = @("Remove", "Redistribute")
    $selectedAction = $actionOptions | Out-GridView -Title "Select Action" -OutputMode Single -PassThru -Default "Remove"

    foreach ($content in $selectedContent) {
        $packageID = $content.PackageID
        $contentType = $content.Type
        $packageState = $content.State
        if ($selectedAction) {
            Write-Host "Performing action '$selectedAction' for content with Package ID: $packageID, Type: $contentType, and State: $packageState"
            Remove-SCCMContent -PackageID $packageID -ContentType $contentType -PackageState $packageState -Action $selectedAction -FailedContent $failedContent
        }
    }
}
