# Prompt user for computer name
$computerName = Read-Host "Enter the computer name"

# Prompt for admin credentials
$credential = Get-Credential -Message "Enter the admin credentials for $computerName"

try {
    # Create CIM session using provided credentials
    $cimSession = New-CimSession -ComputerName $computerName -Credential $credential -ErrorAction Stop

    # Query network adapters remotely
    $networkAdapter = Get-NetAdapter -CimSession $cimSession | Where-Object { $_.InterfaceDescription -like '*wifi*' }

    if ($networkAdapter) {
        Write-Host "Wireless network adapter found on $computerName:`n" -ForegroundColor Red
        Write-Host "Adapter: $($networkAdapter.InterfaceDescription)" -ForegroundColor Red
        Write-Host "MAC Address: $($networkAdapter.MacAddress)" -ForegroundColor Red
        Write-Host "Status: $($networkAdapter.Status)" -ForegroundColor Red
    } else {
        Write-Host "No wireless network adapter found on $computerName." -ForegroundColor Green
    }
}
catch {
    Write-Host -ForegroundColor Red "Error occurred while querying network adapters on $computerName: $_"
}
finally {
    # Close the CIM session
    if ($cimSession) {
        Remove-CimSession -CimSession $cimSession
    }
}