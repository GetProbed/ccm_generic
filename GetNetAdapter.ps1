# Prompt user for computer name
$computerName = Read-Host "Enter the computer name"

# Function to check if the computer is reachable
function Test-ComputerReachable {
    param (
        [string]$ComputerName
    )

    $pingResult = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet
    return $pingResult
}

# Check if the computer is reachable
if (Test-ComputerReachable -ComputerName $computerName) {
    # Prompt for admin credentials
    $credential = Get-Credential -Message "Enter the admin credentials for $computerName"

    try {
        # Create CIM session using provided credentials
        $cimSession = New-CimSession -ComputerName $computerName -Credential $credential -ErrorAction Stop

        # Query network adapters remotely
        $networkAdapter = Get-NetAdapter -CimSession $cimSession | Where-Object { $_.InterfaceDescription -like '*wifi*' }

        if ($networkAdapter) {
            Write-Host "Wireless network adapter found on $computerName :`n"
            Write-Host "Adapter: $($networkAdapter.InterfaceDescription)"
            Write-Host "MAC Address: $($networkAdapter.MacAddress)"
            Write-Host "Status: $($networkAdapter.Status)"
        } else {
            Write-Host "No wireless network adapter found on $computerName."
        }
    }
    catch {
        Write-Host "Error occurred while querying network adapters on $computerName : $_"
    }
    finally {
        # Close the CIM session
        if ($cimSession) {
            Remove-CimSession -CimSession $cimSession
        }
    }
} else {
    Write-Host "Computer $computerName is not reachable."
}
