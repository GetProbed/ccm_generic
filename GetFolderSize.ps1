# Define the list of servers
$servers = @("Server1", "Server2", "Server3")

# Define the folder paths you want to query
$folderPaths = @(
    "\\Server1\Path\To\Folder1",
    "\\Server2\Path\To\Folder2"
)

# Create an array to store the results
$results = @()

foreach ($server in $servers) {
    Write-Host "Querying folders on $server"
    
    if (Test-Connection -ComputerName $server -Count 1 -Quiet) {
        foreach ($folderPath in $folderPaths) {
            $query = "SELECT * FROM Win32_Directory WHERE Name='$folderPath'"
            $folder = Get-WmiObject -Query $query -ComputerName $server
            if ($folder) {
                $sizeInBytes = [double]$folder.Size
                $sizeInGB = [Math]::Round($sizeInBytes / 1GB, 2)

                # Create an object to store the result
                $resultObject = New-Object PSObject -Property @{
                    "ServerName" = $server
                    "FolderPath" = $folderPath
                    "SizeGB" = $sizeInGB
                }

                # Add the result object to the results array
                $results += $resultObject

                Write-Host "Server: $server, Folder: $folderPath, Size: $sizeInGB GB"
            } else {
                Write-Host "Unable to retrieve size for folder on $server : $folderPath"
            }
        }
    } else {
        Write-Host "Server $server is not reachable."
    }
}

# Export the results to a CSV file
$results | Export-Csv -Path "FolderSizes.csv" -NoTypeInformation
