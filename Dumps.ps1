############################  Fetching the ERP ###############################
# Define the source file path on the network shared drive
$latestDate = Get-Date -Format "dd-MMMM-yy"
$sourceFilePath = "\\fsx01\Global Technology\Unified Service Desk\ERP ACTIVE Users data\Latest ERP\June 2025\$latestDate.xlsx"

# Define the destination folder on your local machine
$destinationFolder = "C:\Temp_1"


# Ensure the destination folder exists; create it if it doesn't
if (!(Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
}

# Define the full destination file path
$destinationFilePath = Join-Path -Path $destinationFolder -ChildPath (Split-Path -Leaf $sourceFilePath)

# Copy the file from the network shared drive to the local folder
Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -Force

# Verify if the file was copied successfully
if (Test-Path -Path $destinationFilePath) {
    Write-Host "File successfully copied to $destinationFilePath"
} else {
    Write-Host "Failed to copy the file."
}

############################## Fetch AD dump ################################## 
$cmdCommand = 'csvde -f AD_Data.csv -r "(&(objectClass=user)(objectCategory=computer))" -l "DN,objectClass,cn,userAccountControl,lastLogonTimestamp,operatingSystem,OperatingSystemVersion,whenChanged,pwdLastSet,whenCreated"'  # Replace 'dir' with your desired CMD command
Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmdCommand" -NoNewWindow -Wait

############################# Fetch vendor dump ###############################
# it will open the vendor file cmds as well - Start-Process - -FilePath "C:\temp\Fetch_VendorDump_Manager.ps1"    
#Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -File 'C:\temp\Fetch_VendorDump_Manager.ps1'" -NoNewWindow -Wait
#Start-Job -ScriptBlock { & 'C:\temp\Fetch_VendorDump_Manager.ps1' }
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/RitviqEXL/Endp-stuff/refs/heads/main/Fetch_VendorDump_Manager.ps1" -OutFile "C:\Temp_1\Fetch_VendorDump_Manager.ps1"
Write-Host "File downloaded successfully to C:\Temp_1\Fetch_VendorDump_Manager.ps1"
Unblock-file -Path "C:\Temp_1\Fetch_VendorDump_Manager.ps1"
Start-Process -FilePath "powershell.exe" -ArgumentList "-File C:\Temp_1\Fetch_VendorDump_Manager.ps1" -Wait -NoNewWindow

########################## Opening All the files ###############################
Start-process -FilePath "C:\Temp_1\AD_Data.csv"
Start-process -FilePath "C:\Temp_1\V_Dump.csv"
Start-Process -FilePath  "C:\Temp_1\$latestDate.xlsx"
