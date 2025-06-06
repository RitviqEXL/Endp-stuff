Set-ExecutionPolicy bypass

# Run PS as Admin 

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}



# Set the search base and LDAP filter
$searchBase = "DC=corp,DC=exlservice,DC=com"
$filter = "(&(objectCategory=User)(objectClass=Person)(employeeType=vendor))"

# Set the properties to retrieve
$properties = "distinguishedName", "Enabled", "CN", "description", "whencreated", "displayname", "employeetype", "sAMAccountName", "co", "L", "department", "company", "mail", "manager"

# Create the directory searcher object and set the search scope
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$searchBase")
$searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree

# Set the LDAP filter and properties to retrieve
$searcher.Filter = $filter
$searcher.PropertiesToLoad.AddRange($properties)

# Set the page size and size limit
# By default, the FindAll() method returns a maximum of 1000 results. Hence using the PageSize.
$searcher.PageSize = 1000
$searcher.SizeLimit = 0

# Perform the search and retrieve the results
$results = $searcher.FindAll()

# Create a new CSV file to export the results
$outputFile = "C:\temp\V_Dump.csv"

# Loop through the results and write them to the output file
$results | ForEach-Object {
    $props = [ordered]@{
        DistinguishedName = $_.Properties["distinguishedName"][0]
        Enabled = $_.Properties["Enabled"][0]
        CN = $_.Properties["CN"][0]
        Description = $_.Properties["description"][0]
        WhenCreated = $_.Properties["whencreated"][0]
        DisplayName = $_.Properties["displayname"][0]
        EmployeeType = $_.Properties["employeetype"][0]
        SAMAccountName = $_.Properties["sAMAccountName"][0]
        Country = $_.Properties["co"][0]
        City = $_.Properties["L"][0]
        Department = $_.Properties["department"][0]
        Company = $_.Properties["company"][0]
        Email = $_.Properties["mail"][0]
    }

    # Get the manager's SamAccountName if it exists
    if ($_.Properties.Contains("manager")) {
        $managerDN = $_.Properties["manager"][0]
        $managerEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$managerDN")
        if ($managerEntry.Properties.Contains("sAMAccountName")) {
            $managerSAM = $managerEntry.Properties["sAMAccountName"][0]
            $props.Add("ManagerSAM", $managerSAM)
        } else {
            $props.Add("ManagerSAM", "")
        }
    } else {
        $props.Add("ManagerSAM", "")
    }

    New-Object PSObject -Property $props
} | Export-Csv -Path $outputFile -NoTypeInformation
