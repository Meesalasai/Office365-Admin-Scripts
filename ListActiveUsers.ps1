# Connect to Microsoft 365 Services
# First, install the MSOnline module if not already installed
# Install-Module MSOnline -Force

# Import the required module
Import-Module MSOnline

# Connect to Office 365
$Credential = Get-Credential
Connect-MsolService -Credential $Credential

# Set CSV file path - modify as needed
$csvFilePath = "C:\Reports\Office365_UserLicenses_$(Get-Date -Format 'yyyyMMdd').csv"

# Retrieve all active users with their license information
Write-Host "Retrieving active users and license information. This may take a few minutes..."
$activeUsers = Get-MsolUser -All | Where-Object {$_.BlockCredential -eq $false}

# Create an array to store user license information
$userLicenseInfo = @()

# Process each user to extract detailed license information
foreach ($user in $activeUsers) {
    # Get basic user information
    $userInfo = [PSCustomObject]@{
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        Department = $user.Department
        IsLicensed = $user.IsLicensed
        Licenses = ""
        LicenseDetails = ""
    }
    
    # If the user has licenses, get detailed license information
    if ($user.IsLicensed) {
        # Get license names
        $licenseNames = @()
        foreach ($license in $user.Licenses) {
            $licenseNames += $license.AccountSkuId
        }
        $userInfo.Licenses = $licenseNames -join "; "
        
        # Get service plan details
        $servicePlans = @()
        foreach ($license in $user.Licenses) {
            foreach ($servicePlan in $license.ServiceStatus) {
                if ($servicePlan.ProvisioningStatus -eq "Success") {
                    $servicePlans += "$($servicePlan.ServicePlan.ServiceName): Enabled"
                }
            }
        }
        $userInfo.LicenseDetails = $servicePlans -join "; "
    }
    else {
        $userInfo.Licenses = "No License Assigned"
        $userInfo.LicenseDetails = "N/A"
    }
    
    # Add the user information to the array
    $userLicenseInfo += $userInfo
}

# Export the results to CSV
Write-Host "Exporting data to CSV file: $csvFilePath"
$userLicenseInfo | Export-Csv -Path $csvFilePath -NoTypeInformation

# Display completion summary
Write-Host "Export complete! Total active users processed: $($userLicenseInfo.Count)"
Write-Host "CSV file saved to: $csvFilePath"
