# Install required modules if not already installed
Install-Module -Name MSOnline -Force
Install-Module -Name ExchangeOnlineManagement -Force

# 1.Connect to services
$O365Cred = Get-Credential
Connect-MsolService -Credential $O365Cred
Connect-ExchangeOnline -Credential $O365Cred



# Create the user accounts with a loop (using mock data):




# 2.Sample user data array
$newUsers = @(
    @{FirstName="John"; LastName="Smith"; DisplayName="John Smith"; UserPrincipalName="jsmith@contoso.com"; Password="P@ssw0rd123"; Department="Sales"},
    @{FirstName="Sarah"; LastName="Jones"; DisplayName="Sarah Jones"; UserPrincipalName="sjones@contoso.com"; Password="P@ssw0rd123"; Department="Sales"},
    @{FirstName="Michael"; LastName="Brown"; DisplayName="Michael Brown"; UserPrincipalName="mbrown@contoso.com"; Password="P@ssw0rd123"; Department="Sales"},
    @{FirstName="Emily"; LastName="Davis"; DisplayName="Emily Davis"; UserPrincipalName="edavis@contoso.com"; Password="P@ssw0rd123"; Department="Sales"},
    @{FirstName="David"; LastName="Wilson"; DisplayName="David Wilson"; UserPrincipalName="dwilson@contoso.com"; Password="P@ssw0rd123"; Department="Sales"}
)

# Create each user
foreach ($user in $newUsers) {
    # Create the user account
    New-MsolUser -FirstName $user.FirstName -LastName $user.LastName -DisplayName $user.DisplayName `
                -UserPrincipalName $user.UserPrincipalName -Password $user.Password -ForceChangePassword $false `
                -UsageLocation "US" -Department $user.Department
}



#   3. Assign licenses to all created users:


# Get the SKU ID for Microsoft 365 Business Standard
$licenseSKU = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -like "*BUSINESSSTANDARD*"} | Select-Object -ExpandProperty AccountSkuId

# Assign licenses to each user
foreach ($user in $newUsers) {
    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses $licenseSKU
}



#  4. Add all users to the Sales distribution group:

powershell
# Add each user to the Sales distribution group
foreach ($user in $newUsers) {
    Add-DistributionGroupMember -Identity "Sales@contoso.com" -Member $user.UserPrincipalName
}
