<#
.SYNOPSIS
Configures and/or updates Teams External Access Settings for tenant
.PARAMETER path
The path and filename of a comma separated value (.csv) file of domains and agencies. Defaults to c:\temp\agencylist.csv
#>
param (
   [Parameter()]
   [String]$path = "C:\temp\agencylist.csv"
)
Start-Transcript
Install-Module PowerShellGet -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module AzureAD -Scope CurrentUser -AllowClobber
$Acct = Read-Host -Prompt "Enter Microsoft User Name"
Connect-MicrosoftTeams -AccountId $Acct
$CSSession = New-CsOnlineSession
Import-PSSession $CSSession -AllowClobber
Connect-AzureAD
$Agencies = Import-Csv -Path $path
[System.Collections.ArrayList]$AllowList = @()
$AllowList.AddRange($Agencies.Domain)
[System.Collections.ArrayList]$TenantDomains = @()
$TenantDomains.AddRange((Get-AzureADDomain).Name)
(Compare-Object -ReferenceObject $TenantDomains -DifferenceObject $AllowList -ExcludeDifferent -IncludeEqual).Inputobject  | ForEach-Object {$AllowList.Remove($_)}
#region Test-Federation Function
function Test-FederationConfiguration {
   Try {
       $TFConfig = Get-CsTenantFederationConfiguration
   }
   Catch {}
   switch ($TFConfig) {
       {$PSItem.AllowedDomains.Element -eq '<AllowAllKnownDomains xmlns="urn:schema:Microsoft.Rtc.Management.Settings.Edge.2008" />'}
           {
               Write-Output "Specific Allowed Domains aren't present in External Access List - Importing List for first time."
               Import-Agencies -List $AllowList
           }
       {$PSItem.AllowedDomains.AllowedDomain.Count -ge 1}
           {
               Write-Output "At least one allowed domain has been found in the External Access List - Importing List as additions"
               Add-Agency -List $AllowList
           }
       Default {}
   }
}
#endregion Test-Federation Function
#region Import-Agencies Function
function Import-Agencies {
   [CmdletBinding()]
   param (
       [Parameter(Mandatory=$True)]
       [System.Collections.ArrayList]$List
   )
   Set-CsTenantFederationConfiguration -AllowedDomainsAsAList $List
   $Config = Get-CsTenantFederationConfiguration
   Write-Verbose  "Added the following Domains to the External Access list as an allowed domain $($Config | ForEach-Object  {$_.AllowedDomains.AllowedDomain.Domain}) "
}
#endregion Import-Agencies Function
#region Add-Agency Fuction
function Add-Agency {
   [CmdletBinding()]
   param (
       [Parameter(Mandatory=$True)]
       [System.Collections.ArrayList]$List
   )   
   ForEach ($Agency in $List) {
       Write-Verbose "Adding $Agency to External Access List as an allowed domain"
       $AllowedDomain = New-CsEdgeDomainPattern -Domain $Agency
       Set-CsTenantFederationConfiguration -AllowedDomainsAsAList @{Add=$AllowedDomain}        
   }  
}
#endregion Add-Agency Function
Test-FederationConfiguration
$CSSession | Remove-PSSession
Disconnect-AzureAD
Disconnect-MicrosoftTeams
Stop-Transcript