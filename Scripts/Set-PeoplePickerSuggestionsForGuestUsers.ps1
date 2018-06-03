<#

CSOM is an acronym for Client Side Object Model
CSOM provides a subset of the Server Object Model
CSOM Supports remote execution via JavaScript and .NET
Download Links
SharePoint Online CSOM - http://www.microsoft.com/en-us/download/details.aspx?id=42038
#>

[CmdletBinding()]
Param (
    #security
    [string]$UserName = (Get-Content .\user.txt),
    [string]$UserPassword = (Get-Content .\pw.txt),
    [string]$AdminUrl = "https://x-admin.sharepoint.com",
    #site details
    [string]$SiteUrl = "https://x.sharepoint.com",
    [int]$BatchSize = 5,
    $SiteFilter = '*/teams/*'
)

Try {

    Write-Host "Loading SharePoint CSOM assemblies`n"

    $ClientAssembyPath = Resolve-Path("C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI") -ErrorAction Stop
    $TenantAssembyPath = Resolve-Path("C:\Program Files\SharePoint Client Components\16.0\Assemblies") -ErrorAction Stop
    
    Add-Type -Path ($ClientAssembyPath.Path + "\Microsoft.SharePoint.Client.dll")
    Add-Type -Path ($ClientAssembyPath.Path + "\Microsoft.SharePoint.Client.Runtime.dll")
    Add-Type -Path ($TenantAssembyPath.Path + "\Microsoft.Online.SharePoint.Client.Tenant.dll")
    
} Catch {
    
    Write-Host "Can't load assemblies..." -ForegroundColor Red
    Write-Host $Error[0].Exception.Message -ForegroundColor Red
    Exit
} 

$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($AdminUrl)
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName , ( ConvertTo-SecureString $UserPassword -AsPlainText -Force ))
$Ctx.Credentials = $Credentials
$Ctx.RequestTimeout = -1
$Ctx.PendingRequest.RequestExecutor.RequestKeepAlive = $True
$Ctx.PendingRequest.RequestExecutor.WebRequest.KeepAlive = $False
$Ctx.PendingRequest.RequestExecutor.WebRequest.ProtocolVersion = [System.Net.HttpVersion]::Version10



$Tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($Ctx)

[Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerable]$TenantSiteProperties = $Null


$StartIndex = 0
$TotalSiteCollectionCount = 0
$InScopeSiteCollectionCount = 0
$ModifiedSiteCollectionCount = 0
$MaxSPRequestCalls = 5 #actuall 16 but we stay below
$NumSPRequestCalls = 0
$TotalTime = [System.Diagnostics.Stopwatch]::StartNew()

While (($TenantSiteProperties -eq $Null) -or ($StartIndex -gt 0)) {

    $TenantSiteProperties = $Tenant.GetSiteProperties($StartIndex, $True)
    $Ctx.Load($TenantSiteProperties)
    $Ctx.ExecuteQuery();

    $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Output "-------------------------------------------" 
    ForEach ($Sp in $TenantSiteProperties) {
        
        $TotalSiteCollectionCount++
        
        If ($Sp.Url -like $SiteFilter) {
            $InScopeSiteCollectionCount ++
            $SiteProps = $Tenant.GetSitePropertiesByUrl($SP.Url, $True)
            $Tenant.Context.Load($SiteProps)
            $Tenant.Context.ExecuteQuery()

            If ($SiteProps.ShowPeoplePickerSuggestionsForGuestUsers -eq $False) {

                $numSPRequestCalls = $numSPRequestCalls + 1
                Write-Host "$numSPRequestCalls - $($Sp.Title) => $($Sp.Url)"                           
                
                $Sp.ShowPeoplePickerSuggestionsForGuestUsers = $True
                $ModifiedSiteCollectionCount ++ 
                $Update = $Sp.Update()
                
                If ($numSPRequestCalls -eq $maxSPRequestCalls) {
                    Write-Output "-------------------------------------------" 
                    Write-Output "`nReached request call limit of $numSPRequestCalls. Executing query...."
                    
                    Try {
                  
                        $Ctx.ExecuteQuery()

                    } Catch {

                        Write-Host "Something went wrong" -ForegroundColor Magenta
                        Write-Host $Error[0].Exception -ForegroundColor Magenta
                        
                    } Finally {

                        $StopWatch.Stop()
                        $Output = "Time elapsed for the last execute: {0}s" -f $StopWatch.Elapsed.TotalSeconds
                        Write-Host $Output -ForegroundColor DarkGreen
                        
                       
                        $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

                    }
                    $numSPRequestCalls = 0
                    Write-Output "`n-------------------------------------------" 
                    Start-Sleep -Seconds 1
                }
            }

        }
    }

    $StartIndex = $TenantSiteProperties.NextStartIndex

}


If ($numSPRequestCalls -ne 0) {
    Write-Output "Last execute for $numSPRequestCalls site collections..."
    Try {

        $Ctx.ExecuteQuery() 

    } Catch {

        Write-Host "Something went wrong" -ForegroundColor Magenta
        Write-Host $Error[0].Exception -ForegroundColor Magenta

    } Finally {

        $StopWatch.Stop()
        $Output = "Time elapsed for the last execute: {0}s" -f $StopWatch.Elapsed.TotalSeconds
        
        Write-Host $Output -ForegroundColor DarkGreen


    }
}

$Ctx.Dispose()
$Tenant.Context.Dispose()
$TotalTime.Stop
$Output = "Total elapsed time: {0} minutes" -f $TotalTime.Elapsed.TotalMinutes
Write-Host $Output -ForegroundColor DarkGreen

Write-Host "Total amount of site collections: $TotalSiteCollectionCount"
Write-Host "InScope site collections: $InScopeSiteCollectionCount"
Write-Host "Modified site collections: $ModifiedSiteCollectionCount"

