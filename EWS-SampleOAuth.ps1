<#
//***********************************************************************
//
// EWS-SampleOAuth.ps1
// Modified 10 October 2022
// Last Modifier:  Jim Martin
// Project Owner:  Jim Martin
// Version: v202210101300
//
//***********************************************************************
//
// Copyright (c) 2018 Microsoft Corporation. All rights reserved.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//**********************************************************************​
//
<#
    .SYNOPSIS
        This script provides sample for using OAuth to access a mailbox via EWS
    .DESCRIPTION
       This script performs a search against a mailbox based on provided subject.
            Note: Only searches against well known folder names
    .PARAMETER MailboxName
        Specifies the mailbox to connect
    .PARAMETER FolderName
        Specifies the well known folder name to search
    .PARAMETER TenantId
        Specifies the Exchange Online tenant where mailbox resides
    .PARAMETER Subject
        Specifies the Subject for the search
    .PARAMETER UseImpersonation
        Specifies whether or not EWS impersonation is used
    .PARAMETER EnableLogging
        Specifies whether or not to enable EWS trace logging for troubleshooting. Trace log will be written in current directory.

    .EXAMPLE
		.\EWS-SampleOAuth.ps1 -MailboxName jmartin@contoso.com -FolderName Inbox -TenantId contoso.onmicrosoft.com -Subject "Test message" -UseImpersonation:$False
		This will perform a search against the Inbox folderfor the subject "Test message" using the authenticated user account.

    .EXAMPLE
		.\EWS-SampleOAuth.ps1 -MailboxName jmartin@contoso.com -FolderName Calendar -TenantId contoso.onmicrosoft.com -Subject "Confidential" -EnableLogging:$True
		This will perform a search against the Calendar folder for the subject "Confidential" and create a trace log for troubleshooting purposes.
#>

param(
    [Parameter(Mandatory = $true)] [string] $MailboxName,
    [Parameter(Mandatory = $true)] [string] $FolderName,
    [Parameter(Mandatory = $true)] [string] $TenantId,
    [Parameter(Mandatory = $true)] [string] $Subject,
    [Parameter(Mandatory = $false)] [boolean] $UseImpersonation=$false,
    [Parameter(Mandatory = $false)] [boolean] $EnableLogging=$false
)

function Enable-TraceHandler(){
$sourceCode = @"
    public class ewsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
    {
        public System.String LogFile {get;set;}
        public void Trace(System.String traceType, System.String traceMessage)
        {
            System.IO.File.AppendAllText(this.LogFile, traceMessage);
        }
    }
"@    

    Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $ewsDLL
    $TraceListener = New-Object ewsTraceListener
   return $TraceListener
}

#region Disclaimer
Write-Host -ForegroundColor Yellow '//***********************************************************************'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// Copyright (c) 2018 Microsoft Corporation. All rights reserved.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR'
Write-Host -ForegroundColor Yellow '// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,'
Write-Host -ForegroundColor Yellow '// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE'
Write-Host -ForegroundColor Yellow '// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER'
Write-Host -ForegroundColor Yellow '// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,'
Write-Host -ForegroundColor Yellow '// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '//**********************************************************************​'
#endregion

#region GetOAuthToken
#Check and install Microsoft Authentication Library module
if(!(Get-Module -Name MSAL.PS -ListAvailable -ErrorAction Ignore)){
    try { 
        #Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module -Name MSAL.PS -Repository PSGallery -Force
    }
    catch {
        Write-Warning "Failed to install the Microsoft Authentication Library module."
        exit
    }
    try {
        Import-Module -Name MSAL.PS
    }
    catch {
        Write-Warning "Failed to import the Microsoft Authentication Library module."
    }
}

$ClientID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
$RedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
$RedirectUri = "ms-appx-web://Microsoft.AAD.BrokerPlugin/d3590ed6-52b3-4102-aeff-aad2292ab01c"

$Token = Get-MsalToken -ClientId $ClientID -RedirectUri $RedirectUri  -TenantId $Authority -Scopes 'https://outlook.office365.com/EWS.AccessAsUser.All' -Interactive
$OAuthToken = "Bearer {0}" -f $Token.AccessToken
#endregion

#region LoadEwsManagedAPI
#Check for EWS Managed API, exit if missing
$ewsDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $ewsDLL) {
        Import-Module $ewsDLL
    }
else {
        Write-Warning "This script requires the EWS Managed API 1.2 or later."
        exit
    }
#endregion

#region EwsService
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013 
## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
$service.UserAgent = "EwsPowerShellScript"
$service.Url = "https://outlook.office365.com/ews/exchange.asmx"
$service.HttpHeaders.Clear()
$service.HttpHeaders.Add("Authorization", " $($OAuthToken)")
$service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
if($UseImpersonation) {
    $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
}
#Enable trace logging for troublshooting
if($EnableLogging) {
    Write-Host "EWS trace logging enabled" -ForegroundColor Cyan
    $service.TraceEnabled = $True
    $TraceHandlerObj = Enable-TraceHandler
    $OutputPath = Get-Location
    $TraceHandlerObj.LogFile = "$OutputPath\$MailboxName-$FolderName.log"
    $service.TraceListener = $TraceHandlerObj
}
#endregion

#region ConnectToFolder
Write-Host "Connecting to the $FolderName for $MailboxName..." -ForegroundColor Cyan
$folderid= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$FolderName,$MailboxName)     
$MailboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
#endregion

#region SearchMailboxFolder
Write-Host "Searching the $FolderName of $MailboxName for the subject `'$Subject`'..." -ForegroundColor Cyan
#Define ItemView to retrive just 1000 Items      
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)     
$rptCollection = @()  
$fiItems = $null      
do{   
    $psPropset= New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)    
    $fiItems = $service.FindItems($MailboxFolder.Id,"Subject:$Subject",$ivItemView)     
    if($fiItems.Items.Count -gt 0){  
        [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)    
        foreach($Item in $fiItems.Items){   
            $MailboxObj = New-Object PSObject -Property @{ Mailbox=$MailboxName; Subject=$Item.Subject; From=$Item.From; DateTimeReceived=$Item.DateTimeReceived; Size=$Item.Size; DateTimeSent=$Item.DateTimeSent; HasAttachments=$Item.HasAttachments; MessageClass=$Item.ItemClass; };
            Write-Output $MailboxObj
        }  
    }  
    $ivItemView.Offset += $fiItems.Items.Count
}
while($fiItems.MoreAvailable -eq $true)
#endregion