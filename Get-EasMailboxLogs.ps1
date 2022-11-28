<#
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
// Powershell script to enable and collect EAS mailbox logs.
//
//**********************************************************************​
//
// Syntax for running this script:
//
// .\Get-MailboxLog.ps1 -Mailboxes <Mailboxes> -OutputPath <Path for output file> -Interval <minutes> -MailboxLocation Cloud
//
// The following example collects logs for two mailbox located in the cloud every hour:
//
// .\Get-MailboxLog.ps1 -Mailboxes jim,zeke -OutputPath C:\easLog -Interval 60 -MailboxLocation Cloud
//
// The following example collects logs for an on-premises mailbox every 15 minutes:
//
// .\Get-MailboxLog.ps1 -Mailboxes jim -OutputPath C:\easLog -Interval 60 -MailboxLocation OnPremises
//
//**********************************************************************​
#>
param(
    [Parameter(Mandatory=$true, HelpMessage="The mailboxes to perform the search against.")] [string[]] $Mailboxes,
    [Parameter(Mandatory = $false, HelpMessage="The destination path for the mailbox logs.")] [string] $OutputPath,
    [Parameter(Mandatory = $false, HelpMessage="The interval to capture mailbox logs.")] [int] $Interval = 60,
    [parameter(Mandatory=$true, HelpMessage="The location of the mailbox is either OnPremises or Cloud")] [ValidateSet("OnPremises", "Cloud")] [String[]]$MailboxLocation
)

function Get-FolderPath {   
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the location"
    $folderBrowser.SelectedPath = "C:\"
    $folderPath = $folderBrowser.ShowDialog()
    [string]$oPath = $folderBrowser.SelectedPath
    return $oPath
}

function Get-ExchangeVersion {
    # Check Exchange version for later cmdlets
    if ((Test-Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup\")) {
        if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup\").MsiProductMajor -eq 15) {
            $version = "15"
            return $version
        }
    }
    if ((Test-Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup\")) {
        if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup\").MsiProductMajor -eq 14) {
            $version = "14"
            return $version
        }
    }
    if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Exchange\v8.0\Setup\")) {
        if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Exchange\v8.0\Setup\").MsiProductMajor -eq 8) {
            $version = "12"
            return $version
        }
    }
}

Write-Warning "Do not close this window until you are ready to collect the logs."
$WaitInterval = $Interval*60

#region Validate OutputPath
[boolean]$validPath = $false
while($validPath -eq $false) {
    if($OutputPath -like $null) {[string]$OutputPath = (Get-Location).Path}
    else{
        if($OutputPath.Substring($OutputPath.Length-1,1) -ne "\") {$OutputPath = $OutputPath+"\"}
    }
    if(Test-Path -Path $OutputPath) {$validPath = $true}
    else {
        Write-Warning "An invalid path for the output was provided. Please select the location."
        $ScriptPath = Get-FolderPath
    }
}
#endregion

#region Exchange location
$isOnPrem = $false
if($MailboxLocation -eq "Cloud") {
    try {
        $ModuleAvailable = Get-Module -Name ExchangeOnlineManagement -ListAvailable
        if ($null -ne $ModuleAvailable) {
            if(!(Get-Module -Name ExchangeOnlineManagement)) {
                Connect-ExchangeOnline
            }
        }
		else {
            Write-Warning "ExchangeOnlineManagement module is not installed. Please install and try again."
            exit
        }
    }
	catch {
        Write-Warning "Unable to detect ExchangeOnlineManagement moduled."
    }
}
else {
    $version = Get-ExchangeVersion
    $isOnPrem = $true
}
#endregion

# Looping indefinitely...
while(1) {
    #region Enable Mailbox Logging
    # Ensure that mailbox logging is not disabled after 72 hours
    foreach($Mailbox in $Mailboxes) {
        Write-Host "Enabling mailbox logging for $Mailbox." -ForegroundColor Green
        try { 
            Set-CasMailbox $Mailbox -ActiveSyncDebugLogging:$true -ErrorAction Stop -WarningAction SilentlyContinue }
        catch { 
            Write-Host "Error enabling the ActiveSync mailbox log for $Mailbox. This script must run on the version of Exchange where the mailbox is located." -ForegroundColor White -BackgroundColor Red
        }
    }
    #endregion 
    #region Collect mailbox logs
    Write-Host "Next set of logs will be retrieved at" (Get-Date).AddSeconds($WaitInterval) -ForegroundColor Cyan
    Start-Sleep $WaitInterval
    foreach($Mailbox in $Mailboxes) {
        Write-Host "Getting all devices for mailbox:" $Mailbox
        if ($isOnPrem -eq $false -or $version -eq "15") {
            try { $devices = Get-MobileDeviceStatistics -Mailbox $Mailbox }
            catch { Write-Host "Error locating devices for $Mailbox." -ForegroundColor White -BackgroundColor Red }
        }
        else {
            try { $devices = Get-ActiveSyncDeviceStatistics -Mailbox $Mailbox }
            catch { Write-Host "Error locating devices for $Mailbox." -ForegroundColor Red }
        }
        if ($devices -ne $null) {
            foreach($device in $devices) {
                Write-Host "Downloading logs for device: " $device.DeviceFriendlyName $device.DeviceID -ForegroundColor Cyan
                $fileName = $OutputPath + $Mailbox + "_MailboxLog_" + $device.DeviceFriendlyName + "_" + $device.DeviceID + "_" + (Get-Date).Ticks + ".txt"
                if ($isOnPrem -eq $false -or $version -eq "15") {
                    try { Get-MobileDeviceStatistics $device.Identity -GetMailboxLog -ErrorAction SilentlyContinue | select -ExpandProperty MailboxLogReport | Out-File -FilePath $fileName }
                    catch { Write-Host "Unable to retrieve mailbox log for $device.Identity" -ForegroundColor White -BackgroundColor Red }
                }
                if ($isOnPrem -and $version -eq "14") {
                    try { Get-ActiveSyncDeviceStatistics $device.Identity -GetMailboxLog:$true -ErrorAction SilentlyContinue | select -ExpandProperty MailboxLogReport | Out-File -FilePath $fileName }
                    catch { Write-Host "Unable to retrieve mailbox log for $device" -ForegroundColor Yellow }
                }
                if($isOnPrem -and $version -eq "12") { 
                    try {Get-ActiveSyncDeviceStatistics $device.Identity -GetMailboxLog:$true -ErrorAction SilentlyContinue -OutputPath $OutputPath }
                    catch { Write-Host "Unable to retrieve mailbox log for $device" -ForegroundColor Yellow }
                }
            }
        }
        else { Write-Host "No devices found for $Mailbox." -ForegroundColor Yellow }
    }
    #endregion
    Write-Host "Reminder: Do no close this window until you are ready to collect the logs." -ForegroundColor White -BackgroundColor Red    
}
