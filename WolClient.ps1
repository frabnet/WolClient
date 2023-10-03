#Load localized strings
If ( $psISE ) { Import-LocalizedData -BindingVariable msgTable -FileName "WolClient.psd1" } else { Import-LocalizedData -BindingVariable msgTable }

Function WaitForKey {
    Param ($msg, $error=$false)
    If ($error) {
        Write-Host -ForegroundColor White -BackgroundColor Red $msg
        Write-Host $msgTable.promptExit
    } else {
        Write-Host $msg
    }    
    If ( $psISE ) { Pause } else {  [void][System.Console]::ReadKey($true) }    
}
Function TestTCPPort {
    Param($address, $port, $timeout=500)
    $client = New-Object System.Net.Sockets.TcpClient
    $beginConnect = $client.BeginConnect($address, $port, $null, $null)
    Start-Sleep -Milliseconds $timeout
    $Connected = $client.Connected
    $client.Close()
    Return $Connected
}
Function SendWakeOnLan {
    Write-Host -NoNewline "$($msgTable.sendingWol)... "
    #Ignore self signed cert
    add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    $pfSenseUrl = "https://$($configFile.Settings.pfSense.Host):$($configFile.Settings.pfSense.Port)"
    $Timeout = 10
    #Request homepage to extract csrf token
    $LoginPage = Invoke-WebRequest -TimeoutSec $Timeout -Uri $pfSenseUrl -SessionVariable Session
    $CsrfToken = $LoginPage.InputFields.FindByName('__csrf_magic').Value

    #Login
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $configFile.Settings.pfSense.Username, (ConvertTo-SecureString -AsPlainText -Force ($configFile.Settings.pfSense.Password))
    $Data = @{
	    __csrf_magic=$CsrfToken;
	    usernamefld=$Credential.GetNetworkCredential().UserName;
	    passwordfld=$Credential.GetNetworkCredential().Password;
	    login='Login';
    }
    $Result = Invoke-WebRequest -TimeoutSec $Timeout -WebSession $Session -Uri $pfSenseUrl -Method Post -Body $Data 
    $CsrfToken = $Result.InputFields.FindByName('__csrf_magic').Value

    #Wake on lan
    $Data = @{
	    __csrf_magic=$CsrfToken;
	    if='lan';
	    mac= $configFile.Settings.Pc.Mac;
	    Submit='Send';
    }
    $Result = Invoke-WebRequest -TimeoutSec $Timeout -WebSession $Session -Uri "${pfSenseUrl}/services_wol.php" -Method Post -Body $Data 
    $CsrfToken = $Result.InputFields.FindByName('__csrf_magic').Value

    if ($Result.RawContent -like "*Sent magic packet to*") {
        Write-Host -ForegroundColor Green $msgTable.strOk
    } else {
        Write-Host -ForegroundColor Red $msgTable.strErr
    }

    #Logout
    $Data = @{
        __csrf_magic=$CsrfToken;
        logout=""
    }
    $Result = Invoke-WebRequest -TimeoutSec $Timeout -WebSession $Session -Uri "${pfSenseUrl}/index.php?logout" -Method Post -Body $Data 
    $CsrfToken = $Result.InputFields.FindByName('__csrf_magic').Value
}

# Read config file
If ( -not ( Test-Path -Path .\WolClientConfig.xml ) )  {
    WaitForKey -Msg $msgTable.errConfigNotFound -Error $true
    Exit
}
[xml]$configFile = Get-Content -Path .\WolClientConfig.xml

# Check internet connection (ping www.google.it)
Write-Host -NoNewLine "$($msgTable.checkingInternet)... "
if ( Test-Connection -ComputerName "www.google.it" -Quiet -Count 2 ) {
    Write-Host -ForegroundColor Green $msgTable.strOk
} else {
    Write-Host -ForegroundColor Red $msgTable.strErr
    WaitForKey -Msg $msgTable.errNoInternet -Error $true
    Exit
}

#Check VPN Connection (tcp connect to pfSense:443)
Write-Host -NoNewLine "$($msgTable.checkingVpn)..."
If (-not ( TestTCPPort -address $configFile.Settings.pfSense.Host -port $configFile.Settings.pfSense.Port ) ) {
    # Search for the VPN file
    $VpnFile = $configFile.Settings.Vpn.VpnFile
    if ( $VpnFile -eq "AUTO" ) {$VpnFile = (Get-ChildItem -Path $Env:Userprofile\OpenVPN\config\*.ovpn | Select-Object -First 1).Name }
    # Launches OpenVPN gui with selected VPN profile
    Start-Process -FilePath "$env:programfiles\OpenVPN\bin\openvpn-gui.exe" -ArgumentList "--command connect $VpnFile"
    While ( -not ( TestTCPPort -address $configFile.Settings.pfSense.Host -port $configFile.Settings.pfSense.Port ) ) {        
        Write-Host -NoNewline "."
        Sleep -Milliseconds 500
    }
}
Write-Host -ForegroundColor Green " $($msgTable.strOk)"

#Check if PC is powered on (tcp connect to pc:3389)
Write-Host -NoNewline "$($msgTable.checkingPc)..."
if ( -Not ( TestTCPPort -address $configFile.Settings.Pc.Host -Port 3389 ) ) {
    Write-Host -ForegroundColor Red " $($msgTable.strPcTurnedOff)"
    WaitForKey -Msg $msgTable.promptWol
    SendWakeOnLan
    Write-Host -NoNewLine $msgTable.strWaitingPc
    While ( -Not ( TestTCPPort -address $configFile.Settings.Pc.Host -Port 3389 ) ) {
        Write-Host -NoNewLine "."
        Sleep -Milliseconds 500
    }
}
Write-Host -ForegroundColor Green " $($msgTable.strOk)"

# Start RDP Connection
$RdpContent = "
screen mode id:i:2
session bpp:i:16
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
username:s:$($configFile.Settings.Pc.Username)
connection type:i:3
networkautodetect:i:0
bandwidthautodetect:i:1
displayconnectionbar:i:1
enableworkspacereconnect:i:0
disable wallpaper:i:1
allow font smoothing:i:0
allow desktop composition:i:1
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:$($configFile.Settings.Pc.Host)
audiomode:i:2
redirectprinters:i:0
redirectcomports:i:0
redirectsmartcards:i:0
redirectclipboard:i:1
redirectposdevices:i:0
drivestoredirect:s:
autoreconnection enabled:i:1
authentication level:i:0
prompt for credentials:i:0
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:
gatewayusagemethod:i:4
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:0
promptcredentialonce:i:0
gatewaybrokeringtype:i:0
use redirection server name:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
"
$RdpContent | Out-File "rdp.rdp"
Sleep -Seconds 1
Start-Process -FilePath "$env:SystemRoot\system32\mstsc.exe" -ArgumentList "rdp.rdp"