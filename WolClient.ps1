function WaitForKey {
    param (
        [String] $Msg
    )
    Write-Host $Msg 
    If ( $psISE ) { Pause } else { [void][System.Console]::ReadKey($FALSE) }    
}

If ( -not ( Test-Path -Path .\WolClientConfig.xml ) )  {
    Write-Host "Errore: WolClientConfig.xml non trovato. Adattare e rinominare WolClientConfig_Sample.xml"
    WaitForKey -Msg "Premere un tasto per uscire."
    Exit
}


###########################
# FASE1 Verifica internet #
###########################
Write-Host -NoNewLine "Verifica connessione internet... "
if ( Test-Connection -ComputerName "www.google.it" -Quiet -Count 2 ) {
    Write-Host -ForegroundColor Green "Ok"
} else {
    Write-Host -ForegroundColor Green "Errore"
    Write-Host "Controllare connessione internet e riprovare."
    WaitForKey -Msg "Premere un tasto per uscire."
    Exit
}


######################
# FASE2 Verifica VPN #
######################

Write-Host -NoNewLine "Verifica connessione VPN..."
If (-not ( Test-Connection -ComputerName $configFile.Settings.pfSense.Host -Quiet -Count 2 )) {
    #Ricerca file OVPN
    $VpnFile = $configFile.Settings.Vpn.VpnFile
    if ( $VpnFile -eq "AUTO" ) {$VpnFile = (Get-ChildItem -Path $Env:Userprofile\OpenVPN\config\*.ovpn | Select-Object -First 1).Name }
    #Avvio client OpenVPN    
    Start-Process -FilePath "$env:programfiles\OpenVPN\bin\openvpn-gui.exe" -ArgumentList "--command connect $VpnFile"
    While ( -not ( Test-Connection -ComputerName $configFile.Settings.pfSense.Host -Quiet -Count 2 ) ) {        
        Write-Host -NoNewline "."
        Sleep -Milliseconds 500
    }
}
Write-Host -ForegroundColor Green " Ok"


#####################
# FASE3 Verifica PC #
#####################
function TestConn([string]$srv,$port=135,$timeout=3000,[switch]$verbose){
    #Alternativa a
    #Test-NetConnection -ComputerName $computer -Port $rdpport -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationLevel Quiet
    #con timeout configurabile https://web.archive.org/web/20150405035615/http://poshcode.org/85
    $ErrorActionPreference = "SilentlyContinue"
    $tcpclient = new-Object system.Net.Sockets.TcpClient
    $iar = $tcpclient.BeginConnect($srv,$port,$null,$null)
    $wait = $iar.AsyncWaitHandle.WaitOne($timeout,$false)
    if(!$wait) {
        $tcpclient.Close()
        if($verbose){Write-Host "Connection Timeout"}
        Return $false
    } else {
    $error.Clear()
    $tcpclient.EndConnect($iar) | out-Null
    if(!$?){
        if($verbose){write-host $error[0]};$failed = $true}
        $tcpclient.Close()
    }
    if($failed){return $false}else{return $true}
}


Function SendWakeOnLan {
Write-Host -NoNewline "Invio comando accensione: "
#Ignora errore certificato self-signed
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

$Timeout = 10
#Pagina iniziale (per token csrf)
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
$pfSenseUrl = "https://${pfSenseIp}"
$Data = @{
	__csrf_magic=$CsrfToken;
	if='lan';
	mac= $configFile.Settings.Pc.Mac;
	Submit='Send';
}
$Result = Invoke-WebRequest -TimeoutSec $Timeout -WebSession $Session -Uri "${pfSenseUrl}/services_wol.php" -Method Post -Body $Data 
$CsrfToken = $Result.InputFields.FindByName('__csrf_magic').Value

if ($Result.RawContent -like "*Sent magic packet to*") {
    Write-Host -ForegroundColor Green "Ok"
} else {
    Write-Host -ForegroundColor Red "Errore"
}

#Logout
$Data = @{
    __csrf_magic=$CsrfToken;
    logout=""
}
$Result = Invoke-WebRequest -TimeoutSec $Timeout -WebSession $Session -Uri "${pfSenseUrl}/index.php?logout" -Method Post -Body $Data 
$CsrfToken = $Result.InputFields.FindByName('__csrf_magic').Value
}

Write-Host -NoNewline "Verifica accensione PC..."
if ( -Not ( TestConn -Srv $configFile.Settings.Pc.Host -Port 3389 -Timeout 600 ) ) {
    Write-Host -ForegroundColor Red "Spento"
    WaitForKey -Msg "Premere un tasto per accendere il computer."
    SendWakeOnLan
    Write-Host -NoNewLine "Attesa PC"
    While ( -Not ( TestConn -Srv $configFile.Settings.Pc.Host -Port 3389 -Timeout 600 ) ) {
        Write-Host -NoNewLine "."
        Sleep -Seconds 1
    }
}
Write-Host -ForegroundColor Green " Ok"


###################
# FASE4 Avvio RDP #
###################

$RdpContent = "
screen mode id:i:2
session bpp:i:16
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
username:s:${PcUser}
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
full address:s:${PcIp}
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
Start-Process -FilePath "$env:SystemRoot\system32\mstsc.exe" -ArgumentList "rdp.rdp"