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
    Param($address, $port, $timeout=800)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $task = $client.ConnectAsync($address, $port)
        $completed = $task.Wait($timeout)
        return $completed -and $client.Connected
    } catch { return $false }
    finally { $client.Dispose() }
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
    $LoginPage = Invoke-WebRequest -UseBasicParsing -TimeoutSec $Timeout -Uri $pfSenseUrl -SessionVariable Session
    $CsrfToken = $LoginPage.InputFields.FindByName('__csrf_magic').Value

    #Login
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $configFile.Settings.pfSense.Username, (ConvertTo-SecureString -AsPlainText -Force ($configFile.Settings.pfSense.Password))
    $Data = @{
	    __csrf_magic=$CsrfToken;
	    usernamefld=$Credential.GetNetworkCredential().UserName;
	    passwordfld=$Credential.GetNetworkCredential().Password;
	    login='Login';
    }
    $Result = Invoke-WebRequest -UseBasicParsing -TimeoutSec $Timeout -WebSession $Session -Uri $pfSenseUrl -Method Post -Body $Data 
    $CsrfToken = $Result.InputFields.FindByName('__csrf_magic').Value

    #Wake on lan
    $Data = @{
	    __csrf_magic=$CsrfToken;
	    if='lan';
	    mac= $configFile.Settings.Pc.Mac;
	    Submit='Send';
    }
    $Result = Invoke-WebRequest -UseBasicParsing -TimeoutSec $Timeout -WebSession $Session -Uri "${pfSenseUrl}/services_wol.php" -Method Post -Body $Data 
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
    $Result = Invoke-WebRequest -UseBasicParsing -TimeoutSec $Timeout -WebSession $Session -Uri "${pfSenseUrl}/index.php?logout" -Method Post -Body $Data 
    $CsrfToken = $Result.InputFields.FindByName('__csrf_magic').Value
}

# Read config file
$configPath = (Resolve-Path ".\WolClientConfig.xml").Path
If ( -not ( Test-Path -Path $configPath ) )  {
    WaitForKey -Msg $msgTable.errConfigNotFound -Error $true
    Exit 1
}
[xml]$configFile = Get-Content -Path $configPath

# Ensure <Rdp> section exists (automatically update older configs)
if ($null -eq $configFile.Settings.Rdp) {
    $rdpNode = $configFile.CreateElement("Rdp")
    $rdpNode.SetAttribute("FileHash", "")
    $configFile.Settings.AppendChild($rdpNode) | Out-Null
    $configFile.Save($configPath)
}

# Check internet connection (ping www.google.it)
Write-Host -NoNewLine "$($msgTable.checkingInternet)... "
if ( Test-Connection -ComputerName "www.google.it" -Quiet -Count 2 ) {
    Write-Host -ForegroundColor Green $msgTable.strOk
} else {
    Write-Host -ForegroundColor Red $msgTable.strErr
    WaitForKey -Msg $msgTable.errNoInternet -Error $true
    Exit 1
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
        Start-Sleep -Milliseconds 500
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
        Start-Sleep -Milliseconds 500
    }
}
Write-Host -ForegroundColor Green " $($msgTable.strOk)"

$templatePath  = Join-Path (Resolve-Path ".\").Path "connection_template.rdp"
$generatedPath = Join-Path (Resolve-Path ".\").Path "connection_generated.rdp"
$signedPath    = Join-Path (Resolve-Path ".\").Path "connection_signed.rdp"

# Generate RDP file with populated Host and Username from XML settings
$templateContent = Get-Content -Path $templatePath -Raw
$generatedContent = $templateContent -replace "\[username\]", $configFile.Settings.Pc.Username -replace "\[host\]", $configFile.Settings.Pc.Host
Set-Content -Path $generatedPath -Value $generatedContent -Encoding ASCII

# Sign if content changed or signed file is missing
$sha256 = Get-FileHash -Path $generatedPath -Algorithm SHA256
$currentHash = $sha256.Hash
$needsSigning = ($currentHash -ne $configFile.Settings.Rdp.FileHash) -or (-not (Test-Path $signedPath))
if ($needsSigning) {
    WaitForKey -Msg $msgTable.promptSigning

    $absScript = Join-Path (Resolve-Path ".\").Path "RDPSign.ps1"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = "powershell.exe"
    $psi.Arguments       = "-ExecutionPolicy Bypass -File `"$absScript`" -InputFile `"$generatedPath`" -OutputFile `"$signedPath`""
    $psi.Verb            = "runas"
    $psi.UseShellExecute = $true

    $process = [System.Diagnostics.Process]::Start($psi)
    $process.WaitForExit()

    if ($process.ExitCode -ne 0) {
        WaitForKey -Msg $msgTable.errSigning -error $true
        Exit 1
    }

    $configFile.Settings.Rdp.FileHash = $currentHash
    $configFile.Save($configPath)
}

# Launches connection-signed.rdp
Start-Process -FilePath "$env:SystemRoot\system32\mstsc.exe" -ArgumentList $signedPath