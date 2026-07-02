#Requires -RunAsAdministrator
Param(
    [string] $InputFile  = ".\connection.rdp",
    [string] $OutputFile = "",
    [string] $CAName     = "WolClientRDPPublisher",
    [int]    $Years      = 5
)

# Microsoft documents rdpsign.exe, but not the internal format of RDP signatures.
# This script reproduces the same embedded RDP signature format using only .NET.
# It creates a detached CMS/PKCS#7 signature over the secure RDP settings,
# prepends the 12-byte RDP signature header, and writes the resulting Base64
# payload to the signature:s: field in 64-character chunks separated by
# double spaces, matching rdpsign.exe output.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Security

# Complete list of RDP secure settings that can be signed
# Order is important and matches Microsoft's rdpsign.exe behavior
$secureSettings = @(
    @{ Prefix = "full address:s:";                     Name = "Full Address" },
    @{ Prefix = "alternate full address:s:";           Name = "Alternate Full Address" },
    @{ Prefix = "pcb:s:";                              Name = "PCB" },
    @{ Prefix = "use redirection server name:i:";      Name = "Use Redirection Server Name" },
    @{ Prefix = "server port:i:";                      Name = "Server Port" },
    @{ Prefix = "negotiate security layer:i:";         Name = "Negotiate Security Layer" },
    @{ Prefix = "enablecredsspsupport:i:";             Name = "EnableCredSspSupport" },
    @{ Prefix = "disableconnectionsharing:i:";         Name = "DisableConnectionSharing" },
    @{ Prefix = "autoreconnection enabled:i:";         Name = "AutoReconnection Enabled" },
    @{ Prefix = "gatewayhostname:s:";                  Name = "GatewayHostname" },
    @{ Prefix = "gatewayusagemethod:i:";               Name = "GatewayUsageMethod" },
    @{ Prefix = "gatewayprofileusagemethod:i:";        Name = "GatewayProfileUsageMethod" },
    @{ Prefix = "gatewaycredentialssource:i:";         Name = "GatewayCredentialsSource" },
    @{ Prefix = "support url:s:";                      Name = "Support URL" },
    @{ Prefix = "promptcredentialonce:i:";             Name = "PromptCredentialOnce" },
    @{ Prefix = "gatewaybrokeringtype:i:";             Name = "GatewayBrokeringType" },
    @{ Prefix = "require pre-authentication:i:";       Name = "Require pre-authentication" },
    @{ Prefix = "pre-authentication server address:s:";Name = "Pre-authentication server address" },
    @{ Prefix = "alternate shell:s:";                  Name = "Alternate Shell" },
    @{ Prefix = "shell working directory:s:";          Name = "Shell Working Directory" },
    @{ Prefix = "remoteapplicationprogram:s:";         Name = "RemoteApplicationProgram" },
    @{ Prefix = "remoteapplicationexpandworkingdir:s:";Name = "RemoteApplicationExpandWorkingdir" },
    @{ Prefix = "remoteapplicationmode:i:";            Name = "RemoteApplicationMode" },
    @{ Prefix = "remoteapplicationguid:s:";            Name = "RemoteApplicationGuid" },
    @{ Prefix = "remoteapplicationname:s:";            Name = "RemoteApplicationName" },
    @{ Prefix = "remoteapplicationicon:s:";            Name = "RemoteApplicationIcon" },
    @{ Prefix = "remoteapplicationfile:s:";            Name = "RemoteApplicationFile" },
    @{ Prefix = "remoteapplicationfileextensions:s:";  Name = "RemoteApplicationFileExtensions" },
    @{ Prefix = "remoteapplicationcmdline:s:";         Name = "RemoteApplicationCmdLine" },
    @{ Prefix = "remoteapplicationexpandcmdline:s:";   Name = "RemoteApplicationExpandCmdLine" },
    @{ Prefix = "prompt for credentials:i:";           Name = "Prompt For Credentials" },
    @{ Prefix = "authentication level:i:";             Name = "Authentication Level" },
    @{ Prefix = "audiomode:i:";                        Name = "AudioMode" },
    @{ Prefix = "redirectdrives:i:";                   Name = "RedirectDrives" },
    @{ Prefix = "redirectprinters:i:";                 Name = "RedirectPrinters" },
    @{ Prefix = "redirectcomports:i:";                 Name = "RedirectCOMPorts" },
    @{ Prefix = "redirectsmartcards:i:";               Name = "RedirectSmartCards" },
    @{ Prefix = "redirectposdevices:i:";               Name = "RedirectPOSDevices" },
    @{ Prefix = "redirectclipboard:i:";                Name = "RedirectClipboard" },
    @{ Prefix = "devicestoredirect:s:";                Name = "DevicesToRedirect" },
    @{ Prefix = "drivestoredirect:s:";                 Name = "DrivesToRedirect" },
    @{ Prefix = "loadbalanceinfo:s:";                  Name = "LoadBalanceInfo" },
    @{ Prefix = "redirectdirectx:i:";                  Name = "RedirectDirectX" },
    @{ Prefix = "rdgiskdcproxy:i:";                    Name = "RDGIsKDCProxy" },
    @{ Prefix = "kdcproxyname:s:";                     Name = "KDCProxyName" },
    @{ Prefix = "redirectlocation:i:";                 Name = "RedirectLocation" },
    @{ Prefix = "redirectwebauthn:i:";                 Name = "RedirectWebAuthn" },
    @{ Prefix = "enablerdsaadauth:i:";                 Name = "EnableRdsAadAuth" },
    @{ Prefix = "eventloguploadaddress:s:";            Name = "EventLogUploadAddress" }
)

function Test-LineStartsWith {
    Param(
        [Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Line,
        [Parameter(Mandatory = $true)] [string] $Prefix
    )

    return $Line.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Read-RdpLines {
    Param([Parameter(Mandatory = $true)] [string] $Path)

    # Read with BOM detection to handle both UTF-8 and UTF-16 LE files
    $reader = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
    try {
        $content = $reader.ReadToEnd()
    } finally {
        $reader.Close()
    }

    # Split on any line ending type (Windows, Unix, Mac)
    $lines = [System.Text.RegularExpressions.Regex]::Split($content, "\r\n|\n|\r")
    
    # Remove trailing empty line if present
    if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq "") {
        if ($lines.Count -eq 1) {
            return @()
        }
        $lines = $lines[0..($lines.Count - 2)]
    }

    # Trim trailing whitespace from each line
    return @($lines | ForEach-Object { $_.TrimEnd() })
}

function Format-RdpSignatureLine {
    Param([Parameter(Mandatory = $true)] [string] $SignatureValue)

    $width = 64
    $chunks = New-Object System.Collections.Generic.List[string]
    
    for ($offset = 0; $offset -lt $SignatureValue.Length; $offset += $width) {
        $length = [Math]::Min($width, $SignatureValue.Length - $offset)
        $chunks.Add($SignatureValue.Substring($offset, $length))
    }

    # Join with double spaces and add trailing double spaces (rdpsign.exe format)
    return "signature:s:" + ($chunks.ToArray() -join "  ") + "  "
}

function Add-CertificateToStore {
    Param(
        [Parameter(Mandatory = $true)] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate,
        [Parameter(Mandatory = $true)] [string] $StoreName
    )

    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreName, "LocalMachine")
    try {
        $store.Open("ReadWrite")
        $alreadyPresent = $store.Certificates | Where-Object { $_.Thumbprint -eq $Certificate.Thumbprint } | Select-Object -First 1
        if (-not $alreadyPresent) {
            $store.Add($Certificate)
            Write-Verbose "Added certificate to $StoreName store"
        } else {
            Write-Verbose "Certificate already present in $StoreName store"
        }
    } finally {
        $store.Close()
    }
}

function Add-TrustedRdpPublisherThumbprint {
    Param([Parameter(Mandatory = $true)] [string] $Thumbprint)

    $gpoPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
    $keyName = "TrustedCertThumbprints"
    
    if (-not (Test-Path -LiteralPath $gpoPath)) {
        New-Item -Path $gpoPath -Force | Out-Null
        Write-Verbose "Created registry path: $gpoPath"
    }

    $registryKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
        "SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services",
        $true
    )
    
    if ($null -eq $registryKey) {
        throw "Unable to write key"
    }

    $existingValue = $registryKey.GetValue($keyName, "")
    $existingKind = $null
    
    try {
        $existingKind = $registryKey.GetValueKind($keyName)
    } catch {
        $existingKind = [Microsoft.Win32.RegistryValueKind]::String
    }

    # Handle MultiString registry type
    if ($existingKind -eq [Microsoft.Win32.RegistryValueKind]::MultiString) {
        $values = @($existingValue)
        if ($values -notcontains $Thumbprint) {
            $registryKey.SetValue($keyName, @($values + $Thumbprint), $existingKind)
            Write-Verbose "Added thumbprint to TrustedCertThumbprints (MultiString)"
        }
        return
    }

    # Handle String or REG_SZ registry type (comma-separated)
    $textValue = [string] $existingValue
    if ($textValue -notlike "*$Thumbprint*") {
        $newValue = if ([string]::IsNullOrWhiteSpace($textValue)) { $Thumbprint } else { "$textValue,$Thumbprint" }
        $registryKey.SetValue($keyName, $newValue, $existingKind)
        Write-Verbose "Added thumbprint to TrustedCertThumbprints (String)"
    } else {
        Write-Verbose "Thumbprint already in TrustedCertThumbprints"
    }
}

function New-RdpSignatureValue {
    Param(
        [Parameter(Mandatory = $true)] [string[]] $SignLines,
        [Parameter(Mandatory = $true)] [string[]] $SignNames,
        [Parameter(Mandatory = $true)] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate
    )

    # Construct the message to be signed: 
    # 1. All secure settings lines (CRLF separated)
    # 2. The signscope line
    # 3. CRLF
    # 4. NULL terminator
    # All encoded as UTF-16 LE (Unicode)
    $messageText = (($SignLines -join "`r`n") + "`r`n" + "signscope:s:" + ($SignNames -join ",") + "`r`n" + [char]0)
    $messageBytes = [System.Text.Encoding]::Unicode.GetBytes($messageText)

    # Create a detached CMS/PKCS#7 signature
    $contentInfo = New-Object System.Security.Cryptography.Pkcs.ContentInfo -ArgumentList @(,$messageBytes)
    $signedCms = New-Object System.Security.Cryptography.Pkcs.SignedCms -ArgumentList $contentInfo, $true
    $cmsSigner = New-Object System.Security.Cryptography.Pkcs.CmsSigner -ArgumentList `
        ([System.Security.Cryptography.Pkcs.SubjectIdentifierType]::IssuerAndSerialNumber), $Certificate

    # Configure signer: include certificate, use SHA-256
    $cmsSigner.IncludeOption = [System.Security.Cryptography.X509Certificates.X509IncludeOption]::EndCertOnly
    $cmsSigner.DigestAlgorithm = New-Object System.Security.Cryptography.Oid("2.16.840.1.101.3.4.2.1", "sha256")
    
    # Compute the signature
    $signedCms.ComputeSignature($cmsSigner, $true)
    $cmsBytes = $signedCms.Encode()

    # Build RDP signature format:
    # 4 bytes: version (0x00010001)
    # 4 bytes: signature count (0x00000001)
    # 4 bytes: CMS blob length
    # N bytes: CMS blob
    $memory = New-Object System.IO.MemoryStream
    try {
        $writer = New-Object System.IO.BinaryWriter($memory)
        $writer.Write([UInt32] 0x00010001)  # Version
        $writer.Write([UInt32] 0x00000001)  # Count (always 1)
        $writer.Write([UInt32] $cmsBytes.Length)  # Length of CMS data
        $writer.Write($cmsBytes)  # CMS/PKCS#7 signature
        $writer.Flush()
        
        # Convert to Base64
        return [Convert]::ToBase64String($memory.ToArray())
    } finally {
        $memory.Close()
    }
}

# ==============================================================================
# Main Script
# ==============================================================================

Write-Host "RDP Sign Tool (PowerShell/.NET Implementation)" -ForegroundColor Cyan
Write-Host "No rdpsign.exe required - works on Windows 11 Home!" -ForegroundColor Cyan
Write-Host ""

# Verify input file exists
if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    Write-Error "RDP file not found: $InputFile"
    Exit 1
}
$InputFile = (Resolve-Path -LiteralPath $InputFile).Path
Write-Host "Input file : $InputFile" -ForegroundColor Gray

# Determine output file path
if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $dir = [System.IO.Path]::GetDirectoryName($InputFile)
    if ([string]::IsNullOrEmpty($dir)) { $dir = "." }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $OutputFile = Join-Path $dir ("$base-signed.rdp")
}
Write-Host "Output file: $OutputFile" -ForegroundColor Gray
Write-Host ""

# Certificate: find existing or create new
Write-Host "Checking certificate..." -ForegroundColor Yellow
$certSubject = "CN=$CAName"
$cert = Get-ChildItem Cert:\LocalMachine\My |
    Where-Object { $_.Subject -eq $certSubject -and $_.HasPrivateKey } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (-not $cert) {
    Write-Host "Creating new self-signed certificate..." -ForegroundColor Yellow
    $cert = New-SelfSignedCertificate `
        -Subject $certSubject `
        -CertStoreLocation "Cert:\LocalMachine\My" `
        -Type CodeSigningCert `
        -KeyExportPolicy Exportable `
        -NotAfter (Get-Date).AddYears($Years)
    Write-Host "Certificate created successfully" -ForegroundColor Green
} else {
    Write-Host "Using existing certificate" -ForegroundColor Green
}

# Trust the certificate
Write-Host "Trusting certificate..." -ForegroundColor Yellow
foreach ($storeName in @("Root", "TrustedPublisher")) {
    Add-CertificateToStore -Certificate $cert -StoreName $storeName
}
Add-TrustedRdpPublisherThumbprint -Thumbprint $cert.Thumbprint
Write-Host "Certificate trusted successfully" -ForegroundColor Green
Write-Host ""

# Parse RDP file and build signed version
Write-Host "Processing RDP file..." -ForegroundColor Yellow
$settings = New-Object System.Collections.Generic.List[string]
$fullAddress = $null
$alternateFullAddress = $null
$skipSignatureContinuation = $false

foreach ($line in (Read-RdpLines -Path $InputFile)) {
    # Skip existing signature continuation lines (lines starting with space after signature:s:)
    if ($skipSignatureContinuation) {
        if ($line.StartsWith(" ", [System.StringComparison]::Ordinal)) {
            continue
        }
        $skipSignatureContinuation = $false
    }
    
    # Skip existing signature and signscope lines
    if (Test-LineStartsWith -Line $line -Prefix "signature:s:") {
        $skipSignatureContinuation = $true
        continue
    }
    if (Test-LineStartsWith -Line $line -Prefix "signscope:s:") {
        continue
    }
    
    # Track full address values
    if (Test-LineStartsWith -Line $line -Prefix "full address:s:") {
        $fullAddress = $line.Substring("full address:s:".Length)
    }
    if (Test-LineStartsWith -Line $line -Prefix "alternate full address:s:") {
        $alternateFullAddress = $line.Substring("alternate full address:s:".Length)
    }

    $settings.Add($line)
}

# Add alternate full address if missing (required for signing)
if (-not [string]::IsNullOrEmpty($fullAddress) -and [string]::IsNullOrEmpty($alternateFullAddress)) {
    Write-Host "Adding missing 'alternate full address' field" -ForegroundColor Gray
    $settings.Add("alternate full address:s:$fullAddress")
}

# Extract signable settings in correct order
$signLines = New-Object System.Collections.Generic.List[string]
$signNames = New-Object System.Collections.Generic.List[string]

foreach ($secureSetting in $secureSettings) {
    foreach ($setting in $settings) {
        if (Test-LineStartsWith -Line $setting -Prefix $secureSetting.Prefix) {
            $signNames.Add($secureSetting.Name)
            $signLines.Add($setting)
            break  # Only match each secure setting once
        }
    }
}

if ($signLines.Count -eq 0) {
    Write-Error "No signable RDP settings were found in: $InputFile"
    Exit 1
}

Write-Host "Found $($signLines.Count) signable settings" -ForegroundColor Gray

# Generate signature
Write-Host "Generating signature..." -ForegroundColor Yellow
$signatureValue = New-RdpSignatureValue -SignLines $signLines.ToArray() -SignNames $signNames.ToArray() -Certificate $cert

# Build final output
$outputLines = New-Object System.Collections.Generic.List[string]
$outputLines.AddRange($settings)
$outputLines.Add("signscope:s:" + ($signNames.ToArray() -join ","))
$outputLines.Add((Format-RdpSignatureLine -SignatureValue $signatureValue))

# Write output file with UTF-16 LE + BOM (same as rdpsign.exe)
$outputText = ($outputLines.ToArray() -join "`r`n") + "`r`n"
$utf16WithBom = New-Object System.Text.UnicodeEncoding($false, $true)
[System.IO.File]::WriteAllText($OutputFile, $outputText, $utf16WithBom)

# Verify signature block was written correctly
$fileContent = [System.IO.File]::ReadAllText($OutputFile, [System.Text.Encoding]::Unicode)
$hasSignScope = $fileContent -match '(?m)^signscope:s:.+'
$hasSignature = $fileContent -match '(?m)^signature:s:[A-Za-z0-9+/=]+'

if (-not $hasSignScope -or -not $hasSignature) {
    Write-Error "Signature block missing in signed file."
    Exit 1
}

# Success!
Write-Host ""
Write-Host "SUCCESS" -ForegroundColor Green -BackgroundColor Black
Write-Host "========================================" -ForegroundColor Green
Write-Host "Thumbprint : $($cert.Thumbprint)" -ForegroundColor White
Write-Host "Signed file: $OutputFile" -ForegroundColor White
Write-Host "Valid until: $($cert.NotAfter.ToString('dd/MM/yyyy HH:mm:ss'))" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Green

Exit 0