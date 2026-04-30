#Requires -RunAsAdministrator
Param(
    [string] $InputFile  = ".\connection.rdp",
    [string] $OutputFile = "",
    [string] $CAName     = "WolClientRDPPublisher",
    [int]    $Years      = 5
)

# Based on: https://github.com/IanVanLier/April-2026-security-update-Remote-Desktop-Conection-security-warning-

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Verify input
if (-not (Test-Path $InputFile)) {
    Write-Error "RDP file not found: $InputFile"
    Exit 1
}
$InputFile = (Resolve-Path $InputFile).Path

# Default OutputFile: same dir, base name + -signed.rdp
if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $dir  = [System.IO.Path]::GetDirectoryName($InputFile)
    if ([string]::IsNullOrEmpty($dir)) { $dir = "." }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $OutputFile = Join-Path $dir ("$base-signed.rdp")
}

# Certificate: find or create
$certSubject = "CN=$CAName"
$cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -eq $certSubject } | Select-Object -First 1

if (-not $cert) {
    $cert = New-SelfSignedCertificate -Subject $certSubject -CertStoreLocation "Cert:\LocalMachine\My" -Type CodeSigningCert -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears($Years)

    foreach ($storeName in @("Root", "TrustedPublisher")) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, "LocalMachine")
        $store.Open("ReadWrite")
        $store.Add($cert)
        $store.Close()
    }

    $gpoPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
    $keyName = "TrustedCertThumbprints"
    if (-not (Test-Path $gpoPath)) { New-Item -Path $gpoPath -Force | Out-Null }

    $existingProp = Get-ItemProperty -Path $gpoPath -Name $keyName -ErrorAction SilentlyContinue
    if ($existingProp) {
        $existingType   = (Get-Item $gpoPath).GetValueKind($keyName)
        $existingValues = $existingProp.$keyName
    } else {
        $existingType   = [Microsoft.Win32.RegistryValueKind]::String
        $existingValues = ""
    }

    if ($existingValues -notlike "*$($cert.Thumbprint)*") {
        $newValue = if ([string]::IsNullOrWhiteSpace($existingValues)) { $cert.Thumbprint } else { "$existingValues,$($cert.Thumbprint)" }
        Set-ItemProperty -Path $gpoPath -Name $keyName -Value $newValue -Type $existingType
    }
}

# Create a copy and sign it
Copy-Item -Path $InputFile -Destination $OutputFile -Force
rdpsign.exe /sha256 $cert.Thumbprint "$OutputFile"
if ($LASTEXITCODE -ne 0) {
    Write-Error "rdpsign.exe failed (exit $LASTEXITCODE)."
    Exit 1
}

# Verify signature
$fileContent  = Get-Content $OutputFile -Raw
$hasSignScope = $fileContent -match '(?m)^signscope:s:.+'
$hasSignature = $fileContent -match '(?m)^signature:s:[A-Za-z0-9+/=]+'
if (-not $hasSignScope -or -not $hasSignature) {
    Write-Error "Signature block missing in signed file."
    Exit 1
}

# Output
Write-Host "SUCCESS" -ForegroundColor Green
Write-Host "Thumbprint : $($cert.Thumbprint)"
Write-Host "Signed file: $OutputFile"
Write-Host "Valid until: $($cert.NotAfter.ToString('dd/MM/yyyy'))"

Exit 0