$ErrorActionPreference = 'Stop'
$appName = 'YangiAlifboOffice'
$root = Join-Path $env:LOCALAPPDATA $appName
$officeApps = @(
    @{ Name = 'Word'; Assembly = 'WordUzbekAlphabetAddIn' },
    @{ Name = 'Excel'; Assembly = 'ExcelUzbekAlphabetAddIn' },
    @{ Name = 'PowerPoint'; Assembly = 'PowerPointUzbekAlphabetAddIn' }
)
foreach ($officeApp in $officeApps) {
    $key = "HKCU:\Software\Microsoft\Office\$($officeApp.Name)\Addins\$($officeApp.Assembly)"
    if (Test-Path $key) { Remove-Item -LiteralPath $key -Recurse -Force }
}
if (Test-Path $root) { Remove-Item -LiteralPath $root -Recurse -Force }
$certFile = Join-Path $PSScriptRoot 'YangiAlifboPublisher.cer'
$thumbprint = if (Test-Path $certFile) {
    ([System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certFile)).Thumbprint
}
if ($thumbprint) {
    $cert = Get-ChildItem Cert:\CurrentUser\TrustedPublisher,Cert:\CurrentUser\Root |
        Where-Object { $_.Thumbprint -eq $thumbprint }
} else { $cert = @() }
foreach ($item in $cert) { Remove-Item -LiteralPath $item.PSPath -Force }
Write-Host 'Надстройки удалены. Перезапустите Office.' -ForegroundColor Green
