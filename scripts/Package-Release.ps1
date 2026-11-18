param([string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswhere)) { throw 'Не найден Visual Studio Installer (vswhere.exe).' }
$msbuild = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if (-not $msbuild) { throw 'Не найден MSBuild. Установите Visual Studio с нагрузкой «Разработка Microsoft 365».' }

# Use a neutral local signing identity. Its private key stays in the Windows certificate store.
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq 'CN=Yangi Alifbo' -and $_.HasPrivateKey } | Select-Object -First 1
if (-not $cert) {
    $cert = New-SelfSignedCertificate -Subject 'CN=Yangi Alifbo' -Type CodeSigningCert -CertStoreLocation 'Cert:\CurrentUser\My' -NotAfter (Get-Date).AddYears(5)
}
$stage = Join-Path $repo 'artifacts\YangiAlifboOffice'
if (Test-Path $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null
$projects = @(
    @{ Name='Word'; Project='WordUzbekAlphabetAddIn'; Assembly='WordUzbekAlphabetAddIn' },
    @{ Name='Excel'; Project='ExcelUzbekAlphabetAddIn'; Assembly='ExcelUzbekAlphabetAddIn' },
    @{ Name='PowerPoint'; Project='PowerPointUzbekAlphabetAddIn'; Assembly='PowerPointUzbekAlphabetAddIn' }
)
foreach ($item in $projects) {
    $csproj = Join-Path $repo "$($item.Project)\$($item.Project).csproj"
    & $msbuild $csproj /t:Rebuild "/p:Configuration=$Configuration" '/p:Platform=AnyCPU' "/p:ManifestCertificateThumbprint=$($cert.Thumbprint)" '/p:SignManifests=true' '/v:minimal'
    if ($LASTEXITCODE -ne 0) { throw "Сборка $($item.Name) завершилась с ошибкой." }
    $source = Join-Path $repo "$($item.Project)\bin\$Configuration"
    $destination = Join-Path $stage $item.Name
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Get-ChildItem -LiteralPath $source -File | Where-Object { $_.Extension -notin @('.pfx','.p12','.pdb','.xml') } | Copy-Item -Destination $destination
}
Export-Certificate -Cert $cert -FilePath (Join-Path $stage 'YangiAlifboPublisher.cer') -Force | Out-Null
Copy-Item (Join-Path $PSScriptRoot 'Install.ps1'),(Join-Path $PSScriptRoot 'Uninstall.ps1'),(Join-Path $PSScriptRoot 'Install.cmd'),(Join-Path $PSScriptRoot 'Uninstall.cmd') -Destination $stage
Copy-Item (Join-Path $repo 'README.md') -Destination $stage
$zip = Join-Path $repo 'artifacts\YangiAlifboOffice-Setup.zip'
if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -CompressionLevel Optimal
Write-Host "Архив готов: $zip" -ForegroundColor Green
