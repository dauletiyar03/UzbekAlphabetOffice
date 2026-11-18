$ErrorActionPreference = 'Stop'
$appName = 'YangiAlifboOffice'
$root = Join-Path $env:LOCALAPPDATA $appName
$certPath = Join-Path $PSScriptRoot 'YangiAlifboPublisher.cer'

if (-not (Test-Path $certPath)) { throw 'Не найден файл сертификата YangiAlifboPublisher.cer.' }

$hosts = @(
    @{ Name = 'Word'; Assembly = 'WordUzbekAlphabetAddIn' },
    @{ Name = 'Excel'; Assembly = 'ExcelUzbekAlphabetAddIn' },
    @{ Name = 'PowerPoint'; Assembly = 'PowerPointUzbekAlphabetAddIn' }
)
$installed = @()
foreach ($host in $hosts) {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\$($host.Name)\InstallRoot",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\$($host.Name)\InstallRoot"
    )
    if ($regPaths | Where-Object { Test-Path $_ }) { $installed += $host }
}
if ($installed.Count -eq 0) { throw 'Не найден установленный настольный Office 2016 или новее (Word, Excel, PowerPoint).' }

Write-Host 'Сертификат подписи нужен, чтобы Windows доверяла локальной надстройке.' -ForegroundColor Yellow
Write-Host 'Издатель: Yangi Alifbo (самоподписанный сертификат из этого архива).' -ForegroundColor Yellow
$answer = Read-Host 'Доверять этому сертификату и установить надстройку? (Y/N)'
if ($answer -notin @('Y','y','Д','д')) { Write-Host 'Установка отменена.'; exit 1 }

foreach ($storeName in @('Root','TrustedPublisher')) {
    Import-Certificate -FilePath $certPath -CertStoreLocation "Cert:\CurrentUser\$storeName" | Out-Null
}

if (Test-Path $root) { Remove-Item -LiteralPath $root -Recurse -Force }
New-Item -ItemType Directory -Path $root -Force | Out-Null
foreach ($host in $installed) {
    $source = Join-Path $PSScriptRoot $host.Name
    if (-not (Test-Path (Join-Path $source ($host.Assembly + '.vsto')))) {
        throw "В архиве отсутствуют файлы надстройки $($host.Name)."
    }
    $target = Join-Path $root $host.Name
    Copy-Item -Path $source -Destination $target -Recurse
    $manifest = 'file:///' + ((Join-Path $target ($host.Assembly + '.vsto')).Replace('\','/')) + '|vstolocal'
    $key = "HKCU:\Software\Microsoft\Office\$($host.Name)\Addins\$($host.Assembly)"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name FriendlyName -Value 'O‘zbek alifbosi' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name Description -Value 'Преобразование текста в узбекский латинский алфавит' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $key -Name Manifest -Value $manifest -PropertyType String -Force | Out-Null
    Write-Host "Установлено для $($host.Name)"
}
Write-Host 'Готово. Закройте и снова откройте приложения Office.' -ForegroundColor Green
