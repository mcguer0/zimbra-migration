[CmdletBinding()]
param(
  [Parameter(Mandatory=$true, HelpMessage="user или user@example.com")]
  [string]$User
)

# == Подключаем конфиг ==
. "$PSScriptRoot/../scripts/config.ps1"

# == Локальные функции ==
function Ensure-ExchangeCmdlets {
  if (Get-Command Get-MailContact -ErrorAction SilentlyContinue) { return }
  Write-Host "Подгружаю Exchange cmdlets..."
  try {
    $snap = Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "Exchange" }
    if ($snap) {
      foreach ($s in $snap) { Add-PSSnapin $s.Name -ErrorAction Stop }
      if (Get-Command Get-MailContact -ErrorAction SilentlyContinue) { return }
    }
  } catch {}
  try {
    $sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeMgmtHost) -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $sess -DisableNameChecking | Out-Null
    return
  } catch {
    throw "Не удалось загрузить Exchange cmdlets (ни snapin, ни remoting). Проверь запуск на сервере Exchange."
  }
}

# == Инициализация окружения ==
Ensure-ExchangeCmdlets

if ($User -notlike "*@*") { $User = "$User@$Domain" }

$contact = Get-MailContact -Identity $User -ErrorAction SilentlyContinue
if (-not $contact) {
  Write-Host "Контакт $User не найден." -ForegroundColor Yellow
  return
}

Write-Host "Контакт найден: $($contact.PrimarySmtpAddress)" -ForegroundColor Green
$groups = Get-DistributionGroup -ResultSize Unlimited -Filter "Members -eq '$($contact.DistinguishedName)'"
if ($groups) {
  Write-Host "Состоит в группах:" -ForegroundColor Cyan
  $groups | ForEach-Object { Write-Host $_.PrimarySmtpAddress }
} else {
  Write-Host "Не состоит ни в одной группе рассылки."
}
