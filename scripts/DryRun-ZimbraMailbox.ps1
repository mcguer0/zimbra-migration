# Требует: config.ps1
# Экспортирует: Invoke-DryRunZimbraMailbox

function Invoke-DryRunZimbraMailbox([string]$UserInput) {
  # Нормализуем user/email
  if ($UserInput -like "*@*") { $UserEmail = $UserInput; $Alias = ($UserInput -split "@")[0] }
  else                        { $Alias = $UserInput; $UserEmail = "$UserInput@$Domain" }

  Write-Host "=== Dry run for $UserEmail ==="

  # Проверка почтового контакта
  $mailContact = $null
  try { $mailContact = Get-MailContact -Identity $UserEmail -ErrorAction Stop } catch {}
  if (-not $mailContact) { try { $mailContact = Get-Contact -Identity $UserEmail -ErrorAction Stop } catch {} }
  if ($mailContact) {
    Write-Host "MailContact найден: $($mailContact.Identity)"
  } else {
    Write-Host "MailContact не найден."
  }

  # Проверка mailbox
  $mailbox = $null
  try { $mailbox = Get-Mailbox -Identity $UserEmail -ErrorAction Stop } catch {}
  if ($mailbox) {
    Write-Host "Mailbox найден."
  } else {
    Write-Host "Mailbox не найден."
  }

  # Сбор членств в рассылках
  $groups = @()
  try {
    $recipient = Get-Recipient -Filter "EmailAddresses -eq '$UserEmail' -or PrimarySmtpAddress -eq '$UserEmail'" -ErrorAction Stop
    if ($recipient) {
      $groups = Get-DistributionGroupsByMember $UserEmail |
        Select-Object DisplayName,PrimarySmtpAddress |
        Sort-Object DisplayName
    }
  } catch {}

  if ($groups -and $groups.Count -gt 0) {
    $info = $groups | ForEach-Object { "{0}/{1}" -f $_.DisplayName, $_.PrimarySmtpAddress }
    Write-Host ("Состоит в {0} рассылк(ах): {1}" -f $groups.Count, ($info -join ", "))
  } else {
    Write-Host "Не состоит ни в одной рассылке (AD-группе)."
  }

  # Тестовое подключение к серверам
  Write-Host "`n=== Проверка подключений ==="
  $tests = @(
    @{Name="Exchange IMAP";      Host=$ExchangeImapHost; Port=$ExchangeImapPort},
    @{Name="Zimbra IMAP";        Host=$ZimbraImapHost;   Port=$ZimbraImapPort},
    @{Name="Exchange PowerShell";Host=$ExchangeMgmtHost; Port=80},
    @{Name="Zimbra SSH";         Host=$ZimbraSshHost;    Port=22}
  )
  foreach ($t in $tests) {
    Write-Host ("{0} ({1}:{2})..." -f $t.Name, $t.Host, $t.Port)
    try {
      $ok = Test-NetConnection -ComputerName $t.Host -Port $t.Port -InformationLevel Quiet -WarningAction SilentlyContinue
      if ($ok) { Write-Host "  OK" } else { Write-Warning "  Нет соединения" }
    } catch {
      Write-Warning ("  Ошибка: {0}" -f $_.Exception.Message)
    }
  }

  return @{
    UserEmail = $UserEmail
    Alias     = $Alias
  }
}

