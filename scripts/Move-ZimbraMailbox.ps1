# Требует: config.ps1, utils.ps1
# Экспортирует: Invoke-MoveZimbraMailbox

function Invoke-MoveZimbraMailbox([string]$UserInput) {
  # Нормализация user/email
  if ($UserInput -like "*@*") { $UserEmail = $UserInput; $Alias = ($UserInput -split "@")[0] }
  else                        { $Alias = $UserInput; $UserEmail = "$UserInput@$Domain" }

  Write-Host "=== Подготовка ящика в Exchange для $UserEmail ==="

  # Проверка почтового контакта и сбор его членств в рассылках (если есть), затем отключение контакта. PS 5.1 совместимо
  $contactGroups = @()   # сюда соберём рассылки (AD-группы), где состоит контакт
  try {
    $mailContact = $null
    try { $mailContact = Get-Contact -Identity $UserEmail -ErrorAction Stop } catch {}
    if (-not $mailContact) { try { $mailContact = Get-MailContact -Identity $UserEmail -ErrorAction Stop } catch {} }

    if ($mailContact) {
      Write-Host "Найден MailContact: $($mailContact.Identity) — проверяю членства в рассылках..."
      try {
        # Ищем AD-объект контакта и собираем группы, где он состоит как member
        $adContact = $null
        try {
          $adContact = Get-ADObject -LDAPFilter "(&(objectClass=contact)(|(mail=$UserEmail)(proxyAddresses=SMTP:$UserEmail)(proxyAddresses=smtp:$UserEmail)))" -Properties DistinguishedName,memberOf -ErrorAction Stop
        } catch {}
        if ($adContact) {
          $groupDns = @($adContact.memberOf)
          if ($groupDns.Count -gt 0) {
            $contactGroups = $groupDns | ForEach-Object {
              try { Get-ADGroup -Identity $_ -Properties DistinguishedName,Name,mail -ErrorAction Stop } catch {}
            } | Where-Object { $_ } | Sort-Object Name
            if ($contactGroups.Count -gt 0) {
              Write-Host ("Контакт состоит в {0} групп(ах): {1}" -f $contactGroups.Count, (($contactGroups | Select-Object -Expand Name) -join ", "))
            } else {
              Write-Host "Контакт не состоит ни в одной рассылке (AD-группе)."
            }
          } else {
            Write-Host "Контакт не состоит ни в одной рассылке (AD-группе)."
          }
        } else {
          Write-Host "AD-объект контакта не найден по mail/proxyAddresses — пропускаю сбор групп."
        }
      } catch {
        Write-Warning "Не удалось определить членства контакта в группах: $($_.Exception.Message)"
      }

      # Теперь удаляем сам почтовый контакт
      try {
        Remove-MailContact -Identity $mailContact.Identity -Confirm:$false -ErrorAction Stop
        Write-Host "MailContact удалён: $($mailContact.Identity)"

        if ($adContact) {
          try {
            Remove-ADObject -Identity $adContact.DistinguishedName -Confirm:$false -ErrorAction Stop
            Write-Host "AD-объект контакта удалён: $($adContact.DistinguishedName)"
          } catch {
            Write-Warning "Не удалось удалить AD-объект контакта '$($adContact.DistinguishedName)': $($_.Exception.Message)"
          }
        }
      } catch {
        Write-Warning "Не удалось удалить MailContact '$($mailContact.Identity)': $($_.Exception.Message)"
      }
    } else {
      Write-Host "MailContact не найден для $UserEmail (ok)"
    }
  } catch {
    Write-Warning "Проверка/удаление MailContact завершилась ошибкой: $($_.Exception.Message)"
  }

  # Mailbox: существует?
  try {
    $null = Get-Mailbox -Identity $UserEmail -ErrorAction Stop
    Write-Host "Mailbox уже существует."
  } catch {
    Write-Host "Mailbox не найден. Enable-Mailbox для '$Alias'..."
    Enable-Mailbox -Identity $Alias -PrimarySmtpAddress $UserEmail -Alias $Alias -ErrorAction Stop | Out-Null
    Write-Host "Mailbox включён. Пауза 60 сек для репликации..."
    Start-Sleep -Seconds 60
  }

  # UPN (если задан UpnSuffix — используем его; иначе домен)
  if (-not $UpnSuffix -or [string]::IsNullOrWhiteSpace($UpnSuffix)) { $UpnSuffix = $Domain }
  try {
    $ldap = "(|(sAMAccountName=$Alias)(userPrincipalName=$UserEmail))"
    $adUser = Get-ADUser -LDAPFilter $ldap -Properties userPrincipalName,samAccountName -ErrorAction Stop
    $desiredUpn = "$($adUser.SamAccountName)@$UpnSuffix"
    if ($adUser.UserPrincipalName -ne $desiredUpn) {
      Set-ADUser -Identity $adUser -UserPrincipalName $desiredUpn -ErrorAction Stop
      Write-Host "UPN обновлён: $desiredUpn"
    } else {
      Write-Host "UPN уже корректный: $desiredUpn"
    }

    # Если ранее был найден и отключён MailContact, и мы собрали список его рассылок —
    # добавляем текущего пользователя (AD user) в эти группы
    if ($contactGroups -and $contactGroups.Count -gt 0) {
      Write-Host "Добавляю пользователя $($adUser.SamAccountName) в рассылки вместо контакта..."
      foreach ($g in $contactGroups) {
        try {
          Add-ADGroupMember -Identity $g.DistinguishedName -Members $adUser -ErrorAction Stop
          Write-Host "  + $($g.Name) — добавлен"
        } catch {
          # Если уже член — подавим ошибку с понятным сообщением
          Write-Warning ("  ! {0} — {1}" -f $g.Name, $_.Exception.Message)
        }
      }
    }
  } catch {
    Write-Warning "Не удалось привести UPN к $UpnSuffix для '$Alias': $($_.Exception.Message)"
  }

  # Включаем IMAP
  $cas = Get-CASMailbox -Identity $UserEmail -ErrorAction Stop
  if (-not $cas.ImapEnabled) {
    Set-CASMailbox -Identity $UserEmail -ImapEnabled $true -ErrorAction Stop
    Write-Host "IMAP включён."
  } else {
    Write-Host "IMAP уже включён."
  }

  # FullAccess для админа
  $perm = Get-MailboxPermission -Identity $UserEmail | Where-Object { $_.User -eq $AdminLogin -and $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }
  if (-not $perm) {
    Add-MailboxPermission -Identity $UserEmail -User $AdminLogin -AccessRights FullAccess -InheritanceType All -AutoMapping:$false | Out-Null
    Write-Host "FullAccess выдан $AdminLogin."
  } else {
    Write-Host "FullAccess уже есть."
  }
  Start-Sleep -Seconds 2

  # Готовим удалённый bash-скрипт imapsync
  $AdminImapB64   = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($AdminImapPasswordPlain))
  $ssl1 = if ($ZimbraImapSSL)   { "--ssl1" } else { "" }
  $ssl2 = if ($ExchangeImapSSL) { "--ssl2" } else { "" }
  $ts = Get-Date -Format "yyyyMMdd-HHmmss"
  $RemoteUserTag  = ($UserEmail -replace "@","_" -replace "/","_")
  $RemoteLog      = "/tmp/imapsync-$RemoteUserTag-$ts.log"
  $RemoteScript   = "/tmp/run-imapsync-$RemoteUserTag-$ts.sh"

  $bash = @'
#!/usr/bin/env bash
set -euo pipefail
IMAPSYNC="__IMAPSYNC_PATH__"
if ! command -v "$IMAPSYNC" >/dev/null 2>&1; then echo "imapsync not found at $IMAPSYNC" >&2; exit 127; fi
TS="$(date +%Y%m%d-%H%M%S)"
LOGFILE="__REMOTE_LOG__"
ADMIN_B64='__ADMIN_IMAP_B64__'
ADMIN_PLAIN="$(printf %s "$ADMIN_B64" | base64 -d)"
mkdir -p "$(dirname "$LOGFILE")" || true
: > "$LOGFILE"
exec > >(tee -a "$LOGFILE") 2>&1
echo "[imapsync] start for __USER_EMAIL__ at $TS, log: $LOGFILE"

TRIES=5
DELAY=15
attempt=1
rc=111

while [ $attempt -le $TRIES ]; do
  echo "[imapsync] attempt $attempt/$TRIES"
  "$IMAPSYNC" \
    --host1 "__ZIMBRA_IMAP_HOST__" --port1 "__ZIMBRA_IMAP_PORT__" __SSL1__ \
    --user1 "__USER_EMAIL__" \
    --authuser1 "__ADMIN_LOGIN__" --password1 "$ADMIN_PLAIN" \
    --host2 "__EXCHANGE_IMAP_HOST__" --port2 "__EXCHANGE_IMAP_PORT__" __SSL2__ \
    --user2 "__USER_EMAIL__" \
    --authuser2 "__ADMIN_LOGIN__" --password2 "$ADMIN_PLAIN" \
    --useuid \
    --syncinternaldates \
    --idatefromheader \
    --automap \
    --regextrans2 's#^Sent$#Sent Items#' \
    --regextrans2 's#^Trash$#Deleted Items#' \
    --regextrans2 's#^Junk$#Junk Email#' \
    --exclude '(Trash|Junk|Spam)$' \
    --skipemptyfolders \
    --addheader \
    --fastio1 --fastio2 \
    --nofoldersizes \
    --usecache \
    --useheader 'Message-Id' \
    --delete2duplicates \
    --timeout 600 \
    --logfile "$LOGFILE"

  rc=$?
  echo "[imapsync] exit code: $rc"
  if [ $rc -eq 0 ]; then break; fi
  if [ $attempt -lt $TRIES ]; then
    echo "[imapsync] will retry after ${DELAY}s ..."
    sleep $DELAY
    DELAY=$((DELAY*2))
  fi
  attempt=$((attempt+1))
done
exit $rc
'@

  $repl = @{
    "__IMAPSYNC_PATH__"       = $ImapSyncPath
    "__REMOTE_LOG__"          = $RemoteLog
    "__ADMIN_IMAP_B64__"      = $AdminImapB64
    "__USER_EMAIL__"          = $UserEmail
    "__ZIMBRA_IMAP_HOST__"    = $ZimbraImapHost
    "__ZIMBRA_IMAP_PORT__"    = "$ZimbraImapPort"
    "__EXCHANGE_IMAP_HOST__"  = $ExchangeImapHost
    "__EXCHANGE_IMAP_PORT__"  = "$ExchangeImapPort"
    "__SSL1__"                = $ssl1
    "__SSL2__"                = $ssl2
    "__ADMIN_LOGIN__"         = $AdminLogin
  }
  foreach ($k in $repl.Keys) { $bash = $bash.Replace($k, $repl[$k]) }
  $bashLF = $bash -replace "`r?`n", "`n"

  # SSH к Zimbra и потоковый запуск
  Ensure-Module Posh-SSH
  $zSess = New-SSHSess -SshHost $ZimbraSshHost -SshUser $ZimbraSshUser -SshPass $ZimbraSshPasswordPlain
  if (-not $zSess) { throw "SSH к $ZimbraSshHost не установлен" }

  $LocalLog = Join-Path $LocalLogDir ("imapsync-{0}-{1}.log" -f ($UserEmail -replace "@","_"), $ts)
  try {
    $scriptB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($bashLF))
    $uploadCmd = "bash -lc 'printf %s ""$scriptB64"" | base64 -d > ""$RemoteScript"" && chmod 600 ""$RemoteScript""'"
    Invoke-SSHCommand -SessionId $zSess.SessionId -Command $uploadCmd | Out-Null

    Write-Host "Запускаю imapsync и стримлю вывод..."
    $stream = New-SSHShellStream -SessionId $zSess.SessionId -TerminalName 'xterm' -BufferSize 8192
    $startCmd = "bash -lc 'bash ""$RemoteScript""; printf ""__END__:%s\n"" `$?; exit'"
    $stream.WriteLine($startCmd)

    $writer = New-Object System.IO.StreamWriter($LocalLog, $false, [Text.Encoding]::UTF8)
    try {
      $exitCode = $null
      while ($true) {
        if ($stream.DataAvailable) {
          $chunk = $stream.Read()
          if ($chunk) {
            Write-Host -NoNewline $chunk
            $writer.Write($chunk); $writer.Flush()
            if ($chunk -match "__END__:(\d+)") { $exitCode = [int]$matches[1]; break }
            elseif ($chunk -match "__END__:(True|False)") {
            if ($matches[1] -eq "True") {
              $exitCode = 0
            } else {
              $exitCode = 1
            }
            break
          }
          }
        } elseif ($stream.IsClosed) {
          break
        } else {
          Start-Sleep -Milliseconds 200
        }
      }
    } finally {
      $writer.Close()
    }
    if ($null -eq $exitCode) { $exitCode = 1 }
  }
  finally {
    Invoke-SSHCommand -SessionId $zSess.SessionId -Command ("bash -lc 'rm -f ""{0}"" ""{1}""'" -f $RemoteScript, $RemoteLog) | Out-Null
    Remove-SSHSession -SessionId $zSess.SessionId | Out-Null
  }

  return @{
    ExitCode = $exitCode
    LocalLog = $LocalLog
    UserEmail = $UserEmail
    Alias = $Alias
  }
}
