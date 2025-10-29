# Требует: config.ps1, utils.ps1
# Экспортирует: Invoke-MoveZimbraMailbox

function Invoke-MoveZimbraMailbox([string]$UserInput, [switch]$Staged, [switch]$Activate) {
  # Нормализация user/email
  if ($UserInput -like "*@*") { $UserEmail = $UserInput; $Alias = ($UserInput -split "@")[0] }
  else                        { $Alias = $UserInput; $UserEmail = "$UserInput@$Domain" }

  Write-Host "=== Подготовка ящика в Exchange для $UserEmail ==="

  # Проверка контакта и групп
  $contactGroups = @()
  try {
    $contact = Get-MailContact -Identity $UserEmail -ErrorAction SilentlyContinue
    if ($contact) {
      Write-Host "Найден контакт $UserEmail. Сохраняю группы..."
      $contactGroups = Get-DistributionGroup -ResultSize Unlimited -Filter "Members -eq '$($contact.DistinguishedName)'" -ErrorAction SilentlyContinue
      # Группы, где контакт разрешён как отправитель (Delivery Management)
      $contactAllowedGroups = @()
      try {
        $allDg = Get-DistributionGroup -ResultSize Unlimited
        foreach ($dg in $allDg) {
          $allowedDns = @()
          if ($dg.AcceptMessagesOnlyFromSendersOrMembers) {
            foreach ($id in $dg.AcceptMessagesOnlyFromSendersOrMembers) {
              try { $r = Get-Recipient -Identity $id -ErrorAction Stop } catch { $r = $null }
              if ($r -and $r.DistinguishedName) { $allowedDns += $r.DistinguishedName }
            }
          }
          if ($allowedDns -and ($allowedDns -contains $contact.DistinguishedName)) { $contactAllowedGroups += $dg }
        }
        if ($contactAllowedGroups.Count -gt 0) {
          Write-Host ("Контакт разрешён отправителем в группах: {0}" -f ($contactAllowedGroups.PrimarySmtpAddress -join ', '))
        }
      } catch {
        Write-Warning ("Не удалось определить Delivery Management группы: {0}" -f $_.Exception.Message)
      }
      if ($contactGroups) {
        Write-Host ("Контакт состоит в группах: {0}" -f ($contactGroups.PrimarySmtpAddress -join ', '))
      } else {
        Write-Host "Контакт не состоит ни в одной группе."
      }

      if ($Staged) {
        Write-Host "Добавляю пользователя в группы контакта (по объекту mailbox)..."
        # Разрешим добавление по объекту самого временного ящика, чтобы не попасть в контакт
        $rcp = $null
        try { $rcp = Get-Recipient -Identity "${Alias}_1" -ErrorAction Stop } catch {
          try { $rcp = Get-Recipient -Identity $Alias -ErrorAction Stop } catch { $rcp = $null }
        }
        foreach ($g in $contactGroups) {
          try {
            Set-DistributionGroup -Identity $g.Identity -RequireSenderAuthenticationEnabled $false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            if ($rcp) {
              $added = $false
              try {
                Add-DistributionGroupMember -Identity $g.Identity -Member $rcp.Identity -ErrorAction Stop
                $added = $true
              } catch {
                if ($_.Exception.Message -match 'already a member') {
                  Write-Host "Уже состоит в группе $($g.PrimarySmtpAddress)"; $added = $true
                } elseif ($_.Exception.Message -match 'неправильный тип учетной записи|Invalid argument|local group') {
                  # Fallback: AD-уровень
                  try { Add-ADGroupMember -Identity $g.DistinguishedName -Members $rcp.DistinguishedName -ErrorAction Stop; $added = $true }
                  catch {
                    $m = $_.Exception.Message
                    if ($m -match 'already.*member|уже.*(состоит|присутствует)') { $added = $true }
                    else { Write-Warning ("[AD] Не удалось добавить в {0}: {1}" -f $g.PrimarySmtpAddress, $m) }
                  }
                } else { Write-Warning ("Не удалось добавить в {0}: {1}" -f $g.PrimarySmtpAddress, $_.Exception.Message) }
              }
              if ($added) { Write-Host "Добавлен в группу $($g.PrimarySmtpAddress)" }
            } else {
              # если не удалось резолвить получателя — пробуем по адресу как раньше
              try { Add-DistributionGroupMember -Identity $g.Identity -Member $UserEmail -ErrorAction Stop; Write-Host "Добавлен (по адресу) в группу $($g.PrimarySmtpAddress)" }
              catch { Write-Warning ("Не удалось добавить (по адресу) в {0}: {1}" -f $g.PrimarySmtpAddress, $_.Exception.Message) }
            }
          } catch {
            Write-Warning ("Не удалось добавить в группу {0}: {1}" -f $g.PrimarySmtpAddress, $_.Exception.Message)
          }
        }
        $AliasTemp = "${Alias}_1"
        $TempEmail = "$AliasTemp@$Domain"
        Write-Host "Контакт остаётся до финального запуска."
      } elseif ($Activate) {
        $AliasTemp = "${Alias}_1"
        $TempEmail = "$AliasTemp@$Domain"
        $swapOk = $false
        try {
          Write-Host "Подменяю отправителя в Delivery Management (контакт -> mailbox)..."
          Replace-AcceptedSender -OldContactSmtp $UserEmail -NewMailboxId $TempEmail -ErrorAction Stop | Out-Null
          $swapOk = $true
        } catch {
          Write-Warning ("Ошибка подмены отправителя для групп: {0}" -f $_.Exception.Message)
        }
        if (-not $swapOk -and $contactAllowedGroups) {
          foreach ($dg in $contactAllowedGroups) {
            try {
              Set-DistributionGroup -Identity $dg.Identity -AcceptMessagesOnlyFromSendersOrMembers @{ Add = $TempEmail; Remove = $contact.Identity } -ErrorAction Stop
              Write-Host ("[Fallback] Delivery Management обновлён: {0}" -f $dg.Name)
            } catch {
              Write-Warning ("[Fallback] Не удалось обновить Delivery Management для {0}: {1}" -f $dg.Name, $_.Exception.Message)
            }
          }
        }

        Write-Host "Обеспечиваю членство пользователя в группах (перед удалением контакта)..."
        # Добавляем сам почтовый ящик по объекту, игнорируя "already a member"
        $userRcp = $null
        try { $userRcp = Get-Recipient -Identity $Alias -ErrorAction Stop } catch {
          try { $userRcp = Get-Recipient -Identity $TempEmail -ErrorAction Stop } catch { $userRcp = $null }
        }
        foreach ($g in $contactGroups) {
          try {
            # Снимем требование аутентификации отправителя, чтобы не мешало добавлению
            try { Set-DistributionGroup -Identity $g.Identity -RequireSenderAuthenticationEnabled $false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null } catch {}
            if ($userRcp) {
              $added = $false
              try {
                Add-DistributionGroupMember -Identity $g.Identity -Member $userRcp.Identity -ErrorAction Stop
                $added = $true
              } catch {
                if ($_.Exception.Message -match 'already a member') {
                  Write-Host "Уже состоит в группе $($g.PrimarySmtpAddress)"; $added = $true
                } elseif ($_.Exception.Message -match 'неправильный тип учетной записи|Invalid argument|local group') {
                  # Fallback через AD
                  try { Add-ADGroupMember -Identity $g.DistinguishedName -Members $userRcp.DistinguishedName -ErrorAction Stop; $added = $true }
                  catch {
                    $m = $_.Exception.Message
                    if ($m -match 'already.*member|уже.*(состоит|присутствует)') { $added = $true }
                    else {
                      # Последняя попытка — добавить по SMTP/алиасу через Exchange cmdlet
                      try { Add-DistributionGroupMember -Identity $g.Identity -Member $UserEmail -ErrorAction Stop; $added = $true }
                      catch { Write-Warning ("[AD] Не удалось добавить в {0}: {1}" -f $g.PrimarySmtpAddress, $m) }
                    }
                  }
                } else {
                  Write-Warning ("Не удалось добавить в {0}: {1}" -f $g.PrimarySmtpAddress, $_.Exception.Message)
                }
              }
              if ($added) { Write-Host "Добавлен в группу $($g.PrimarySmtpAddress)" }
            } else {
              # fallback по адресу
              try { Add-DistributionGroupMember -Identity $g.Identity -Member $UserEmail -ErrorAction Stop; Write-Host "Добавлен (по адресу) в группу $($g.PrimarySmtpAddress)" }
              catch { Write-Warning ("Не удалось добавить (по адресу) в {0}: {1}" -f $g.PrimarySmtpAddress, $_.Exception.Message) }
            }
          } catch {
            Write-Warning ("Не удалось обеспечить членство в группе {0}: {1}" -f $g.PrimarySmtpAddress, $_.Exception.Message)
          }
        }

        Write-Host "Удаляю контакт $UserEmail..."
        Remove-MailContact -Identity $contact.Identity -Confirm:$false -ErrorAction Stop
        Write-Host "Контакт удалён."
        try {
          Write-Host "Переименовываю временный ящик $TempEmail в $UserEmail..."
          $mbx = $null
          try { $mbx = Get-Mailbox -Identity $TempEmail -ErrorAction Stop } catch {}
          $mbxId = if ($mbx) { $mbx.Identity } else { $TempEmail }

          # Сначала приводим Alias и набор адресов (add/remove), затем делаем primary
          try { Set-Mailbox -Identity $mbxId -Alias $Alias -ErrorAction Stop } catch {}
          try { Set-Mailbox -Identity $mbxId -EmailAddresses @{Add=$UserEmail; Remove=$TempEmail} -ErrorAction SilentlyContinue } catch {}

          $success = $false
          $delay = 10
          for ($i=1; $i -le 6 -and -not $success; $i++) {
            try {
              Set-Mailbox -Identity $mbxId -PrimarySmtpAddress $UserEmail -ErrorAction Stop
              $success = $true
            } catch {
              $msg = $_.Exception.Message
              if ($msg -match 'already.*used|уже используется') {
                $holder = $null
                try { $holder = Get-Recipient -Identity $UserEmail -ErrorAction Stop } catch {}
                if ($holder -and $mbx -and ($holder.DistinguishedName -ne $mbx.DistinguishedName)) {
                  Write-Warning ("Адрес {0} занят объектом: {1}. Повтор через {2} c..." -f $UserEmail,$holder.Identity,$delay)
                } else {
                  Write-Warning ("Адрес {0} ещё не освободился (репликация?). Повтор через {1} c..." -f $UserEmail,$delay)
                }
                Start-Sleep -Seconds $delay
                $delay = [Math]::Min($delay*2, 60)
              } else {
                throw
              }
            }
          }

          if ($success) {
            # после успешного переименования все дальнейшие операции должны идти по новому адресу
            $mailboxIdentity = $UserEmail
            # удалить старый адрес _1 как вторичный, если остался
            try { Set-Mailbox -Identity $UserEmail -EmailAddresses @{ Remove = $TempEmail } -ErrorAction SilentlyContinue } catch {}
            Write-Host "Пауза 30 сек для репликации адресов/политик..."
            Start-Sleep -Seconds 30
            # Финальная проверка членства после переименования
            if ($contactGroups -and $contactGroups.Count -gt 0) {
              foreach ($g in $contactGroups) {
                try {
                  try { Set-DistributionGroup -Identity $g.Identity -RequireSenderAuthenticationEnabled $false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null } catch {}
                  Add-DistributionGroupMember -Identity $g.Identity -Member $UserEmail -ErrorAction Stop
                  Write-Host ("[Final] Добавлен в группу {0}" -f $g.PrimarySmtpAddress)
                } catch {
                  if ($_.Exception.Message -notmatch 'already a member') {
                    Write-Warning ("[Final] Не удалось добавить в {0}: {1}" -f $g.PrimarySmtpAddress, $_.Exception.Message)
                  }
                }
              }
            }
          } else {
            Write-Warning ("Не удалось назначить PrimarySmtpAddress = {0}. Оставил адрес в списке proxy (smtp:). Продолжаю работать с временным адресом {1}." -f $UserEmail,$TempEmail)
            $mailboxIdentity = $TempEmail
          }
        } catch {
          Write-Warning ("Не удалось переименовать временный ящик {0}: {1}" -f $TempEmail, $_.Exception.Message)
        }
      } else {
        Write-Host "Контакт остаётся (не указан -Activate)."
      }
    }
  } catch {
    Write-Warning ("Не удалось обработать контакт {0}: {1}" -f $UserEmail, $_.Exception.Message)
  }

  # Раннюю попытку подготовки Delivery Management убрали — выполняется строго после Enable-Mailbox

    # Mailbox: существует?
  try {
    $mailboxIdentity = if ($Staged -and $TempEmail) { $TempEmail } else { $UserEmail }
    $null = Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
    Write-Host "Mailbox уже существует."
  } catch {
    if ($Activate) {
      Write-Warning "Mailbox $mailboxIdentity не найден."
    } elseif ($Staged -and $contact) {
      Write-Host "Mailbox не найден. Enable-Mailbox для '$Alias' с временным алиасом '$AliasTemp'..."
      Enable-Mailbox -Identity $Alias -PrimarySmtpAddress $TempEmail -Alias $AliasTemp -ErrorAction Stop | Out-Null
      Write-Host "Mailbox включён. Пауза 30 сек для репликации..."
      Start-Sleep -Seconds 30
      Set-Mailbox -Identity $TempEmail -HiddenFromAddressListsEnabled $true -ErrorAction Stop
      try {
        Set-ADUser -Identity $Alias -EmailAddress $UserEmail -ErrorAction Stop
      } catch {
        Write-Warning ("Не удалось обновить поле mail для {0}: {1}" -f $Alias, $_.Exception.Message)
      }
      # Добавляем mailbox в белые списки отправителей для групп, где контакт был разрешён
      try {
        Write-Host "Добавляю mailbox в Delivery Management групп (без удаления контакта)..."
        Replace-AcceptedSender -OldContactSmtp $UserEmail -NewMailboxId $TempEmail -AddOnly -ErrorAction Stop | Out-Null
      } catch {
        Write-Warning ("Не удалось обновить Delivery Management для {0}: {1}" -f $UserEmail, $_.Exception.Message)
      }
    } else {
      Write-Host "Mailbox не найден. Enable-Mailbox для '$Alias'..."
      Enable-Mailbox -Identity $Alias -PrimarySmtpAddress $UserEmail -Alias $Alias -ErrorAction Stop | Out-Null
      Write-Host "Mailbox включён. Пауза 30 сек для репликации..."
      Start-Sleep -Seconds 30
      if ($Staged) {
        Set-Mailbox -Identity $UserEmail -HiddenFromAddressListsEnabled $true -ErrorAction Stop
      }
    }
  }

  if ($Staged) {
    try {
      Set-ADUser -Identity $Alias -EmailAddress $UserEmail -ErrorAction Stop
      Write-Host "Поле mail обновлено: $UserEmail"
    } catch {
      Write-Warning ("Не удалось обновить поле mail для {0}: {1}" -f $Alias, $_.Exception.Message)
    }
  }

  if ($Activate) {
    try {
      Set-Mailbox -Identity $mailboxIdentity -HiddenFromAddressListsEnabled $false -ErrorAction Stop
      Write-Host "Учетная запись активирована."
    } catch {
      Write-Warning ("Не удалось активировать учетную запись {0}: {1}" -f $Alias, $_.Exception.Message)
    }
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

  } catch {
    Write-Warning "Не удалось привести UPN к $UpnSuffix для '$Alias': $($_.Exception.Message)"
  }

  # Включаем IMAP
  $cas = Get-CASMailbox -Identity $mailboxIdentity -ErrorAction Stop
  if (-not $cas.ImapEnabled) {
    Set-CASMailbox -Identity $mailboxIdentity -ImapEnabled $true -ErrorAction Stop
    Write-Host "IMAP включён."
  } else {
    Write-Host "IMAP уже включён."
  }

  # FullAccess для админа
  $perm = Get-MailboxPermission -Identity $mailboxIdentity | Where-Object { $_.User -eq $AdminLogin -and $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }
  if (-not $perm) {
    Add-MailboxPermission -Identity $mailboxIdentity -User $AdminLogin -AccessRights FullAccess -InheritanceType All -AutoMapping:$false | Out-Null
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
set -uo pipefail
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
DELAY=5
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
    --tmpdir __CACHE_DIR__ \
    --usecache \
    --useheader 'Message-Id' \
    --delete2duplicates \
    --timeout1 45 --timeout2 45 \
    --logfile "$LOGFILE"

  rc=$?
  echo "[imapsync] exit code: $rc"
  if [ $rc -eq 0 ]; then break; fi
  # Retry only on Exchange (host2) authentication failure
  if grep -q 'Host2 failure: Error login on' "$LOGFILE"; then
    if [ $attempt -lt $TRIES ]; then
      echo "[imapsync] host2 auth failed, retry after ${DELAY}s ..."
      sleep $DELAY
      attempt=$((attempt+1))
      continue
    fi
  fi
  # Not an auth failure or retries exhausted
  break
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
    "__CACHE_DIR__"           = $cache_dir
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
    $lineBuf = ""
    try {
      $exitCode = $null
      while ($true) {
        if ($stream.DataAvailable) {
          $chunk = $stream.Read()
          if ($chunk) {
            Write-Host -NoNewline $chunk
            $writer.Write($chunk); $writer.Flush()
            # Try to parse full lines from chunk to detect imapsync total/progress
            $lineBuf += $chunk
            while ($lineBuf -match "^(.*?)(?:`r?`n)") {
              $line  = $matches[1]
              $lineBuf = $lineBuf.Substring($matches[0].Length)
              try {
                if ($line -match 'Progression\s*:\s*(\d+)\s*/\s*(\d+)') {
                  Write-Host ("__IMAPSYNC_PROGRESS__:{0}/{1}" -f $matches[1], $matches[2])
                } elseif ($line -match '(?i)(?:messages\s+copied|copied)\s*:?\s*(\d+)\s*/\s*(\d+)') {
                  Write-Host ("__IMAPSYNC_PROGRESS__:{0}/{1}" -f $matches[1], $matches[2])
                }
              } catch {}
            }
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
