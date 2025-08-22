# Требует: config.ps1, utils.ps1
# Экспортирует: Update-PMGTransport

function Update-PMGTransport([string]$UserEmail) {
  Ensure-Module Posh-SSH
  $pmgSess = New-SSHSess -SshHost $PMGHost -SshUser $PMGUser -SshPass $PMGPasswordPlain
  if (-not $pmgSess) { return @{ Success=$false; Error="SSH к PMG не установлен" } }

  try {
    $line = "$UserEmail smtp:[$ExchangeImapHost]:25"
    $pmgScript = @'
#!/usr/bin/env bash
set -euo pipefail
FILE="/etc/pmg/transport"
ADDR="__ADDR__"
LINE="__LINE__"
if grep -q -E "^${ADDR}[[:space:]]" "$FILE"; then
  sed -i "s|^${ADDR}[[:space:]].*|${LINE}|" "$FILE"
else
  echo "${LINE}" >> "$FILE"
fi
postmap "$FILE"
systemctl reload postfix
'@
    $pmgScript = $pmgScript.Replace("__ADDR__", $UserEmail).Replace("__LINE__", $line)
    $pmgScriptLF = $pmgScript -replace "`r?`n", "`n"
    $pmgB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pmgScriptLF))
    $pmgRemote = "/tmp/pmg-transport-$(( $UserEmail -replace '@','_' )).sh"
    $pmgUpload = "bash -lc 'printf %s ""$pmgB64"" | base64 -d > ""$pmgRemote"" && chmod 700 ""$pmgRemote"" && bash ""$pmgRemote""; rc=`$?; rm -f ""$pmgRemote""; printf ""__PMG_END__:%s\n"" `\$rc'"
    $pmgRun = Invoke-SSHCommand -SessionId $pmgSess.SessionId -Command $pmgUpload

    $ok = $false
    if ($pmgRun.Output -and ($pmgRun.Output -join "`n") -match "__PMG_END__:(\d+)") {
      $ok = ([int]$matches[1] -eq 0)
    } else {
      $ok = ($pmgRun.ExitStatus -eq 0)
    }
    if ($ok) { return @{ Success=$true; Line=$line } }
    else { return @{ Success=$false; Error="Команда завершилась с ошибкой: $($pmgRun.Error -join ' ')" } }
  } finally {
    Remove-SSHSession -SessionId $pmgSess.SessionId | Out-Null
  }
}
