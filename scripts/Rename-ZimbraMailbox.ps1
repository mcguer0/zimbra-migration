# Требует: config.ps1, utils.ps1
# Экспортирует: Rename-ZimbraMailbox

function Rename-ZimbraMailbox([string]$UserEmail,[string]$Alias) {
  Ensure-Module Posh-SSH
  $zSess = New-SSHSess -SshHost $ZimbraSshHost -SshUser $ZimbraSshUser -SshPass $ZimbraSshPasswordPlain
  if (-not $zSess) { return @{ Success=$false; Error="SSH к $ZimbraSshHost не установлен" } }

  $oldEmail = "{0}_old@{1}" -f $Alias, $Domain
  try {
    $cmd = ("bash -lc 'su - zimbra -c ""zmprov ra {0} {1}""'") -f $UserEmail, $oldEmail
    $rn = Invoke-SSHCommand -SessionId $zSess.SessionId -Command $cmd
    if ($rn.ExitStatus -ne 0) {
      $oldEmail = "{0}_old_{1}@{2}" -f $Alias, (Get-Date -Format "yyyyMMddHHmm"), $Domain
      $cmd2 = ("bash -lc 'su - zimbra -c ""zmprov ra {0} {1}""'") -f $UserEmail, $oldEmail
      $rn2 = Invoke-SSHCommand -SessionId $zSess.SessionId -Command $cmd2
      if ($rn2.ExitStatus -ne 0) {
        return @{ Success=$false; Error=($rn2.Error -join ' ') }
      }
    }
    return @{ Success=$true; NewEmail=$oldEmail }
  } finally {
    Remove-SSHSession -SessionId $zSess.SessionId | Out-Null
  }
}
