# Экспортирует: Ensure-Module, New-SSHSess

function Ensure-Module([string]$Name) {
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    throw "Модуль $Name не найден. Установи: Install-Module $Name -Scope AllUsers -Force"
  }
  Import-Module $Name -ErrorAction Stop
}

function New-SSHSess([string]$SshHost,[string]$SshUser,[string]$SshPass) {
  $sec  = ConvertTo-SecureString $SshPass -AsPlainText -Force
  $cred = New-Object System.Management.Automation.PSCredential ($SshUser, $sec)
  $params = @{ ComputerName = $SshHost; Credential = $cred }
  $cmd = Get-Command New-SSHSession
  if ($cmd.Parameters.ContainsKey('AcceptKey'))         { $params['AcceptKey'] = $true }
  if ($cmd.Parameters.ContainsKey('ConnectionTimeout')) { $params['ConnectionTimeout'] = 30 }
  $res = New-SSHSession @params
  if ($res -is [System.Array]) { $res = $res[0] }
  if (-not $res) { throw "Не удалось открыть SSH к $SshHost" }
  return $res
}

