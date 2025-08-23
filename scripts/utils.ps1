# Экспортирует: Ensure-Module, New-SSHSess, Get-DistributionGroupsByMember

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

function Get-DistributionGroupsByMember([string]$mail) {
  if (-not $mail) { return @() }

  try {
    $recipient = Get-Recipient -Identity $mail -ErrorAction Stop
  } catch {
    return @()
  }

  $dn = $recipient.DistinguishedName
  $escapedDn = $dn -replace "'", "''"

  # 1) пробуем OPATH (быстро)
  $groups = @()
  try {
    $groups = Get-DistributionGroup -Filter "Members -eq '$escapedDn'" -ResultSize Unlimited
  } catch {
    $groups = @()
  }

  # 2) если пусто — резерв через AD/LDAP (надёжно)
  if (-not $groups -or $groups.Count -eq 0) {
    try {
      $groups = Get-ADGroup -LDAPFilter "(member=$dn)" -Properties mail,displayName,distinguishedName
    } catch {
      $groups = @()
    }
  }

  $groups | Select-Object `
    @{n='DisplayName';e={$_.DisplayName}},
    @{n='PrimarySmtpAddress';e={ if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress } else { $_.mail } }},
    @{n='DistinguishedName';e={$_.DistinguishedName}}
}
