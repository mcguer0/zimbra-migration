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

function Get-DistributionGroupsByMember {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [string]$Identity
  )

  if (-not $Identity) { return @() }

  $dn = $null
  $email = $null

  if ($Identity -like '*@*') {
    $email = $Identity
    try {
      $recipient = Get-Recipient -Filter "EmailAddresses -eq '$Identity' -or PrimarySmtpAddress -eq '$Identity'" -ErrorAction Stop
      if ($recipient) { $dn = $recipient.DistinguishedName }
    } catch {}
  } else {
    $dn = $Identity
  }

  $groups = @()

  if ($dn) {
    try {
      $groups = Get-ADGroup -LDAPFilter "(member=$dn)" -Properties mail,displayName |
        Select-Object @{Name='DisplayName';Expression={$_.DisplayName}}, @{Name='PrimarySmtpAddress';Expression={$_.mail}}, DistinguishedName |
        Sort-Object DisplayName
      return $groups
    } catch {}
  }

  try {
    $groups = Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
      $dg = $_
      try {
        $members = Get-DistributionGroupMember $dg.Identity -ResultSize Unlimited
        if ($dn) {
          if ($members.DistinguishedName -contains $dn) { $dg }
        } elseif ($email) {
          if ($members.PrimarySmtpAddress -contains $email) { $dg }
        }
      } catch {}
    } | Where-Object { $_ } |
      Select-Object DisplayName, PrimarySmtpAddress, DistinguishedName |
      Sort-Object DisplayName
  } catch {}

  return $groups
}
