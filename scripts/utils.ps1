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

function Get-DistributionGroupsByMember([string]$Identity) {
  $recipient = $null
  try {
    if ($Identity -like '*,*') {
      $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
    } else {
      $recipient = Get-Recipient -Filter "EmailAddresses -eq '$Identity' -or PrimarySmtpAddress -eq '$Identity'" -ErrorAction Stop
    }
  } catch {}
  if (-not $recipient) { return @() }
  $userEmail = $recipient.PrimarySmtpAddress
  $userDN    = $recipient.DistinguishedName
  $groups = Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
    $dg = $_
    try {
      $members = Get-DistributionGroupMember $dg.Identity -ResultSize Unlimited
      if ($members.PrimarySmtpAddress -contains $userEmail -or $members.DistinguishedName -contains $userDN) {
        [PSCustomObject]@{
          DisplayName        = $dg.DisplayName
          PrimarySmtpAddress = $dg.PrimarySmtpAddress
          DistinguishedName  = $dg.DistinguishedName
        }
      }
    } catch {}
  } | Where-Object { $_ } | Sort-Object DisplayName
  return ,$groups
}

