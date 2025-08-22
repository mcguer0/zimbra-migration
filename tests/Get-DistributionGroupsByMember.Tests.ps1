$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here/../scripts/utils.ps1"

Describe "Get-DistributionGroupsByMember" {
  Context "AD lookup" {
    Mock Get-Recipient { [pscustomobject]@{ DistinguishedName='CN=User,DC=example,DC=com' } }
    Mock Get-ADGroup { param($LDAPFilter) @([pscustomobject]@{ DisplayName='GroupA'; mail='groupa@example.com'; DistinguishedName='CN=GroupA,DC=example,DC=com' }) }
    It "returns groups with required properties" {
      $res = Get-DistributionGroupsByMember 'user@example.com'
      $res | Should -HaveCount 1
      $res[0].DisplayName | Should -Be 'GroupA'
      $res[0].PrimarySmtpAddress | Should -Be 'groupa@example.com'
      $res[0].DistinguishedName | Should -Be 'CN=GroupA,DC=example,DC=com'
    }
  }

  Context "fallback enumeration" {
    Mock Get-Recipient { [pscustomobject]@{ DistinguishedName='CN=User,DC=example,DC=com'; PrimarySmtpAddress='user@example.com' } }
    Mock Get-ADGroup { throw 'AD not available' }
    Mock Get-DistributionGroup { @([pscustomobject]@{ Identity='GroupB'; DisplayName='GroupB'; PrimarySmtpAddress='groupb@example.com'; DistinguishedName='CN=GroupB,DC=example,DC=com' }) }
    Mock Get-DistributionGroupMember { param($Identity) @([pscustomobject]@{ DistinguishedName='CN=User,DC=example,DC=com' }) }
    It "uses distribution group enumeration" {
      $res = Get-DistributionGroupsByMember 'user@example.com'
      $res | Should -HaveCount 1
      $res[0].DisplayName | Should -Be 'GroupB'
    }
  }
}

Describe "Scripts use Get-DistributionGroupsByMember" {
  It "Move-ZimbraMailbox uses helper" {
    $content = Get-Content "$here/../scripts/Move-ZimbraMailbox.ps1" -Raw
    $content | Should -Match 'Get-DistributionGroupsByMember'
  }
  It "DryRun-ZimbraMailbox uses helper" {
    $content = Get-Content "$here/../scripts/DryRun-ZimbraMailbox.ps1" -Raw
    $content | Should -Match 'Get-DistributionGroupsByMember'
  }
}
