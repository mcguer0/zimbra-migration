[CmdletBinding(DefaultParameterSetName="Export")]
param(
  [Parameter(ParameterSetName="Export", Mandatory=$true, HelpMessage="���� � CSV ��� ��������")]
  [string]$Export,

  [Parameter(ParameterSetName="Export")]
  [string]$SourceOU = "OU=������-�����,DC=metall-zavod,DC=local",

  [Parameter(ParameterSetName="Import", Mandatory=$true, HelpMessage="���� � CSV ��� �������")]
  [string]$Import,

  [Parameter(ParameterSetName="Import")]
  [string]$TargetOU = "OU=ZimbraContactsForExchange,OU=AdressBook,DC=metall-zavod,DC=local"
)

if ($PSCmdlet.ParameterSetName -eq "Export") {
  Get-ADUser -SearchBase $SourceOU `
             -SearchScope Subtree `
             -LDAPFilter '(&(objectCategory=person)(objectClass=user)(mail=*))' `
             -Properties mail,displayName,givenName,sn,sAMAccountName,company,department,title,
                        telephoneNumber,mobile,facsimileTelephoneNumber,streetAddress,l,st,postalCode,co,c,info |
    Select-Object `
        @{n='Name';e={ if ($_.displayName) { $_.displayName } elseif ($_.Name) { $_.Name } else { $_.sAMAccountName } }},
        @{n='DisplayName';e={$_.displayName}},
        @{n='FirstName';e={$_.givenName}},
        @{n='LastName';e={$_.sn}},
        @{n='Alias';e={ ($_.sAMAccountName -replace '[^a-zA-Z0-9._-]','' ).ToLower() }},
        @{n='ExternalEmailAddress';e={$_.mail}},
        @{n='Company';e={$_.company}},
        @{n='Department';e={$_.department}},
        @{n='Title';e={$_.title}},
        @{n='Phone';e={$_.telephoneNumber}},
        @{n='MobilePhone';e={$_.mobile}},
        @{n='Fax';e={$_.facsimileTelephoneNumber}},
        @{n='StreetAddress';e={$_.streetAddress}},
        @{n='City';e={$_.l}},
        @{n='StateOrProvince';e={$_.st}},
        @{n='PostalCode';e={$_.postalCode}},
        @{n='CountryOrRegion';e={ if ($_.co) { $_.co } else { $_.c } }},
        @{n='HiddenFromAddressListsEnabled';e={$false}},
        @{n='Notes';e={ if ($_.info) { $_.info } else { '������������ �� AD' } }} |
    Export-Csv -Path $Export -NoTypeInformation -Encoding UTF8

  Import-Csv $Export | Measure-Object
  return
}

if ($PSCmdlet.ParameterSetName -eq "Import") {
  if (-not (Test-Path $Import)) { throw "CSV not found: $Import" }
  $rows = Import-Csv -Path $Import

  $created=0; $updated=0; $skipped=0

  foreach ($r in $rows) {
      $name       = if ($r.PSObject.Properties.Match('Name').Count)       { [string]$r.Name } else { $null }
      $display    = if ($r.PSObject.Properties.Match('DisplayName').Count){ [string]$r.DisplayName } else { $null }
      $firstName  = if ($r.PSObject.Properties.Match('FirstName').Count)  { [string]$r.FirstName } else { $null }
      $lastName   = if ($r.PSObject.Properties.Match('LastName').Count)   { [string]$r.LastName } else { $null }
      $aliasCsv   = if ($r.PSObject.Properties.Match('Alias').Count)      { [string]$r.Alias } else { $null }
      $email      = if ($r.PSObject.Properties.Match('ExternalEmailAddress').Count) { [string]$r.ExternalEmailAddress } else { $null }

      if ($name)      { $name = $name.Trim() }
      if ($display)   { $display = $display.Trim() }
      if ($firstName) { $firstName = $firstName.Trim() }
      if ($lastName)  { $lastName = $lastName.Trim() }
      if ($aliasCsv)  { $aliasCsv = $aliasCsv.Trim() }
      if ($email)     { $email = $email.Trim() }

      if ([string]::IsNullOrWhiteSpace($email)) {
          Write-Warning ("SKIP: '{0}' � ������ ExternalEmailAddress" -f $name); $skipped++; continue
      }

      $alias = $aliasCsv
      if ([string]::IsNullOrWhiteSpace($alias)) {
          if ($email -match '^(?<local>[^@]+)@') { $alias = $matches['local'] }
      }

      $mc = $null
      $mc = Get-MailContact -Filter "ExternalEmailAddress -eq '$email'" -ErrorAction SilentlyContinue
      if (-not $mc) { $mc = Get-MailContact -Filter "PrimarySmtpAddress -eq '$email'" -ErrorAction SilentlyContinue }
      if (-not $mc -and $alias) { $mc = Get-MailContact -Filter "Alias -eq '$alias'" -ErrorAction SilentlyContinue }
      if (-not $mc -and $name)  { $mc = Get-MailContact -Filter "DisplayName -eq '$name'" -ErrorAction SilentlyContinue }
      if (-not $mc -and $name)  { $mc = Get-MailContact -Filter "Name -eq '$name'" -ErrorAction SilentlyContinue }

      if (-not $mc) {
          try {
              if ($alias) {
                  $mc = New-MailContact -Name $name -ExternalEmailAddress $email -OrganizationalUnit $TargetOU -FirstName $firstName -LastName $lastName -Alias $alias -ErrorAction Stop
              } else {
                  $mc = New-MailContact -Name $name -ExternalEmailAddress $email -OrganizationalUnit $TargetOU -FirstName $firstName -LastName $lastName -ErrorAction Stop
              }
              $created++; Write-Host ("CREATED: {0} <{1}>" -f $name,$email)
          }
          catch {
              $msg = $_.Exception.Message
              if ($msg -match 'already exists|��� ����������|Object .* already exists') {
                  $mc = $null
                  $mc = Get-MailContact -Filter "ExternalEmailAddress -eq '$email'" -ErrorAction SilentlyContinue
                  if (-not $mc) { $mc = Get-MailContact -Filter "PrimarySmtpAddress -eq '$email'" -ErrorAction SilentlyContinue }
                  if (-not $mc -and $alias) { $mc = Get-MailContact -Filter "Alias -eq '$alias'" -ErrorAction SilentlyContinue }
                  if (-not $mc -and $name)  { $mc = Get-MailContact -Filter "DisplayName -eq '$name'" -ErrorAction SilentlyContinue }
                  if (-not $mc -and $name)  { $mc = Get-MailContact -Filter "Name -eq '$name'" -ErrorAction SilentlyContinue }

                  if (-not $mc) {
                      $adContact = $null
                      if ($alias) { $adContact = Get-Contact -Filter "Alias -eq '$alias'" -ErrorAction SilentlyContinue }
                      if (-not $adContact -and $name) { $adContact = Get-Contact -Filter "DisplayName -eq '$name'" -ErrorAction SilentlyContinue }
                      if (-not $adContact -and $name) { $adContact = Get-Contact -Filter "Name -eq '$name'" -ErrorAction SilentlyContinue }

                      if ($adContact) {
                          try {
                              Enable-MailContact -Identity $adContact.Identity -ExternalEmailAddress $email -ErrorAction Stop | Out-Null
                              $mc = Get-MailContact -Identity $adContact.Identity -ErrorAction SilentlyContinue
                              Write-Warning ("MAIL-ENABLED existing AD contact: '{0}' <{1}>" -f $name,$email)
                          } catch {
                              Write-Warning ("ERROR mail-enabling '{0}': {1}" -f $name,$_.Exception.Message)
                          }
                      }
                  }

                  if (-not $mc) {
                      Write-Warning ("CONFLICT: '{0}' <{1}> � ������ � ����� ������ ����������, �� �� ����� �� ������. �������." -f $name,$email)
                      $skipped++; continue
                  } else {
                      Write-Warning ("FOUND-EXISTING: '{0}' ��� ����������; ��������." -f $name)
                  }
              } else {
                  Write-Warning ("ERROR create '{0}' <{1}>: {2}" -f $name,$email,$msg)
                  $skipped++; continue
              }
          }
      }

      $company    = if ($r.PSObject.Properties.Match('Company').Count)    { [string]$r.Company } else { $null }
      $department = if ($r.PSObject.Properties.Match('Department').Count) { [string]$r.Department } else { $null }
      $title      = if ($r.PSObject.Properties.Match('Title').Count)      { [string]$r.Title } else { $null }
      $phone      = if ($r.PSObject.Properties.Match('Phone').Count)      { [string]$r.Phone } else { $null }
      $mobile     = if ($r.PSObject.Properties.Match('MobilePhone').Count){ [string]$r.MobilePhone } else { $null }
      $fax        = if ($r.PSObject.Properties.Match('Fax').Count)        { [string]$r.Fax } else { $null }
      $street     = if ($r.PSObject.Properties.Match('StreetAddress').Count){ [string]$r.StreetAddress } else { $null }
      $city       = if ($r.PSObject.Properties.Match('City').Count)       { [string]$r.City } else { $null }
      $state      = if ($r.PSObject.Properties.Match('StateOrProvince').Count){ [string]$r.StateOrProvince } else { $null }
      $postal     = if ($r.PSObject.Properties.Match('PostalCode').Count) { [string]$r.PostalCode } else { $null }
      $country    = if ($r.PSObject.Properties.Match('CountryOrRegion').Count){ [string]$r.CountryOrRegion } else { $null }
      $notes      = if ($r.PSObject.Properties.Match('Notes').Count)      { [string]$r.Notes } else { $null }

      $hidden = $null
      if ($r.PSObject.Properties.Match('HiddenFromAddressListsEnabled').Count -and $r.HiddenFromAddressListsEnabled) {
          $v = ($r.HiddenFromAddressListsEnabled.ToString()).Trim().ToLower()
          if ($v -eq 'true') { $hidden = $true } elseif ($v -eq 'false') { $hidden = $false }
      }

      try {
          if ($display) { Set-Contact -Identity $mc.Identity -DisplayName $display -ErrorAction SilentlyContinue | Out-Null }

          Set-Contact -Identity $mc.Identity `
              -FirstName $firstName `
              -LastName $lastName `
              -Company $company `
              -Department $department `
              -Title $title `
              -Phone $phone `
              -MobilePhone $mobile `
              -Fax $fax `
              -StreetAddress $street `
              -City $city `
              -StateOrProvince $state `
              -PostalCode $postal `
              -CountryOrRegion $country `
              -Notes $notes `
              -ErrorAction SilentlyContinue | Out-Null

          Set-MailContact -Identity $mc.Identity -ExternalEmailAddress $email -ErrorAction SilentlyContinue | Out-Null
          if ($alias -and $mc.Alias -ne $alias) { Set-MailContact -Identity $mc.Identity -Alias $alias -ErrorAction SilentlyContinue | Out-Null }
          if ($hidden -ne $null) { Set-MailContact -Identity $mc.Identity -HiddenFromAddressListsEnabled:$hidden -ErrorAction SilentlyContinue | Out-Null }

          $updated++; Write-Host ("UPDATED: {0} <{1}>" -f $mc.DisplayName,$email)
      }
      catch {
          Write-Warning ("ERROR update '{0}' <{1}>: {2}" -f $name,$email,$_.Exception.Message)
          $skipped++
      }
  }

  Write-Host '==== Summary ===='
  Write-Host ("Created: {0}" -f $created)
  Write-Host ("Updated: {0}" -f $updated)
  Write-Host ("Skipped: {0}" -f $skipped)
}
