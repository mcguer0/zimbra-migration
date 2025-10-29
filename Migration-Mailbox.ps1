<#
  Migration-Mailbox.ps1 — запуск миграции одного пользователя либо пакета (-Path)
  Переписано с нуля. Текст на русском в UTF‑8 with BOM.
#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
  # Режим 1: одиночная миграция
  [Parameter(ParameterSetName = 'Single', Mandatory = $true, HelpMessage = 'user или user@example.com')]
  [string]$User,

  # Режим 2: пакетная миграция из файла
  [Parameter(ParameterSetName = 'Batch', Mandatory = $true, HelpMessage = 'Путь к списку users.txt')]
  [string]$Path,

  # Форсировать финальные шаги (активация, PMG, переименование)
  [switch]$Force,

  # «Сухой» прогон без внесения изменений
  [switch]$Dryrun,

  # Постепенная миграция (staged)
  [switch]$Staged
)

# == Конфиг и функции ==
. "$PSScriptRoot\scripts\config.ps1"
. "$PSScriptRoot\scripts\utils.ps1"             # Ensure-Module, New-SSHSess
. "$PSScriptRoot\scripts\Move-ZimbraMailbox.ps1"
. "$PSScriptRoot\scripts\DryRun-ZimbraMailbox.ps1"
. "$PSScriptRoot\scripts\Rename-ZimbraMailbox.ps1"
. "$PSScriptRoot\scripts\Update-PMGTransport.ps1"
. "$PSScriptRoot\scripts\Replace-AcceptedSender.ps1"

# == Доступ к Exchange cmdlets ==
function Ensure-ExchangeCmdlets {
  if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
  Write-Host 'Подключаю Exchange cmdlets...'
  try {
    $snap = Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Exchange' }
    if ($snap) {
      foreach ($s in $snap) { Add-PSSnapin $s.Name -ErrorAction Stop }
      if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
    }
  } catch {}
  try {
    $sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeMgmtHost) -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $sess -DisableNameChecking | Out-Null
  } catch {
    throw 'Не удалось подключить Exchange cmdlets (ни snapin, ни remoting). Запустите скрипт на хосте Exchange.'
  }
}

# == Инициализация окружения ==
Ensure-ExchangeCmdlets
Ensure-Module Posh-SSH
Ensure-Module ActiveDirectory
if (-not (Test-Path $LocalLogDir)) { New-Item -Path $LocalLogDir -ItemType Directory -Force | Out-Null }

# == Простой индикатор прогресса по количеству job ==

# == Миграция одного пользователя ==
function Invoke-OneUser([string]$UserInput) {
  if ($Dryrun) {
    Invoke-DryRunZimbraMailbox -UserInput $UserInput | Out-Null
    return
  }

  $move = Invoke-MoveZimbraMailbox -UserInput $UserInput -Staged:$Staged -Activate:$Force
  $UserEmail = $move.UserEmail
  $Alias     = $move.Alias

  if ($move.ExitCode -eq 0) {
    Write-Host ("`n✅ Почтовый ящик {0} мигрирован успешно. Лог: {1}" -f $UserEmail, $move.LocalLog)

    if ($Force) {
      $pmg = Update-PMGTransport -UserEmail $UserEmail
      if ($pmg.Success) { Write-Host ("Обновлён transport в PMG: {0}" -f $pmg.Line) }
      else { Write-Warning ("Не удалось обновить transport в PMG: {0}" -f $pmg.Error) }

      $rename = Rename-ZimbraMailbox -UserEmail $UserEmail -Alias $Alias
      if ($rename.Success) { Write-Host ("Переименовано в {0}" -f $rename.NewEmail) }
      else { Write-Warning ("Не удалось переименовать: {0}" -f $rename.Error) }
    } else {
      Write-Host 'Финальные шаги (PMG/переименование) выполняются ключом -Force.'
    }
  } else {
    Write-Warning ("`n❌ Почтовый ящик {0} завершился с кодом {1}. См. лог: {2}" -f $UserEmail, $move.ExitCode, $move.LocalLog)
  }
}

# == Точка входа ==
if ($PSCmdlet.ParameterSetName -eq 'Single') {
  $jobs = @()
  $state = @{}
  $u = $User
  Write-Host ''
  Write-Host ("==================== {0} ====================" -f $u) -ForegroundColor Cyan
  $jobs += Start-Job -Name ("mig-" + ($u -replace "[^\w\.-]","_")) -ArgumentList @(
      $u, $PSScriptRoot, [bool]$Force, [bool]$Dryrun, [bool]$Staged
    ) -ScriptBlock {
      param($UserInput, $Root, [bool]$Force, [bool]$Dryrun, [bool]$Staged)

      . "$Root\scripts\config.ps1"
      . "$Root\scripts\utils.ps1"
      . "$Root\scripts\Move-ZimbraMailbox.ps1"
      . "$Root\scripts\DryRun-ZimbraMailbox.ps1"
      . "$Root\scripts\Rename-ZimbraMailbox.ps1"
      . "$Root\scripts\Update-PMGTransport.ps1"
      . "$Root\scripts\Replace-AcceptedSender.ps1"

      function Ensure-ExchangeCmdlets {
        if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
        try {
          $snap = Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Exchange' }
          if ($snap) {
            foreach ($s in $snap) { Add-PSSnapin $s.Name -ErrorAction SilentlyContinue }
            if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
          }
        } catch {}
        try {
          $sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeMgmtHost) -Authentication Kerberos -ErrorAction Stop
          Import-PSSession $sess -DisableNameChecking | Out-Null
        } catch {
          throw 'Не удалось подключить Exchange cmdlets внутри фонового задания.'
        }
      }

      Ensure-ExchangeCmdlets
      Ensure-Module Posh-SSH
      Ensure-Module ActiveDirectory
      if (-not (Test-Path $LocalLogDir)) { New-Item -Path $LocalLogDir -ItemType Directory -Force | Out-Null }

      if ($Dryrun) {
        Invoke-DryRunZimbraMailbox -UserInput $UserInput | Out-Null
        return
      }

      $move = Invoke-MoveZimbraMailbox -UserInput $UserInput -Staged:$Staged -Activate:$Force
      $UserEmail = $move.UserEmail
      $Alias     = $move.Alias

      if ($move.ExitCode -eq 0) {
        Write-Host ("`n✅ Почтовый ящик {0} мигрирован успешно. Лог: {1}" -f $UserEmail, $move.LocalLog)
        if ($Force) {
          $pmg = Update-PMGTransport -UserEmail $UserEmail
          if ($pmg.Success) { Write-Host ("Обновлён transport в PMG: {0}" -f $pmg.Line) } else { Write-Warning ("Не удалось обновить transport в PMG: {0}" -f $pmg.Error) }
          $rename = Rename-ZimbraMailbox -UserEmail $UserEmail -Alias $Alias
          if ($rename.Success) { Write-Host ("Переименовано в {0}" -f $rename.NewEmail) } else { Write-Warning ("Не удалось переименовать: {0}" -f $rename.Error) }
        } else {
          Write-Host 'Финальные шаги (PMG/переименование) выполняются ключом -Force.'
        }
      } else {
        Write-Warning ("`n❌ Почтовый ящик {0} завершился с кодом {1}. См. лог: {2}" -f $UserEmail, $move.ExitCode, $move.LocalLog)
      }
    }

  # Простой прогресс: количество завершённых заданий из общего числа
  Write-Host ("`nВсего запущенных заданий: {0}" -f $jobs.Count)
  $total = $jobs.Count
  $progressId = 1
  $activity = 'Миграция почтовых ящиков'
  try {
    while ($true) {
      $current = Get-Job -Id ($jobs | Select-Object -ExpandProperty Id)
      # Строка прогресса под пользователем из маркеров дочерней задачи
      foreach ($j in $current) {
        $out = Receive-Job -Job $j -Keep -ErrorAction SilentlyContinue
        if ($out) {
          foreach ($line in ($out | ForEach-Object { $_.ToString() })) {
            if ($line -match '^__IMAPSYNC_PROGRESS__:(\d+)/(\d+)') {
              $d = [int]$matches[1]; $t = [int]$matches[2]
              Write-Host ("  {0}: {1}/{2}" -f $u, $d, $t)
            }
          }
        }
      }
      $done    = ($current | Where-Object { $_.State -in 'Completed','Failed','Stopped' }).Count
      $running = ($current | Where-Object { $_.State -eq 'Running' }).Count
      $percent = if ($total -gt 0) { [int](($done / $total) * 100) } else { 100 }
      $status  = ([string]::Format('{0}/{1} завершено; выполняется: {2}', $done, $total, $running))
      Write-Progress -Id $progressId -Activity $activity -Status $status -PercentComplete $percent
      if ($done -ge $total) { break }
      Start-Sleep -Milliseconds 500
    }
  } finally {
    Write-Progress -Id $progressId -Activity $activity -Completed
  }

  foreach ($j in $jobs) { Receive-Job -Job $j -Keep | Write-Output }
  Remove-Job -Job $jobs -Force -ErrorAction SilentlyContinue | Out-Null
  return
}

# Пакетный режим (-Path)
if (-not (Test-Path $Path)) { throw ("Файл не найден: {0}" -f $Path) }
$list = Get-Content -Path $Path -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
if (-not $list -or $list.Count -eq 0) { Write-Host 'Список пуст.'; return }

$jobs = @()
$userMap = @{}
foreach ($u in $list) {
  Write-Host ''
  Write-Host ("==================== {0} ====================" -f $u) -ForegroundColor Cyan
  $job = Start-Job -Name ("mig-" + ($u -replace "[^\w\.-]","_")) -ArgumentList @(
      $u, $PSScriptRoot, [bool]$Force, [bool]$Dryrun, [bool]$Staged
    ) -ScriptBlock {
      param($UserInput, $Root, [bool]$Force, [bool]$Dryrun, [bool]$Staged)

      # Подгружаем зависимости в контексте job
      . "$Root\scripts\config.ps1"
      . "$Root\scripts\utils.ps1"
      . "$Root\scripts\Move-ZimbraMailbox.ps1"
      . "$Root\scripts\DryRun-ZimbraMailbox.ps1"
      . "$Root\scripts\Rename-ZimbraMailbox.ps1"
      . "$Root\scripts\Update-PMGTransport.ps1"
      . "$Root\scripts\Replace-AcceptedSender.ps1"

      function Ensure-ExchangeCmdlets {
        if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
        try {
          $snap = Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Exchange' }
          if ($snap) {
            foreach ($s in $snap) { Add-PSSnapin $s.Name -ErrorAction SilentlyContinue }
            if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
          }
        } catch {}
        try {
          $sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeMgmtHost) -Authentication Kerberos -ErrorAction Stop
          Import-PSSession $sess -DisableNameChecking | Out-Null
        } catch {
          throw 'Не удалось подключить Exchange cmdlets внутри фонового задания.'
        }
      }

      Ensure-ExchangeCmdlets
      Ensure-Module Posh-SSH
      Ensure-Module ActiveDirectory
      if (-not (Test-Path $LocalLogDir)) { New-Item -Path $LocalLogDir -ItemType Directory -Force | Out-Null }

      if ($Dryrun) {
        Invoke-DryRunZimbraMailbox -UserInput $UserInput | Out-Null
        return
      }

      $move = Invoke-MoveZimbraMailbox -UserInput $UserInput -Staged:$Staged -Activate:$Force
      $UserEmail = $move.UserEmail
      $Alias     = $move.Alias

      if ($move.ExitCode -eq 0) {
        Write-Host ("`n✅ Почтовый ящик {0} мигрирован успешно. Лог: {1}" -f $UserEmail, $move.LocalLog)
        if ($Force) {
          $pmg = Update-PMGTransport -UserEmail $UserEmail
          if ($pmg.Success) { Write-Host ("Обновлён transport в PMG: {0}" -f $pmg.Line) } else { Write-Warning ("Не удалось обновить transport в PMG: {0}" -f $pmg.Error) }
          $rename = Rename-ZimbraMailbox -UserEmail $UserEmail -Alias $Alias
          if ($rename.Success) { Write-Host ("Переименовано в {0}" -f $rename.NewEmail) } else { Write-Warning ("Не удалось переименовать: {0}" -f $rename.Error) }
        } else {
          Write-Host 'Финальные шаги (PMG/переименование) выполняются ключом -Force.'
        }
      } else {
        Write-Warning ("`n❌ Почтовый ящик {0} завершился с кодом {1}. См. лог: {2}" -f $UserEmail, $move.ExitCode, $move.LocalLog)
      }
    }
  $jobs += $job
  $userMap[$job.Id] = $u
}

if ($jobs.Count -gt 0) {
  Write-Host ("`nВсего запущенных заданий: {0}" -f $jobs.Count)
  # Простой индикатор прогресса ожидания
  $total = $jobs.Count
  $progressId = 1
  $activity = 'Миграция почтовых ящиков'
  try {
    while ($true) {
      $current = Get-Job -Id ($jobs | Select-Object -ExpandProperty Id)
      # Выводим строку прогресса под каждым пользователем, если пришли маркеры
      foreach ($j in $current) {
        $out = Receive-Job -Job $j -Keep -ErrorAction SilentlyContinue
        if ($out) {
          foreach ($line in ($out | ForEach-Object { $_.ToString() })) {
            if ($line -match '^__IMAPSYNC_PROGRESS__:(\d+)/(\d+)') {
              $d = [int]$matches[1]; $t = [int]$matches[2]
              $usr = if ($userMap.ContainsKey($j.Id)) { $userMap[$j.Id] } else { $j.Name }
              Write-Host ("  {0}: {1}/{2}" -f $usr, $d, $t)
            }
          }
        }
      }
      $done    = ($current | Where-Object { $_.State -in 'Completed','Failed','Stopped' }).Count
      $running = ($current | Where-Object { $_.State -eq 'Running' }).Count
      $percent = if ($total -gt 0) { [int](($done / $total) * 100) } else { 100 }
      $status  = ([string]::Format('{0}/{1} завершено; выполняется: {2}', $done, $total, $running))
      Write-Progress -Id $progressId -Activity $activity -Status $status -PercentComplete $percent
      if ($done -ge $total) { break }
      Start-Sleep -Milliseconds 500
    }
  } finally {
    Write-Progress -Id $progressId -Activity $activity -Completed
  }

  foreach ($j in $jobs) { Receive-Job -Job $j -Keep | Write-Output }
  Remove-Job -Job $jobs -Force -ErrorAction SilentlyContinue | Out-Null
}

Write-Host "Готово."
