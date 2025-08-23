[CmdletBinding(DefaultParameterSetName="Single")]
param(
  # Режим 1: один пользователь
  [Parameter(ParameterSetName="Single", Mandatory=$true, HelpMessage="user или user@example.com")]
  [string]$User,

  # Режим 2: пакетный (файл со списком пользователей)
  [Parameter(ParameterSetName="Batch", Mandatory=$true, HelpMessage="Путь к файлу users.txt")]
  [string]$Path,

  # Без вопросов (автоматически переименовывать и писать transport)
  [switch]$Force,

  # Сухой прогон без переноса
  [switch]$Dryrun
)

# == Подключаем конфиг и функции ==
. "$PSScriptRoot\scripts\config.ps1"           # переменные окружения
. "$PSScriptRoot\scripts\utils.ps1"            # Ensure-Module, New-SSHSess
. "$PSScriptRoot\scripts\Move-ZimbraMailbox.ps1"
. "$PSScriptRoot\scripts\DryRun-ZimbraMailbox.ps1"
. "$PSScriptRoot\scripts\Rename-ZimbraMailbox.ps1"
. "$PSScriptRoot\scripts\Update-PMGTransport.ps1"

# == Локальные функции, специфичные для "основного" скрипта ==

function Ensure-ExchangeCmdlets {
  if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
  Write-Host "Подгружаю Exchange cmdlets..."
  try {
    $snap = Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "Exchange"}
    if ($snap) {
      foreach ($s in $snap) { Add-PSSnapin $s.Name -ErrorAction Stop }
      if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) { return }
    }
  } catch {}
  try {
    $sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeMgmtHost) -Authentication Kerberos -ErrorAction Stop
    Import-PSSession $sess -DisableNameChecking | Out-Null
    return
  } catch {
    throw "Не удалось загрузить Exchange cmdlets (ни snapin, ни remoting). Проверь запуск на сервере Exchange."
  }
}

# == Инициализация окружения ==
Ensure-ExchangeCmdlets
Ensure-Module Posh-SSH
Ensure-Module ActiveDirectory
if (-not (Test-Path $LocalLogDir)) { New-Item -Path $LocalLogDir -ItemType Directory -Force | Out-Null }

# == Обработчик одного пользователя ==
function Invoke-OneUser([string]$UserInput) {
  if ($Dryrun) {
    Invoke-DryRunZimbraMailbox -UserInput $UserInput | Out-Null
    return
  }

  # Выполняем перенос
  $move = Invoke-MoveZimbraMailbox -UserInput $UserInput
  $UserEmail = $move.UserEmail
  $Alias     = $move.Alias

  if ($move.ExitCode -eq 0) {
    Write-Host "`n✅ Миграция $UserEmail завершена успешно. Лог: $($move.LocalLog)"

    if ($Force) {
      $pmg = Update-PMGTransport -UserEmail $UserEmail
      if ($pmg.Success) { Write-Host "Транспорт обновлён: $($pmg.Line)" }
      else { Write-Warning "Не удалось обновить transport на PMG: $($pmg.Error)" }

      $rename = Rename-ZimbraMailbox -UserEmail $UserEmail -Alias $Alias
      if ($rename.Success) {
        Write-Host "Переименовано в $($rename.NewEmail)"
        # после переименования исходный адрес освобождается — транспорт всё равно на старый адрес
      } else {
        Write-Warning "Не удалось переименовать: $($rename.Error)"
      }
    } else {
      Write-Host "Переименование Zimbra-аккаунта и обновление transport на PMG пропущены (не указан -Force)."
    }
  } else {
    Write-Warning "`n⚠️ Миграция $UserEmail завершилась с кодом $($move.ExitCode). См. лог: $($move.LocalLog)"
  }
}

# == Точка входа ==
if ($PSCmdlet.ParameterSetName -eq "Single") {
  Invoke-OneUser -UserInput $User
} else {
  if (-not (Test-Path $Path)) { throw "Файл не найден: $Path" }
  $list = Get-Content -Path $Path -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith("#") }
  foreach ($u in $list) {
    Write-Host ""
    Write-Host "==================== $u ====================" -ForegroundColor Cyan
    try { Invoke-OneUser -UserInput $u } catch { Write-Warning "Ошибка при миграции $u : $($_.Exception.Message)" }
  }
  Write-Host "`nГотово."
}
