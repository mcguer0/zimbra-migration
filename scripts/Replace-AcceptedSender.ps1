# Экспортирует: Replace-AcceptedSender

function Replace-AcceptedSender {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true, HelpMessage="SMTP контакта, который использовался как отправитель")]
        [string] $OldContactSmtp,

        [Parameter(Mandatory=$true, HelpMessage="Идентификатор нового почтового ящика (UPN/SMTP/Alias/DN)")]
        [string] $NewMailboxId,

        [Parameter(HelpMessage="Только добавить ящик, контакт не удалять из белого списка")]
        [switch] $AddOnly
    )

    function Ensure-ExchangeCmdlets {
        if (Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue) { return }
        Write-Host "Подгружаю Exchange cmdlets..."
        try {
            $snap = Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Exchange' }
            if ($snap) {
                foreach ($s in $snap) { Add-PSSnapin $s.Name -ErrorAction Stop }
                if (Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue) { return }
            }
        } catch {}
        try {
            if (-not $script:ExchangeMgmtHost) { $script:ExchangeMgmtHost = 'localhost' }
            $sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $script:ExchangeMgmtHost) -Authentication Kerberos -ErrorAction Stop
            Import-PSSession $sess -DisableNameChecking | Out-Null
        } catch {
            throw "Не удалось загрузить Exchange cmdlets (ни snapin, ни remoting). Запускайте на сервере Exchange или настройте remoting."
        }
    }

    Ensure-ExchangeCmdlets

    # Получаем объекты
    $old = Get-MailContact -Identity $OldContactSmtp -ErrorAction SilentlyContinue
    if (-not $old) { throw "Контакт не найден: $OldContactSmtp" }

    $new = Get-Mailbox -Identity $NewMailboxId -ErrorAction SilentlyContinue
    if (-not $new) { throw "Mailbox не найден: $NewMailboxId" }

    Write-Host ("Ищу группы, где контакт '{0}' разрешён как отправитель..." -f $old.PrimarySmtpAddress)

    # Преобразуем список разрешённых отправителей в DN для сравнения
    $groups = @()
    $all = Get-DistributionGroup -ResultSize Unlimited
    foreach ($dg in $all) {
        $allowed = @()
        if ($dg.AcceptMessagesOnlyFromSendersOrMembers) {
            foreach ($id in $dg.AcceptMessagesOnlyFromSendersOrMembers) {
                try {
                    $r = Get-Recipient -Identity $id -ErrorAction Stop
                    if ($r -and $r.DistinguishedName) { $allowed += $r.DistinguishedName }
                } catch {}
            }
        }
        if ($allowed -and ($allowed -contains $old.DistinguishedName)) { $groups += $dg }
    }

    if (-not $groups) { Write-Host "Групп не найдено."; return }

    Write-Host ("Найдено групп: {0}" -f $groups.Count)
    $done = 0
    foreach ($dg in $groups) {
        $hasNew = $false
        # проверим, не добавлен ли уже новый отправитель
        if ($dg.AcceptMessagesOnlyFromSendersOrMembers) {
            foreach ($id in $dg.AcceptMessagesOnlyFromSendersOrMembers) {
                try {
                    $r = Get-Recipient -Identity $id -ErrorAction Stop
                    if ($r -and $r.DistinguishedName -eq $new.DistinguishedName) { $hasNew = $true; break }
                } catch {}
            }
        }

        $op = @{ Add = $new.Identity }
        if (-not $AddOnly) { $op['Remove'] = $old.Identity }

        if ($PSCmdlet.ShouldProcess($dg.Identity, ("Update AcceptMessagesOnlyFromSendersOrMembers: {0}" -f ($op.GetEnumerator() | ForEach-Object { "{0}={1}" -f $_.Key,$_.Value } -join '; ')))) {
            try {
                if ($AddOnly -and $hasNew) {
                    Write-Host ("[SKIP] $($dg.Name): отправитель уже добавлен")
                } else {
                    Set-DistributionGroup -Identity $dg.Identity -AcceptMessagesOnlyFromSendersOrMembers $op -ErrorAction Stop
                    $done++
                    Write-Host ("[OK] Обновлено: {0}" -f $dg.Name)
                }
            } catch {
                Write-Warning ("Ошибка обновления {0}: {1}" -f $dg.Name, $_.Exception.Message)
            }
        }
    }

    Write-Host ("Готово. Обновлено групп: {0}" -f $done)
}

