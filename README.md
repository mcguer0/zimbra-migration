# Zimbra → Exchange Migration (PowerShell)

Набор сценариев для поэтапной миграции почтовых ящиков и контактов из Zimbra в Microsoft Exchange/Active Directory. Включает автоматическое управление членством в рассылках и «разрешёнными отправителями» (Delivery Management), а также интеграцию с Proxmox Mail Gateway.

## Возможности

- Перенос почты через `imapsync` с потоковым логом и повторами.
- Staged/Activate‑подход:
  - `-Staged`: создаёт временный mailbox `<alias>_1`, добавляет его в группы контакта и в Delivery Management групп, где контакт был разрешён; контакт остаётся.
  - `-Force`: финализирует перенос — делает swap (контакт → mailbox) в Delivery Management, удаляет контакт, переименовывает `<alias>_1 → <alias>`, удаляет старый адрес `_1`, активирует ящик в GAL; дополнительно (часть `-Force`) переименовывает Zimbra‑ящик в `<alias>_old@…` и обновляет PMG transport.
- Синхронизация членства: переносит членство рассылок от контакта к mailbox (есть fallback через Add‑ADGroupMember для проблемных групп).
- Обслуживание Exchange: включает IMAP, выдаёт FullAccess администратору, приводит UPN к заданному суффиксу.
- Утилиты: экспорт/импорт контактов, поиск контакта и его групп, ручная замена отправителя в Delivery Management.

## Структура

- `Migration-Mailbox.ps1` — точка входа для миграции одного/нескольких пользователей.
- `scripts/config.example.ps1` — пример конфига; скопируйте в `scripts/config.ps1` и заполните.
- `scripts/Move-ZimbraMailbox.ps1` — логика подготовки/активации и запуска `imapsync`.
- `scripts/Replace-AcceptedSender.ps1` — замена контакта на mailbox во всех группах, где контакт разрешён отправителем.
- `scripts/Update-PMGTransport.ps1` — обновление `transport` в Proxmox Mail Gateway.
- `scripts/Rename-ZimbraMailbox.ps1` — переименование Zimbra‑ящика в `<alias>_old@…`.
- `scripts/Find-Contact.ps1` — поиск контакта и его групп в Exchange/AD.
- `scripts/Contact-Manager.ps1` — экспорт/импорт контактов.
- `lists/` — рабочие CSV для контактов и рассылок.

Примечание: файл `Contact.ps1` в корне помечен как устаревший и сохранён для совместимости; используйте `scripts/Contact-Manager.ps1`.

## Требования

- Windows PowerShell 5.1.
- Командлеты Exchange локально либо через ремоут на `$ExchangeMgmtHost`.
- Модули: `Posh-SSH`, `ActiveDirectory`.
- На Zimbra установлен `imapsync` и доступ по SSH; доступ по SSH к PMG (если используете transport).
- Учетная запись запуска имеет права в Exchange/AD и SSH.

Установка модулей:

```powershell
Install-Module Posh-SSH -Scope AllUsers -Force
# Модуль ActiveDirectory — через RSAT / AD DS Tools
```

## Настройка

1) Скопируйте и отредактируйте конфиг:

```powershell
Copy-Item .\scripts\config.example.ps1 .\scripts\config.ps1
# затем заполните scripts/config.ps1
```

Ключевые параметры `scripts/config.ps1`:

```powershell
$Domain               = 'example.com'
$ContactsSourceOU     = ''
$ContactsTargetOU     = ''
$DistributionGroupsOU = ''
$AdminLogin           = 'EXCH\\migration'
$ExchangeImapHost     = 'mail01.example.com'
$ZimbraImapHost       = 'zimbra.example.com'
$AdminImapPasswordPlain = 'secret'
$ZimbraSshHost        = 'zimbra.example.com'
$ZimbraSshUser        = 'root'
$ZimbraSshPasswordPlain = 'rootpass'
$ImapSyncPath         = '/usr/bin/imapsync'
$PMGHost              = 'pmg.example.com'
$PMGUser              = 'root'
$PMGPasswordPlain     = 'rootpass'
$LocalLogDir          = .\ImapSyncLogs
$ExchangeMgmtHost     = 'localhost'
$UpnSuffix            = 'example.com'
```

Пароли лежат открыто — ограничьте доступ к репозиторию ACL.

## Сценарии запуска

- Сухой прогон (без переноса):

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov -Dryrun
```

- Подготовка (Staged): создаёт `<alias>_1`, добавляет в группы и Delivery Management, переносит почту, не удаляя контакт:

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov -Staged
```

- Финализация (Force): swap в Delivery Management, удаление контакта, переименование `<alias>_1 → <alias>`, удаление адреса `_1`, активация, дельта‑перенос, PMG и переименование Zimbra:

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov -Force
```

- Пакетно:

```powershell
.\Migration-Mailbox.ps1 -Path .\users.txt -Staged
.\Migration-Mailbox.ps1 -Path .\users.txt -Force
```

`users.txt` — `user` или `user@example.com` построчно; строки `#` игнорируются.

## Что делает миграция (по порядку)

1. Нормализует пользователя в SMTP и ищет контакт.
2. Сохраняет: группы, где контакт — член; и группы, где контакт разрешён отправителем (Delivery Management).
3. `-Staged`:
   - включает временный mailbox `<alias>_1` и скрывает его из адресных списков;
   - добавляет `<alias>_1` в членство групп контакта и в Delivery Management (без удаления контакта);
   - приводит UPN к `$UpnSuffix`, включает IMAP, выдаёт FullAccess; запускает `imapsync`.
4. `-Force`:
   - меняет отправителя в Delivery Management (контакт → mailbox) с fallback‑логикой, удаляет контакт;
   - гарантирует членство mailbox во всех группах контакта (fallback через Add‑ADGroupMember);
   - переименовывает `<alias>_1 → <alias>` с ожиданием репликации и удаляет прокси‑адрес `_1`;
   - активирует ящик в GAL, приводит UPN, включает IMAP, выдаёт FullAccess; запускает дельта‑перенос; обновляет PMG и переименовывает Zimbra‑ящик.

## Утилиты

- Замена отправителя в Delivery Management:

```powershell
# Подготовка: добавить mailbox, не удаляя контакт
scripts/Replace-AcceptedSender.ps1
Replace-AcceptedSender -OldContactSmtp user@domain -NewMailboxId user@domain -AddOnly

# Swap: заменить контакт на mailbox
Replace-AcceptedSender -OldContactSmtp user@domain -NewMailboxId user@domain
```

- Поиск контакта и его групп: `scripts/Find-Contact.ps1 -User user`
- Экспорт/импорт контактов: `scripts/Contact-Manager.ps1` (CSV лежат в `lists/`).

## Логи и проверка

- Локальный лог: `ImapSyncLogs\imapsync-<user>-<timestamp>.log`
- На Zimbra: `/tmp/imapsync-<user>-<timestamp>.log` (удаляется в конце).

Проверка:

```powershell
Get-Mailbox ivan.petrov | fl PrimarySmtpAddress,Alias,EmailAddresses
Get-CASMailbox ivan.petrov | fl ImapEnabled
Get-MailboxPermission ivan.petrov | ? { $_.User -match 'migration' }
```

## Рекомендации по Delivery Management

- Держите список «кто может писать» в базовой группе (например, `all_users`) и добавляйте её в DM нужных рассылок.
- Динамические группы не принимают статическое членство — предупреждения в логах в этом случае ожидаемы.

## Ограничения

- Репликация AD/Exchange может занимать минуты; сценарий делает повторы, но отдельные свойства применяются не мгновенно.
- Если контакт удалён до `-Force` без предварительного экспорта, восстановить его включения в DM невозможно штатно.

## Вклад

Скрипты предоставляются «как есть». PR/issue приветствуются.

