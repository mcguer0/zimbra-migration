# Zimbra Migration — набор скриптов для перехода на Exchange

Проект содержит сценарии для автоматизации переноса почтовых ящиков и контактов из Zimbra в Microsoft Exchange.

## Назначение

Скрипт `Migration-Mailbox.ps1` автоматизирует миграцию ящиков **Zimbra → Exchange** через `imapsync`:

- создаёт/включает mailbox в Exchange и **ждёт 60 сек** для репликации;
- включает IMAP и выдаёт **FullAccess** администратору;
- переносит почту через `imapsync` (со стороны Zimbra);
- при запуске с `-Dryrun` выполняет сухой прогон: проверяет наличие mailbox и подключение к серверам;
- при использовании `-Force` переименовывает старый ящик в Zimbra (`user_old@domain`);
- при использовании `-Force` создаёт/обновляет `transport` в **PMG**;
- приводит **UPN (Имя для входа)** к суффиксу `example.com` (или `$UpnSuffix` из конфига);
- пишет подробные логи на Windows и на сервере Zimbra.

---

## Состав

- `Migration-Mailbox.ps1` — основной сценарий.
- `config.ps1` — настройки (лежит рядом).
- `users.txt` — список пользователей для пакетного режима (по одному в строке, пустые строки и строки `#комментарий` игнорируются).
- `scripts/Contact-Manager.ps1` — экспорт/импорт контактов.
- `scripts/Find-Contact.ps1` — поиск контакта и его групп в Exchange.

Пример `users.txt`:

```
# список к миграции
ivan.petrov
maria.ivanova@example.com
```

---

## Требования

- Windows PowerShell 5.1 (рекомендовано запускать на сервере Exchange или админ-станции в домене).
- Доступны cmdlets Exchange (локально или через ремоут на `$ExchangeMgmtHost`).
- Модули:  
  - `Posh-SSH` (для SSH на Zimbra/PMG),  
  - `ActiveDirectory` (для смены UPN).
- На сервере Zimbra установлен `imapsync` и доступ по SSH.
- Доступ по SSH к PMG (если используете обновление `transport`).
- Учётка, запускающая скрипт, имеет права:
  - в Exchange (Enable-Mailbox, Set-CASMailbox, Add-MailboxPermission),
  - в AD (изменение `userPrincipalName`),
  - SSH к Zimbra/PMG.

Установка недостающих модулей (при необходимости):

```powershell
Install-Module Posh-SSH -Scope AllUsers -Force
# Модуль ActiveDirectory ставится через RSAT / роли AD DS Tools
```

---

## Настройка `config.ps1`

Откройте и заполните ключевые поля:

```powershell
$Domain                 = "example.com"
$AdminLogin             = "EXCH\imapadmin"        # учётка, получающая FullAccess и IMAP proxy-auth

$ExchangeImapHost       = "mail01.example.com"
$ZimbraImapHost         = "zimbra.example.com"

$AdminImapPasswordPlain = "SuperSecret"

$ZimbraSshHost          = "zimbra.example.com"
$ZimbraSshUser          = "root"
$ZimbraSshPasswordPlain = "RootPass"
$ImapSyncPath           = "/usr/bin/imapsync"

$PMGHost                = "pmg.example.com"
$PMGUser                = "root"
$PMGPasswordPlain       = "RootPass"

$LocalLogDir            = ".\ImapSyncLogs"
$ExchangeMgmtHost       = "localhost"

# Важно для «Имя для входа» (UPN):
$UpnSuffix              = "example.com"
```

> Пароли хранятся в открытом виде. Ограничьте доступ к файлам NTFS-правами. При желании можно позже перевести на хранилище учётных данных.

---

## Как это работает (поток для каждого пользователя)

1. Нормализует адрес: `user` → `user@$Domain`.
2. Ищет контакт с таким адресом в Exchange. При запуске с `-Staged` пользователь добавляется в группы контакта, контакт остаётся, а рабочий ящик ещё не создаётся. При финальном запуске с `-Force` контакт удаляется.
3. Проверяет mailbox в Exchange. При подготовке (`-Staged`) создаётся временный ящик `<alias>_1` с адресом `<alias>_1@$Domain`, он отключается и скрывается из адресных списков. При последующем запуске с `-Activate` этот временный ящик переименовывается в боевой `<alias>@$Domain`. Если mailbox отсутствует, выполняется **Enable-Mailbox** и пауза 60 сек.
4. Приводит UPN к `$UpnSuffix` (например, `mailtest@example.com`).
5. Включает IMAP: `Set-CASMailbox -ImapEnabled $true`.
6. Выдаёт FullAccess администратору `$AdminLogin`.
7. На Zimbra по SSH готовит и запускает **imapsync** с повторами и логированием (stream в файл в `$LocalLogDir`).
8. Если перенос **успешен**:
   - при запуске с `-Force` переименовывает старый ящик в `user_old@domain` (при конфликте добавит таймштамп);
   - при запуске с `-Force` создаёт/обновляет запись `transport` в PMG (`user@domain smtp:[ExchangeHost]:25`);
   - без `-Force` эти шаги пропускаются.
9. Убирает временные файлы на Zimbra, закрывает SSH.
10. Пишет итог в консоль и путь к локальному логу.

---

## Запуск

По умолчанию выполняется перенос почты. Для подготовки ящика без удаления контакта и активации используйте `-Staged`. Для сухого прогона без переноса используйте `-Dryrun`. Чтобы автоматически переименовать Zimbra-аккаунт и настроить transport на PMG, добавьте ключ `-Force`.

### Один пользователь

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov
# или явный адрес
.\Migration-Mailbox.ps1 -User ivan.petrov@example.com
```

### Пакетный режим

```powershell
.\Migration-Mailbox.ps1 -Path .\users.txt
```

### С переименованием и транспортом

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov -Force
# либо
.\Migration-Mailbox.ps1 -Path .\users.txt -Force
```

### Предварительная подготовка (staged)

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov -Staged
# либо пакетно
.\Migration-Mailbox.ps1 -Path .\users.txt -Staged
```

Такой запуск создаёт временный mailbox `<alias>_1` (отключённый и скрытый из адресных списков) и переносит пользователя в группы контакта, оставляя сам контакт. Позже выполните миграцию повторно с `-Activate` (и при необходимости с `-Force`), чтобы удалить контакт и переименовать временный ящик в боевой.

### Сухой прогон

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov -Dryrun
```

Без переноса почты будут выведены сведения о наличии mailbox и доступности серверов.

---

## Где смотреть логи

- Windows (подробный stream-лог `imapsync`): `C:\ImapSyncLogs\imapsync-<user>-<timestamp>.log`
- На Zimbra (временный): `/tmp/imapsync-<user>-<timestamp>.log` — удаляется в конце работы.

---

## Проверка результата

В Exchange:

```powershell
Get-Mailbox -Identity ivan.petrov@example.com
Get-CASMailbox -Identity ivan.petrov@example.com | fl ImapEnabled
Get-MailboxPermission -Identity ivan.petrov@example.com | ? {$_.User -match 'imapadmin'}
```

В AD (UPN):

```powershell
Get-ADUser -LDAPFilter "(sAMAccountName=ivan.petrov)" -Properties userPrincipalName | fl userPrincipalName
```

В PMG (по SSH):

```bash
grep "^ivan.petrov@example.com" /etc/pmg/transport
postmap -q "ivan.petrov@example.com" /etc/pmg/transport
```

В Zimbra (по SSH):

```bash
su - zimbra -c 'zmprov ga ivan.petrov_old@example.com | egrep "mail|zimbraMailAlias"'
```

---

## Типичные проблемы и решения

- **Нет Exchange cmdlets** → запускайте на сервере Exchange или задайте `$ExchangeMgmtHost` и проверьте Kerberos/WinRM.
- **`Posh-SSH`/`ActiveDirectory` не найдены** → установите модуль/RSAT и перезапустите консоль админа.
- **SSH недоступен** → проверьте хост/порт/пароль/файрвол на Zimbra и PMG.
- **`imapsync` не найден** → установите на Zimbra и задайте корректный `$ImapSyncPath`.
- **Нет прав на смену UPN** → дайте учётке право изменять атрибуты пользователей в AD.

---

## Замечания по безопасности

- Держите `config.ps1` в защищённой папке (ограничьте ACL).
- Логи могут содержать адреса и структуру ящиков — не передавайте их третьим лицам.

---

## Быстрый чек-лист перед запуском

- [ ] Заполнен `config.ps1` (особенно `$Domain`, `$UpnSuffix`, `$ExchangeImapHost`, `$ZimbraImapHost`, пароли).
- [ ] На Zimbra есть `imapsync` и доступ по SSH.
- [ ] На PMG есть доступ по SSH (если используете transport).
- [ ] На Windows есть `Posh-SSH` и `ActiveDirectory`.
- [ ] Есть права в Exchange/AD.

---

## Работа с контактами

### Экспорт и импорт
Скрипт `Contact.ps1` может экспортировать контакты из Active Directory в CSV и импортировать их в Exchange. Все CSV сохраняются в папку `lists` в корне репозитория.

```powershell
# Экспорт из AD
./Contact.ps1 -Export contacts.csv

# Импорт в OU из CSV
./Contact.ps1 -Import contacts.csv

# Импорт только одного контакта
./Contact.ps1 -Import contacts.csv -Contact ivan.petrov@example.com
```

### Экспорт групп рассылки из Zimbra
```powershell
./Contact.ps1 -ExportDistributionGroup
```

### Поиск контакта
Для быстрого поиска контакта и его групп используйте `scripts/Find-Contact.ps1`:

```powershell
./scripts/Find-Contact.ps1 -User ivan.petrov@example.com
```

Скрипт выводит адрес контакта и список рассылок, в которых он состоит.
