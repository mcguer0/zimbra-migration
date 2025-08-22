# Migration-Mailbox — краткая инструкция

## Назначение

Скрипт автоматизирует миграцию ящиков **Zimbra → Exchange** через `imapsync`:

- создаёт/включает mailbox в Exchange и **ждёт 20 сек** для репликации;
- включает IMAP и выдаёт **FullAccess** администратору;
- переносит почту через `imapsync` (со стороны Zimbra);
- **(по выбору)** переименовывает старый ящик в Zimbra (`user_old@domain`);
- **(по выбору)** создаёт/обновляет `transport` в **PMG**;
- приводит **UPN (Имя для входа)** к суффиксу `mtzd.ru` (или `$UpnSuffix` из конфига);
- пишет подробные логи на Windows и на сервере Zimbra.

---

## Состав

- `Migration-Mailbox.ps1` — основной сценарий.
- `config.ps1` — настройки (лежит рядом).
- `users.txt` — список пользователей для пакетного режима (по одному в строке, пустые строки и строки `#комментарий` игнорируются).

Пример `users.txt`:

```
# список к миграции
ivan.petrov
maria.ivanova@mtzd.ru
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
$Domain                 = "mtzd.ru"
$AdminLogin             = "EXCH\imapadmin"        # учётка, получающая FullAccess и IMAP proxy-auth

$ExchangeImapHost       = "mail01.mtzd.ru"
$ZimbraImapHost         = "zimbra.mtzd.ru"

$AdminImapPasswordPlain = "SuperSecret"

$ZimbraSshHost          = "zimbra.mtzd.ru"
$ZimbraSshUser          = "root"
$ZimbraSshPasswordPlain = "RootPass"
$ImapSyncPath           = "/usr/bin/imapsync"

$PMGHost                = "pmg.mtzd.ru"
$PMGUser                = "root"
$PMGPasswordPlain       = "RootPass"

$LocalLogDir            = ".\ImapSyncLogs"
$ExchangeMgmtHost       = "localhost"

# Важно для «Имя для входа» (UPN):
$UpnSuffix              = "mtzd.ru"
```

> Пароли хранятся в открытом виде. Ограничьте доступ к файлам NTFS-правами. При желании можно позже перевести на хранилище учётных данных.

---

## Как это работает (поток для каждого пользователя)

1. Нормализует адрес: `user` → `user@$Domain`.
2. Проверяет mailbox в Exchange; если нет — **Enable-Mailbox** и пауза 20 сек.
3. Приводит UPN к `$UpnSuffix` (например, `mailtest@mtzd.ru`).
4. Включает IMAP: `Set-CASMailbox -ImapEnabled $true`.
5. Выдаёт FullAccess администратору `$AdminLogin`.
6. На Zimbra по SSH готовит и запускает **imapsync** с повторами и логированием (stream в файл в `$LocalLogDir`).
7. Если перенос **успешен**:
   - **спросит** переименовать старый ящик в `user_old@domain` (или пропустить).  
     При конфликте добавит таймштамп.
   - **спросит** создать/обновить запись `transport` в PMG (`user@domain smtp:[ExchangeHost]:25`).
   - При параметре `-Force` вопросы **не задаются** — оба шага выполняются автоматически.
8. Убирает временные файлы на Zimbra, закрывает SSH.
9. Пишет итог в консоль и путь к локальному логу.

---

## Запуск

### Один пользователь

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov
# или явный адрес
.\Migration-Mailbox.ps1 -User ivan.petrov@mtzd.ru
```

### Пакетный режим

```powershell
.\Migration-Mailbox.ps1 -Path .\users.txt
```

### Без вопросов (как раньше)

```powershell
.\Migration-Mailbox.ps1 -User ivan.petrov -Force
# либо
.\Migration-Mailbox.ps1 -Path .\users.txt -Force
```

---

## Где смотреть логи

- Windows (подробный stream-лог `imapsync`): `C:\ImapSyncLogs\imapsync-<user>-<timestamp>.log`
- На Zimbra (временный): `/tmp/imapsync-<user>-<timestamp>.log` — удаляется в конце работы.

---

## Проверка результата

В Exchange:

```powershell
Get-Mailbox -Identity ivan.petrov@mtzd.ru
Get-CASMailbox -Identity ivan.petrov@mtzd.ru | fl ImapEnabled
Get-MailboxPermission -Identity ivan.petrov@mtzd.ru | ? {$_.User -match 'imapadmin'}
```

В AD (UPN):

```powershell
Get-ADUser -LDAPFilter "(sAMAccountName=ivan.petrov)" -Properties userPrincipalName | fl userPrincipalName
```

В PMG (по SSH):

```bash
grep "^ivan.petrov@mtzd.ru" /etc/pmg/transport
postmap -q "ivan.petrov@mtzd.ru" /etc/pmg/transport
```

В Zimbra (по SSH):

```bash
su - zimbra -c 'zmprov ga ivan.petrov_old@mtzd.ru | egrep "mail|zimbraMailAlias"'
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
