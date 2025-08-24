# === Конфигурация миграции Zimbra -> Exchange ===
# Сохрани этот файл рядом с Migration-Mailbox.ps1

# Домен почты
$Domain                 = ""

# OU для экспорта и импорта контактов
$ContactsSourceOU       = ""
$ContactsTargetOU       = ""

# Админ (используется и как IMAP proxy-auth, и для FullAccess)
$AdminLogin             = ""

# IMAP: приемник — Exchange
$ExchangeImapHost       = ""
$ExchangeImapPort       = 993
$ExchangeImapSSL        = $true

# IMAP: источник — Zimbra
$ZimbraImapHost         = ""
$ZimbraImapPort         = 993
$ZimbraImapSSL          = $true

# Пароль IMAP-админа
$AdminImapPasswordPlain = ""

# SSH к Zimbra (где установлен imapsync)
$ZimbraSshHost          = ""
$ZimbraSshUser          = ""
$ZimbraSshPasswordPlain = ""
$ImapSyncPath           = "/usr/bin/imapsync"

# PMG (Proxmox Mail Gateway) — обновление transport
$PMGHost                = ""
$PMGUser                = ""
$PMGPasswordPlain       = ""

# Логи на Windows
$LocalLogDir            = Join-Path (Split-Path $PSScriptRoot -Parent) 'ImapSyncLogs'

# Хост для подключения к Exchange Management PowerShell (если локально — оставь localhost)
$ExchangeMgmtHost       = "localhost"
