# Экспорт пользователей, групп и компьютеров из AD

## Требования для использования

+ Операционная система Windows 7+, Windows 2012+. Рекомендуемая Windows 10x64.1803+, Windows 2019x64
+ Windows PowerShell 5+, Рекомендуется Windows PowerShell 5.1
+ Remote Server Administration Tools for Windows 10 (или другой для соответвующей версии ОС)
+ Права на чтение данных из ActiveDirectory (Read all user information) [Дополнительно](https://social.technet.microsoft.com/Forums/en-US/c8b5886a-f0f1-4e20-b083-d36521d4dec6/delegation-to-read-all-users-properties-in-the-domain?forum=winserverDS)

### Запуск

Пример запуска:

```

powershell.exe -ExecutionPolicy Bypass -Command "./export-ad.ps1" -base DC=acme``,DC=local -server dc.acme.local -outfilename export-ad

```
Параметры:

| Имя         | Назначение                                      |
|-------------|-------------------------------------------------|
| base        | Корневая OU для экспорта                        |
| server      | Имя домен-контроллера                           |
| user             | [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
| pwd              | [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
| outfilename | Имя файла результатов                           |
| makves_url  | URL-адрес сервера Makves. Например: http://192.168.0.77:8000          |
| makves_user | Имя пользователя Makves под которым данные отправляются на сервер     |
| makves_pwd  | Пароль пользователя Ьфлмуы под которым данные отправляются на сервер  |

После запуска, если не задан параметр user, будет выведено окно логина на домен-контроллер, нужно ввести логин-пароль пользователя имеющего право читать учетные данные из домена
