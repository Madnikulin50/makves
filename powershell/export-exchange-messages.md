# Экспорт сообщений из Exchange

## Подготовка

+ Exchange Management Shell
+ Microsoft Exchange Web Services Managed API https://www.microsoft.com/en-us/download/details.aspx?id=42951
+ Для работы режима compliance необходимо установить:
  + Microsoft Office 2010 Filter Packs https://www.microsoft.com/en-ie/download/details.aspx?id=17062
  + PDF iFilter 64 11.0.01 https://supportdownloads.adobe.com/detail.jsp?ftpID=5542

## Запуск

Запуск нужно осуществлять из Windows PowerShell с включенным Exchange Management Shell

Пример запуска

```
./export-exchange-messages.ps1 -outfilename export-exchange-messages
```

Параметры:

| Имя         | Назначение                                      |
|-------------|-------------------------------------------------|
| outfilename | Имя файла результатов                           |
| user        | [Обязательный] Имя пользователя под которым производится запрос |
| domain | [Обязательный] Домен пользователя под которым производится запрос |
| pwd         | [Обязательный] Пароль пользователя под которым производится запрос |
| save_body     | Сохранять тело почтового сообщения                    |
| compliance  | Вычисление соответсвие текста стандартам  |
| start  | Метка времени для измения файлов в формате "yyyyMMddHHmmss"       |
| startfn | Имя файла для метки времени |
| makves_url  | URL-адрес сервера Makves. Например: http://192.168.0.77:8000          |
| makves_user | Имя пользователя Makves под которым данные отправляются на сервер     |
| makves_pwd  | Пароль пользователя под которым данные отправляются на сервер  |
