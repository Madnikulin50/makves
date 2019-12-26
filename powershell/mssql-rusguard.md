# Передача событий из базы RusGuard

## Требования для использования

+ Операционная система Windows 7+, Windows 2012+. Рекомендуемая Windows 10x64.1803+, Windows 2019x64
+ Windows PowerShell 5+, Рекомендуется Windows PowerShell 5.1

## Запуск

Пример запуска для сохранения данных в файл

```

powershell.exe -ExecutionPolicy Bypass -Command "./mssql-rusguard.ps1" -outfilename rusguard

```

Пример запуска с передачей данных на сервер MAKVES

```

powershell.exe -ExecutionPolicy Bypass -Command "./mssql-rusguard.ps1" -makves_url "http://192.168.0.77:8000"

```


Параметры:

| Имя         | Назначение                                                                   |
|-------------|------------------------------------------------------------------------------|
| connection  | Строка соединения с базой данных. Например:server=localhost;user id=sa;password=pwd                                    |
| outfilename | Имя файла результатов                                                        |
| makves_url  | URL-адрес сервера Makves. Например: http://192.168.0.77:8000          |
| makves_user | Имя пользователя Makves под которым данные отправляются на сервер     |
| makves_pwd  | Пароль пользователя Makves под которым данные отправляются на сервер  |
| start  | Метка времени для измения файлов в формате "yyyyMMddHHmmss"       |
| startfn | Имя файла для метки времени |
