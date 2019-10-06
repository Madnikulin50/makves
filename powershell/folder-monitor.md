## Агент изменений файлов

### Функционал 
+ Мониторинг папки и передача измений  на анализ в платформу MAKVES по протоколу HTTP/HTTPS

### Требования для использования
+ Операционная система Windows 7+, Windows 2012+. Рекомендуемая Windows 10x64.1803+, Windows 2019x64
+ Windows PowerShell 5+, Рекомендуется Windows PowerShell 5.1

### Запуск

Запуск агента с передачей данных по протоколу HTTP
```
powershell.exe -ExecutionPolicy Bypass -Command "./folder-monitor.ps1" -folder c:\work -url "http://10.0.0.10:8000" -user admin -pwd admin
```

Параметры:

| Имя         | Назначение                                      |
|-------------|-------------------------------------------------|
| folder | Имя папки                           |
| url | Адрес сервера                           |
| user | Пользователь                           |
| pwd | Пароль                           |


