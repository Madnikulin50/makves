## Агент мониторинга действий пользователя

### Функционал 
+ Съемка снимков экрана
+ Определение смены активного окна или его заголовка
+ Журналирование клавиатурного ввода (keylogger)
+ Передача данных в плаьформу MAKVES по протоколу HTTP/HTTPS

### Требования для использования
+ Операционная система Windows 7+, Windows 2012+. Рекомендуемая Windows 10x64.1803+, Windows 2019x64
+ Windows PowerShell 5+, Рекомендуется Windows PowerShell 5.1

### Запуск

Запуск агента с передачей данных по протоколу HTTP
```
powershell.exe -ExecutionPolicy Bypass -Command "./user-agent.ps1" -url "http://10.0.0.10:8000" -user admin -pwd admin
```


Параметры:

| Имя         | Назначение                                      |
|-------------|-------------------------------------------------|
| url | Адрес сервера                           |
| user | Пользователь                           |
| pwd | Пароль                           |

