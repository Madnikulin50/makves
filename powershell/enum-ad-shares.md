## Сбор данных о сетевых папках в домене

### Подготовка

- Перед запуском должен быть установлен Remote Server Administration Tools for Windows 10 (или другой для соответвующей версии ОС)

### Запуск

Пример запуска

```

powershell.exe -ExecutionPolicy Bypass -Command "./export-ad-shares.ps1" -base DC=acme``,DC=local -server dc.acme.local -outfilename export-shares

```
Параметры:

| Имя         | Назначение                                      |
|-------------|-------------------------------------------------|
| base        | Корневая OU для экспорта                        |
| server      | Имя домен-контроллера                           |
| outfilename | Имя файла результатов                           |

