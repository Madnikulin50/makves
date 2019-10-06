# Скрипты для наполнения базы объетов Makves IRP

## Подготовка к работе

1. Убедиться что установлен Windows Management Framework 5.0 или 5.1
2. Убедиться что выключены все галочки в Remote Server Administration Tools (RSAT). См. https://blogs.msdn.microsoft.com/adpowershell/2009/03/24/active-directory-powershell-installation-using-rsat-on-windows-7/


## Организация репозитория
У каждого скрипта есть свой файл описания с тем же именем и расширением md

## Список скриптов

| Имя                       | Назначение                                                              |
|---------------------------|-------------------------------------------------------------------------|
| export-ad                 | Экспорт информации о пользователях, группах и компьютерах  из AD        |
| export-events             | Экспорт информации о событиях из Windows EventLog                 |
| export-mb-folders         | Экспорт почтовых ящиков из Exchange                               |
| explore-folder            | Сбор информации о файлах                                                |
| folder-evtx-monitor            | Агент мониторинга за Windows Event Log                                                |
| folder-monitor            | Агент мониторинга за папкой файловой системы                                                |
| user-agent            | Агент мониторинга действий пользователя                                                |


[DEV] -Находятся в процессе разработки
