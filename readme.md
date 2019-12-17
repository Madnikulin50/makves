## Контейнер MAKVES

Система реагирования на инциденты информационной безопасности и аудита информацинных ресурсов

### Системные требования

#### Hardware
  + Процессор i5 4 core (Рекомендуется i5 не ниже 6 поколения с 8 ядрами)
  + Оперативная память 8GB (Рекомендуется 16 GB)
  + Жесткий диск 128 GB (Рекомендуется 256 GB)
  
#### Windows
  + i5 4 ядра (Рекомендуется i5 не ниже 6 поколения с 8 ядрами), 16 GB RAM, 128 GB HDD (Рекомендуется 256 GB) 
  
  
  + _Вариант 1_ Физическая машина
      + Windows 10/2016/2019
      + Docker Desktop 18.* -для Windows 10/2016, Windows Server 2019 входит в поставку (Может быть установлено во время пилота. Нужен административный доступ)
      
  + _Вариант 2_ Виртуальная машина или физическая машина
      + Windows 10 x64, Windows Server 2016/2019
      + PostgreSQL https://www.postgresql.org/download/windows/
      + Нужен административный доступ для установки системы
  
#### Linux
  + Ubunty 16.*, 18.*
  + Docker 18.*
  + docker-compose 3.6+

#### MacOS X
  + MacOS 10.14.*
  + Docker 18.*
  
### Необходимо к прибытию инженера
  + Рабочая станция в соответствии с указанными выше требованиями
  + Доступ в интернет на рабочей станции
  + Собранная скриптами информация об объектах компьютерной инфраструктуры (Желательно).
    + Информация о пользователя и компьютерах в домене(используется скрипт export-ad.ps1)
    + Информация о папках общего доступа в домене(используется скрипт explore-folder.ps1)
    + Информация о событиях (используется скрипт export-events.ps1 или копии файлов EventLog (.evtx) c контроллера домена или машина на которую осужествляется EventLog Forwarding)

### Docker
#### Установка

Создать рабочий каталог
Сохранить в нем файл docker-compose.yml
В рабочем каталоге выполнить комманды:

Загрузка последней версии MAKVES
````
 docker-compose pull 
````

#### Запуск
   
````
 docker-compose up
````

После старта системы Web-консоль будет доступна по адресу http://localhost:8000. По-умолчанию для входа следует использовать пользователя admin и пароль admin.

Рекомендуем загрузить настроки по-умолчанию из этого репозитория
 
### MSI

#### Установка

1. Установить программу из установщика [makves.msi](https://github.com/Madnikulin50/makves/releases "Переход по ссылке")
2. Создать в PostgreSQL пустую базу данных
3. В конфигурационном файле изменить настройки подключения к базе данных PostgresSQL
4. Зарегистрировать makves как сервис:
                                      

#### Запуск

Запустить службу makves

### Загрузка объектов компьютерной инфраструктуры
 
 Для загрузки данных используются файлы созданные скриптами:
  
  + [export-ad.ps1](https://github.com/Madnikulin50/makves-mvp/blob/master/export-ad.ps1 "Переход по ссылке")  - Экспорт данных о пользователях, группах и компьютерах из AD.  [Документация](https://github.com/Madnikulin50/makves-mvp/blob/master/export-ad.md "Описание").
  + [explore-folder.ps1](https://github.com/Madnikulin50/makves-mvp/blob/master/explore-folder.ps1 "Переход по ссылке") - Экспорт данных о папках и файлах. [Документация](https://github.com/Madnikulin50/makves-mvp/blob/master/explore-folder.md "Описание").
  + [export-events.ps1](https://github.com/Madnikulin50/makves-mvp/blob/master/export-events.ps1 "Переход по ссылке") - Экспорт событий из журналов Windows. [Документация](https://github.com/Madnikulin50/makves-mvp/blob/master/export-events.md "Описание")
  + [export-mb-folders.ps1](https://github.com/Madnikulin50/makves-mvp/blob/master/export-mb-folders.ps1 "Переход по ссылке") - Экспорт информации о почтовых ящиках MS Exchange. [Документация](https://github.com/Madnikulin50/makves-mvp/blob/masterexport-mb-folders.md "Описание")
    
 Созданные этими скриптами файлы следует загружать в форму __Данные/Загрузка__
  
 Кроме того, данные о событиях могут быть загружены напрямую из evtx-файлов
  
 
 

