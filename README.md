# OpticGOST 1.2
OpticGOST - набор инструментов (а именно - макросы для САПР Zemax и надстройка для MS Excel), автоматизирующих процесс оформления КД на оптические приборы.

Надстройка OpticGOST для Excel заполняет таблицы оптического выпуска и др. конструкторских документов на основе данных из Zemax. Для этого все необходимые конструктивные параметры и значения аберраций сохраняются из Zemax в файл JSON, который открывается в Excel через интерфейс OpticGOST. 

Есть возможность строить таблицы (конструктивных параметров, параметров опт. деталей, хода лучей) и по стандартным текстовым отчётам Zemax: Prescription data и Raytrace. Для автоматического экспорта всех необходимых отчётов и графиков имеется макрос analysis_export.zpl.

## Функционал 
#### [В версии 1.3] Автозаполнение таблиц оптического выпуска по данным из файла lensdata.json

1. С помощью макроса JSONconfig создайте файл настроек экспорта %lens_name%_config.txt. В этом файле указываются координаты лучей, для которых вычисляются аберрации. По умолчанию за меридиональную принимается плоскость (Px=0;Hx=0) и задан набор координат Hy=[0; 1], Py=[0; 0,5; 0,7; 1]. При необходимости отредактируйте файл.
1. Создайте файл JSON макросом JSONexport.
1. Загрузите файл в Excel через меню "lensdata.json".
#### Создание таблицы конструктивных параметров по файлу Prescription Data
![prescription_data_import](./screenshots/prescription_import.png?raw=true)
1. Сохраните отчёт Prescription Data из Zemax в текстовый файл или воспользуйтесь макросом analysis_export. 
1. Откройте файл в Excel из меню Prescription Data. 
#### Заполнение таблицы хода лучей по файлам Zemax Raytrace
![raytrace_import](./screenshots/raytrace_import.png?raw=true)

4 файла Raytrace для апертурного, главного, верхнего и нижнего лучей автоматически экспортируются с нужными настройками макросом analysis_export.zpl. 

#### Создание таблицы оптических деталей по файлу Prescription Data.

#### Автоматический экспорт всех необходимых в КД отчётов Zemax и графиков аберраций
Графики сохраняются в папке рядом с файлом zmx в виде картинок bmp и в протабулированном текстовом виде.

## Установка
### Установка надстройки для Excel

1. В параметрах Excel разрешите выполнение надстроек без цифровой подписи:

![excel_security_settings](./screenshots/security.png?raw=true)
	
2. Перейдите в меню активации надстроек:

        Параметры -> Надстройки -> Управление -> "Надстройки Excel"

3. В окне "Надстройки" нажмите "Обзор" и перейдите в папку OpticGOST/OpticGOST_for_Excel. Выберите OpticGOSTv1.2.xlam

### Установка макроса для Zemax

Запустите "Установка макроса для Zemax.bat" 

При этом скрипт install_set_paths.ps1 создаст файл в папке OpticGOST файл analysis_export.zpl и скопирует его в папку /Documents/Zemax/Macros.
В меню Macros в Zemax должен появиться макрос ANALYSIS_EXPORT
		
	NB! Если вы переместили папку OpticGOST, макрос надо будет переустановить, так как сгенерированный скриптом файл analysis_export zpl содержит ссылки на папку OpticGOST/config.
