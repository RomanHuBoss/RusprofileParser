# RusprofileParser
Парсер данных юридических лиц и ИП, запрашиваемых с сайта rusprofile.ru по ИНН или ОГРН, с использованием VBScript

Запуск скрипта несложен (под Windows): перейти в консоли в папку со скриптом и выполнить нижеследующую команду.

innRequest.vbs -i=[ИНН] -o=[ОГРН] -s=rusprofile -f=[ПАПКА_С_РЕЗУЛЬТАТОМ]

Описание параметров:
1) вместо [ИНН] следует указать числовое значение ИНН;
2) вместо [ОГРН] следует указать числовое значение ОГРН;
3) вместо [ПАПКА_С_РЕЗУЛЬТАТОМ] следует указать путь к папке, в которую будет сохраняться результат

Первый и второй параметры взаимозаменяемы.
При одновременном указании и ИНН, и ОГРН предпочтение отдается ИНН. 

В принципе, скрипт несложно развить до парсинга других ресурсов в том же ключе.



