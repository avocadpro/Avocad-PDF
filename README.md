# ПРАВА И ВЛАДЕЛЕЦ 
Автор даёт согласие на использование скрипта в любых целях  
Больше интересного для САПР Вы найдете на нашей страничке [avocad.pro](https://avocad.pro/)

# ОПИСАНИЕ
Скрипт выполняет экспорта файлов КОМПАС в PDF тремя способами:
- Активного чертежа/спецификации/фрагмента в КОМПАС
- Запущенных чертежей/спецификаций/фрагментов в КОМПАС
- Множества чертежей/спецификаций/фрагментов из выбранной папки
 
# ТРЕБОВАНИЯ
Необходимо наличие:
- [КОМПАС-Макро](https://help.ascon.ru/KOMPAS/21/ru-RU/61_osobennosti_ustanovki_prikladnoj_biblioteki.html)
- КОМПАС V18 и выше (не тестировалось на более ранних)

# КАК ЗАПУСТИТЬ
1. Скачать файл "AvocadPDF.pym"
2. Запустить КОМПАС
3. Находясь на [Стартовой странице](https://help.ascon.ru/KOMPAS/23/ru-RU/idr_mainframe_full.html) перейти в меню "Макросы"
   > "Приложение"-"КОМПАС-Макро"-"Макросы"
5. Нажать кнопку "Добавить" и выбрать файл "AvocadPDF.pym"
7. Скрипт готов к запуску.

> [!IMPORTANT]
> Если путь приложения КОМПАС отличается от стандартного, то перед запуском скрипта необходимо изменить путь в переменной 'kompas_dir_path'
