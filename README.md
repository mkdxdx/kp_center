# kp_center
1
Парсер на python специально для товара с сайта http://center-bespeki.com/
Использование:
  1. Зайти на сайт, выбрать товар, скопировать ссылку на товар.
  2. Вставить ссылку в окно созданное скриптом. Ссылок может быть несколько, каждая на новой строке.
  3. Опционально можно включить другую номенклатуру вручную в формате "Название,кол-во,цена за ед". После ссылок так же можно вставить запятую чтобы указать количество.
  4. Ввести данные клиента внизу слева.
  5. Нажать "Сформировать документ", должен быть установлен любой просмотрщик XLS файлов.
  
Вся номенклатура взятая по ссылка будет зачисляться в отдельную сумму предоплаты поставщику отдельно от общей суммы.
Если оставить область текста пустой и нажать на кнопку, заполнится данными примера.
  
Писалось под python 3.5. 
