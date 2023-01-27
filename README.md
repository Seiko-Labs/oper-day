# Закрытие операционного дня в системе Colvir (Otbasy Bank)

## ШАГ 1.
Переходим в режим «COPPER».
<br>
В появившемся списке необходимо развернуть список. Необходимо развернуть список, отжать «+» на строке 00 (Рис 1).
<br><br>
![Рис. 1](doc_images/0.jpg)
<br><br>
В столбце «Проц. 4» будут стоять «√». Их снимаем, откликиваем на «Снять признак выполнения регламентной процедуры 4» (рис 2) - действия выполняется только для ГО – «00».
<br><br>
![Рис. 2](doc_images/1.jpg)
<br><br>
Обновляем страницу (рис 3).
<br><br>
![Рис. 3](doc_images/2.jpg)
<br><br>
Кликаем на подразделение «00» и запускаем «Регламентная процедура 4» (рис 4).
<br><br>
![Рис. 4](doc_images/3.jpg)
<br><br>
В окне на «Задать дату вручную» проставляем галочку. Задаем текущую дату. Убираем галочку с «Выполнить для всех подчиненных подразделений» и кликаем на «ОК» (рис 5), подтверждаем действие – кликаем «Да» (рис 6).
<br><br>
![Рис. 5](doc_images/4.jpg)
<br><br>
![Рис. 6](doc_images/5.jpg)
<br><br>
Результат в столбце Обр. по 00 будет стоять галочка (Рис 7).
<br><br>
![Рис. 7](doc_images/6.jpg)
<br><br>
Переходим в «Задания на обработку операционных периодов» (рис 8) кликаем на «обновить» до состояния «обработано» (Рис 9).
<br><br>
![Рис. 6](doc_images/5.jpg)
<br><br>
**Необходимо дождаться состояния «обработано» - строчка исчезнет и в столбце проц.4 появится галочка.**

## ШАГ 2.
Выделяем все филиалы от 01 до 31 и запускаем регламент 2 (Рис 10)
<br><br>
![Рис. 10](doc_images/11.jpg)
<br><br>
Проставляем галочку в «задать дату вручную» и проставляем текущую дату, кликаем «ОК» (Рис 11).
<br><br>
![Рис. 11](doc_images/12.jpg)
<br><br>
Также во вкладке «Все задания на обработку ОП» (Рис 12) дожидаемся состояния «обработано».
<br><br>
![Рис. 12](doc_images/13.jpg)
<br><br>
В колонке Проц.2 по мере отработке преставятся галочки (Рис 13)
<br><br>
![Рис. 13](doc_images/14.jpg)
<br><br>

## ШАГ 3.
Запускаем регламент 2 по ГО «00»
<br>
Проставляем галочку в «задать дату вручную» и проставляем текущую дату, снимаем галочку с «Выполнить для всех подчинённых подразделений» и кликаем «ОК» (Рис 14).
<br><br>
![Рис. 14](doc_images/15.jpg)
<br><br>
Также во вкладке «Все задания на обработку ОП» дожидаемся состояния «обработано».

## ШАГ 4.
Робот преступает к отработке процесса после получения выписки МТ-950
<br><br>
Робот запускает COLVIR. После переходим в режим EXTRCT  (рис 15)
<br><br>
![Рис. 15](doc_images/16.jpg)
<br><br>
Выбираем режим «Выписка», в появившемся окне проставляем сегодняшнюю дату и кликаем «ОК». (Рис 16)
<br><br>
![Рис. 16](doc_images/17.jpg)
<br><br>
В списке выбираем строку с 950 Типом (рис 17), после чего проверяем наличие квитовки (рис 18).  (В выписке в столбце КВ должны отсутствовать галочки)
<br><br>
![Рис. 17](doc_images/18.jpg)
<br><br>
![Рис. 18](doc_images/19.jpg)
<br><br>
Для квитовки счетов необходимо перейти в режим «Автоматическая выверка счетов» и выбираем счет KZ139ХХХ02 (рис 19).
<br><br>
Откликиваем «Операции» и «Подготовить документы для системы выверки» (рис 20).
<br><br>
![Рис. 19](doc_images/20.jpg)
<br><br>
В появившемся окне задаем текущую дату и кликаем на «ОК» (Рис 21).
<br><br>
![Рис. 20](doc_images/21.jpg)
<br><br>
Еще раз откликиваем «Операции» и кликаем на «Сквитовать выбранные счета» (Рис 22). 
<br><br>
![Рис. 21](doc_images/22.jpg)
<br><br>
### **Примечание**
<br><br>
После квитовки выйдет результат операции с указанием кол-во успешных\неуспешных операции (Рис 23). 
<br><br>
![Рис. 22](doc_images/23.jpg)
<br><br>
В случае если неуспешно, необходимо будет запустить «Операции» - «Сквитовать все счета» (Рис 24).
<br><br>
![Рис. 23](doc_images/24.jpg)
<br><br>
