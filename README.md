# auto_equip_distribution


### КОРОТКО

#### Для чего

Файл «auto_equip_distribution_v4.3.exe» (номер версии может меняться) создан для быстрого и безошибочного создания Excel таблицы с информацией для распределения закупаемого оборудования между подразделениями (ЛВУМГ).

 

#### Как работает

Заполняем файлы-шаблоны «Договір.xlsx» и «Групування.xlsx». Запускаем «auto_equip_distribution_v4.3.exe», после чего появляется результат - файл «Рознарядка … .xlsx».

Исходной информацией для заполнения файла «Договір.xlsx» служат данные из Договора про закупку: название позиции и закупаемое количество.

В файл «Групування.xlsx» копируем данные из файла Групування, созданного на этапе составления Технических требований для закупки.

При запуске «auto_equip_distribution_v4.3.exe» скрипт ищет в файле «Групування.xlsx» позиции из файла «Договір.xlsx»  и вычисляет суммарное количество заказанное каждым подразделением (ЛВУМГ).

Результаты сохраняются в виде таблицы в файле «Рознарядка … .xlsx».


 

#### Исключение ошибок

Для исключения ошибок ввода и интерпретации информации в скрипт встроен ряд функций.

Для визуального контроля полученных результатов, ниже таблицы с результатами файла «Рознарядка … .xlsx», для каждой позиции добавляются суммарные значения из трех источников: Рознарядка, Договор, Групування. Значения из всех трех источников должны совпадать. Значения выведены в виде сгенерированных Excel формул (вычисляются средствами Excel), что так же позволяет визуально проконтролировать корректность диапазонов ячеек из которых взяты данные.

Помимо Excel формул, суммарные значения (из Рознарядка, Договор, Групування) вычисляются скриптом, и в случае несовпадения, соответствующие ячейки выделяются цветом.

Так же по каждой позиции сравниваются единицы измерения из Договора и Групування. В случае несовпадения соответствующие ячейки с единицами измерения так же выделяются цветом, что в некоторых случаях позволяет быстро идентифицировать проблему несовпадения суммарных значений.

Исходные файлы «Договір.xlsx» и «Групування.xlsx» проверяются на корректность заполнения, что позволяет исключить обработку заведомо некорректной информации.

 

#### Предупреждения об ошибках

В случае возникновения ошибок (некорректные исходные данные, несовпадение суммарных значений и т.д.) в файле-результате «Рознарядка … .xlsx» генерируются соответствующие сообщения об ошибках и рекомендации по их устранению. Сообщения выделяются цветом.

 

 

 

### ПОДРОБНО 
Предполагается, что интерфейс максимально прост и интуитивно понятен. Информацию приведенную ниже можно рассматривать как справочную.

#### Порядок работы

 

1. Содержимое папки dist (файлы: «auto_equip_distribution_v4.3.exe», файлы-шаблоны «Договір.xlsx» и «Групування.xlsx», и опционально – «ЛВУ - коди-назви.xlsx») копируем к себе в отдельную рабочую папку. Далее работаем в этой отдельной папке.


2. Заполняем файл-шаблон «Договір.xlsx». Информацию берем из официального Договора на поставку. Столбцы  «Позиц. за ДОГОВОРОМ», «Од. виміру» и «Кількість» - обязательны для заполнения. Пустые строки недопустимы.

    _Примечание: скрипт анализирует наличие и формат используемых исходных данных во избежание использования некорректной информации._


3. Копируем ячейки с данными из файла «… Групування… . xlsx», созданного на этапе составления Технических требований для закупки, и вставляем их в файл-шаблон «Групування.xlsx». Столбцы «Од. вим.», «Кільк. в од.вим.», «Код», «Позиц. за ДОГОВОРОМ» - обязательны для заполнения. Пустые строки недопустимы.

    _Примечание: Можно использовать оригинальный файл «… Групування… . xlsx», переименовав его в  «Групування.xlsx», но копирование ячеек с данными в файл-шаблон – предпочтительнее, т.к. позволяет проконтролировать соответствие столбцов шапки-шаблона и скопированных данных, а так же исключает ненужную информацию (скрытые листы, встроенные скрипты, отформатированные, но не заполненные ячейки и т.д.). Как и в случае с файлом «Договір.xlsx», скрипт анализирует таблицу файла «Групування.xlsx» на наличие и формат используемых исходных данных._
 

4. Запускаем файл «auto_equip_distribution_v4.3.exe», после чего в рабочей папке появится файл вида «Рознарядка 2021-07-21_T102200.xlsx», где последняя часть имени файла вида «2021-07-21_T102200» будет соответствовать текущей дате и времени (добавлено для уникальности имени файла во избежание перезаписи результата при повторном запуске «auto_equip_distribution_v4.3.exe»).

    _Примечание: Файл вида «Рознарядка 2021-07-21_T102200.xlsx» можно переименовывать и редактировать по своему усмотрению. Файл не содержит скрытой информации или встроенных скриптов._



#### Файл-результат «Рознарядка … .xlsx» (выглядит как «Рознарядка 2021-07-21_T102200.xlsx»)

Сгенерированный файл-результат «Рознарядка … .xlsx» предназначен для предоставления полной информации о распределении оборудования согласно Договора – исходную информацию и результат расчета.

Файл содержит 4 листа: «Договір», «Групування», «Рознарядка. Перевірка», «Рознарядка по регіонах».

 

##### Лист «Договір»

Содержит копию таблицы из исходного файла «Договір.xlsx».

 

##### Лист  «Групування»

Содержит копию таблицы из исходного файла «Групування.xlsx».

 

##### Лист «Рознарядка. Перевірка»

Служит для проверки результата распределения. Этот лист становится активным сразу после формирования файла «Рознарядка … .xlsx».

Здесь, для наглядности контроля суммарных значений и единиц измерения, распределение выполнено по каждому ЛВУМГ (без распределения по Регионам).

Лист содержит результирующую таблицу распределения по ЛВУМГ, где первая колонка данных – номер по порядку, вторая – название ЛВУМГ сформировавшего потребность, начиная с третьей – количество, по каждой позиции (Каждая колонка представляет данные по соответствующей позиции. Номер (название) позиции и единицы измерения указаны в шапке таблицы) в единицах измерения по Договору.

 

Ниже таблицы для каждой позиции добавлены строки с контрольными суммарными значениями для каждой позиции: «Рознарядка. Сумарна кількість», «Договір. Сумарна кількість», «Групування. Сумарна кількість». Значения вычисляются Excel по сгенерировнным скриптом формулам. Диапазоны ячеек с данными берутся из соответствующих листов этого же файла: лист «Рознарядка. Перевірка», лист «Договір», лист «Групування». Контрольные суммарные значения по каждой позиции должны соответствовать количеству из договора (лист «Договір).

Параллельно контрольные суммарные значения вычисляются скриптом опираясь на те же источники.

 

В случае если контрольное суммарное значение в строке «Рознарядка. Сумарна кількість» не соответствует количеству по Договору (строка «Договір. Сумарна кількість»), ячейка с ошибочным значением закрашивается красным цветом, а в конце строки «Рознарядка. Сумарна кількість» появляется предупреждающее сообщение. В строках «Групування. Сумарна кількість» и/или «Групування. Одиниці виміру» появляется «расшифровка» возможной проблемы: ячейки соответствующей позиции с суммарным значением и/или единицами измерения закрашиваются оренжевым цветом, в конце строки появляется информационное сообщение с рекомендацией проверить/откорректировать данные в исходном ФАЙЛЕ «Групування.xlsx».

Если номер (позиция так же может иметь цифро-буквенное или буквенное обозначение) позиции из таблицы «Договір» не найдены в таблице «Групування», то в строке «Групування. Сумарна кількість» вместо суммарного значения появится надпись «поз. відс.», что означает «позиція відсутня». В этом случае необходимо проверить наличие и/или правильность написания соответствующей позиции в ФАЙЛЕ «Групування.xlsx».

 

Перед началом вычислений скрипт проверяет корректность данных в исходных файлах «Договір.xlsx» и «Групування.xlsx». В случае наличия некорректных данных (в одном из исходных файлов или в каждом из них) на листе «Рознарядка. Перевірка» вместо результирующей таблицы распределения выводится сообщение об ошибке заполнения с именем проблемного файла и рекомендацией его откорректировать.

 

##### Лист «Рознарядка по регіонах»

Содержит результирующую таблицу с теми же колонками и данными, что и результирующая таблица на листе «Рознарядка. Перевірка», но с группировкой подразделений (ЛВУМГ) по Регионам.

Для каждого Региона добавляется строка с суммарными значениями количества по каждой позиции.

 

При несовпадении контрольных суммарных значений ячейки таблицы закрашиваются красным цветом, внизу таблицы выводится предупреждающая надпись с рекомендацией перейти на лист «Рознарядка. Перевірка».

 

В случае наличия некорректных данных в исходных файлах «Договір.xlsx» и/или «Групування.xlsx» лист «Рознарядка по регіонах» не формируется.

 

 

##### Файл «ЛВУ - коди-назви.xlsx»

Информация из файла «ЛВУ - коди-назви.xlsx» устанавливает соответствие между: кодами ЛВУМГ (из файла РПП на 2021), названиями ЛВУМГ (старого и нового), и принадлежностью ЛВУМГ к региону (складу).
Это соответствие так же прописано в скрипте файла «auto_equip_distribution_v4.3.exe» и используется в отсутствие файла «ЛВУ - коди-назви.xlsx».
Начиная с v4.3, при наличии файла «ЛВУ - коди-назви.xlsx» в рабочей директории, информация берется из этого файла.
Таким образом, корректируя файл «ЛВУ - коди-назви.xlsx», можно менять текущую информацию о кодах, названиях и принадлежности подразделений к Регионам. Структура файла должна оставаться без изменений. 


 
