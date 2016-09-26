# VBA_functions

Функции для файлов статистики

Список:
* **Replace_symbols(ByVal txt As String)**  - замена символов "~!@/\#$%^:?&*=|`;""" и пробелов на нижнее подчеркивание

* **VBA_Start()** - отключение обновления экрана и пересчетов

* **VBA_Start()**  - возвращение обновлений и пересчетов

* **CreateSh(cr_sh As String)** - создание страницы в случае если такой страницы нет.

* **OpenFile(ByRef patch As String, nm_sh As String) As String** - открытие книги по указанному пути и активация листа. На выходе - название книги

* **openFileCSV(ByRef patch As String)** - открытие CSV

* **getQuartal(month&) As String** - отдает квартал по номеру месяца

* **getMonth_form_00(num_month As Integer) As String** - отдает номер месяца в формате 01

* **getPatchHistTR(brand As String, year As Integer,  ver_year As Integer, thisMonth As Integer, ver_month As Integer) As String** - формирует путь к файлу Total Russia CA:
    * nmBrand - марка формата  LP, KR, RD ...
    * ThisYear - текущий актуальный год
    * VarYear -  переменный год
    * ThisMonth - текуйщий актуальный месяц
    * VarMonth - переменный месяца

* **getLastRow() As Integer** - получение последней строки на активном листе (не работает в текстовых файлах)

* **getLastColumn() As Integer** - получение последнего столбца на активном листе (не работает в текстовых файлах)

* **getClntType(in_data$, i&) as String** - получение типа клинета.
    * in_data - тип клинета из  Total Russia CA
    * i - номер столбца получаемого типа на выходе

<center>

| 1 | 2 | 3 | 4| 
|---|---|---|---|
| салон| salon| salon| single |
| сеть салонов | chain_salons | salon | chain |
| ч/м | hdres | salon | single |
| сеть магазинов | chain_shops | shop | chain |
| магазин | shop | shop | single |
| магазин | shop | shop | single |
| салон-маг. | salon | salon | single |
| (пусто) | other | other | single |
| школа | school | school | single |
| другое | other | other | single |
| нейл-бар | nails_bar | nails | single |
| сеть нейл-баров | chain_nails | nails | chain |
| e-commerce | e-commerce | e-commerce | single |

</center>

* **getMregWhitoutBrand(in_data$) as String** - проверка названия мегарегиона, в случае бренд перед названием мегарегиона, бренд обрезается.

* **getMregExt(in_data_mreg$, in_data_reg$) as String** - разделение мегарегиона Moscow GR на отдельные регионы

* **getMregLat(in_data_mreg As String) As String** - получение английского название мегарегионов из расширенного формата

* **getSalonName(in_sln_nm$, in_sln_addres$, in_city$) as String** - получение полного названия салона + адреса

* **getMonthNumeric(in_data$) as Integer** - получение номера месяца из русского названия

* **getNameMonthEN(in_data%) As String** - получения короткого анг. названия месяца от номера месяца

* **getMonthEng(month$) as String** - получение короткого анг. названия месяца от русского названия

* **getYearType(ThisYear&, in_data&, i&) As Variant**
    * i=1, в случае отсутсвия года отдается 2008
    * i=2, отдается краткий формат года
        * TY
        * PY
        * PPY

* **getMag(in_min_price As Long, in_max_price As Long, in_place As Long, mag_type As String) As Variant** - отдает маг по трем типам бизнеса *(mag_type)*:
    * hair
    * nail
    * skin

    или по одному из параметров мага *(mag_type)*:
    * avg_price
    * place

* **getTypeBusiness(in_brand$) as String** - получение от бренда типа бизнеса:
    * Hair
    * Nails
    * Skin

* **getTypeDN(in_data&) as String** - получение типа DN:
    * Active
    * Closed

* **getRoundNum(in_data as Variant) as Double** - получение округленного значения числа или нуля, в случае если не число

* **getNum2num0&(in_data As Variant) as Double** - получение нуля в случае если пусто или не число

* **getNum2numNull(in_data) As Variant** - получение Empty если не число

* **getNmChainTop(inNmChain$, inCdChain&, inNmTypeClnt$) as String** - получение статуса отсележиваемой в Fr&Ch сети

* **getLTM(in_row&, inThisMonth&, typeFN$) As Variant** - получение данных по LTM *(last twelve months)* :
    * avg_ca - средний оборот за 12 последних месяцев
    * frqOrders - частота срабатывания формате x\12
    * type_avg_ca - тип оборота за LTM c шагами *(2.5, 5, 10, 15, 20, 25, 30, 50, 60, 70, )*
* **getVectoreEV(in_data#) as String** - получение вектора эволюции (+ или -)

* **getMonthlyCA(in_row&, in_month&, in_thisMonth&, in_typeY$, in_typeVal$, in_type_period$) as Double** - получение оборота по месяцам.
    * in_row - номер строки
    * in_month - номер месяца в цикле для суммирования
    * in_thisMonth - номер текущего месяца
    * in_typeY - тип года и формат денег:
        * PY - предущий год
        * TY - текущий год
    * in_typeVal
        * LOR - предущий год в деньгах Партнера
        * PRTN - текущий год в деньгах Партнера
    * in_type_period - тип рассчетного периода:
        * Total - цикл весь год
        * YTD - цикл по отчетному периоду











