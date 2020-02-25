1. ExcelCreateWinForm -создание книги Excel из WinForms с заполнением строка\столбец использу€  Microsoft.Office.Interop.Excel

2. ExcelInOpenXML  -создание книги Excel из  онсольного приложени€ с заполнением строка\столбец,примен€€ стили форматировани€,
	использу€  OpenXML, ќтличие этой библиотеки OpenXML от Microsoft.Office.Interop.Excel в быстродействии которое на пор€док выше.
	(ƒл€ работы с Excel документами необходимо установить расширение DocumentFormat.OpenXML из Nuget, оно позволит создавать Excel 
	документы дл€ версии Microsoft Office не ниже 2010.)

3. ExcelReadData - пример на WinForms доступа и выборки данных из таблиц Excel,например все данные из столбца, с выбором файла дл€ 
	считывани€ данных  использу€  Microsoft.Office.Interop.Excel

