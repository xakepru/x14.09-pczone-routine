Sub highlightcell
Dim oTextTables As Variant 'Массив — все листы книги
Dim oTextTable As Variant 'Переменная для листа, с которым будем работать
Dim oCell As Variant 'Переменная для ячейки
oTextTables = ThisComponent.getTextTables() 'Кладем в переменную все листы книги
oTextTable = oTextTables.getByIndex(1) 'Кладем в переменную второй лист книги
oCell = oTextTable.getCellByPosition(1, 1) 'Первая ячейка второго листа
'Можно было сделать так
oCell = oTextTable.getCellByName("B3")
oCell.setPropertyValue("BackColor", RGB(0, 0, 0)) 'Красим!
End Sub
