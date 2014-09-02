REM  *****  BASIC  *****
Sub WriteGreenText()
Dim oDoc As Object, xText As Object,xTextRange As Object
'Кладем в переменную текущий документ
oDoc = ThisComponent
'Кладем в переменную текст документа
xText = oDoc.getText() xTextRange = xText.getEnd() 'Ищем конец текста
'Начинаем задавать форматирование
xTextRange.CharBackColor = green
TextRange.CharHeight = 16.0
xTextRange.CharPosture = com.sun.star.awt.Font
Slant.ITALIC 'курсив
'Запись текста
xTextRange.setString("зеленый курсив")
End Sub

