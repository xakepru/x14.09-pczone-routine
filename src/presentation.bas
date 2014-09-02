'Создаем чистый слайд в текущем документе
Sub CreateSomeSlide
oDoc = ThisComponent
oDrawPages = oDoc.getDrawPages()
oDrawPage = oDrawPages.insertNewByIndex(oDrawPages.getCount())
'Создаем новое текстовое поле
oTextShape = oDoc.createInstance("com.sun.star.drawing.TextShape")
'Задаем форму полю для текста
aPoint = CreateUnoStruct("com.sun.star.awt.Point")
aSize = CreateUnoStruct("com.sun.star.awt.Size")
'Задаем полю координаты
aPoint.X = 1000
aPoint.Y = 1000
'Задаем ширину и высоту
aSize.Width = 15000
aSize.Height = 5000
'Применяем вышеназванные параметры
oTextShape.setPosition(aPoint)
oTextShape.setSize(aSize)
'Вставляем текстовое поле на слайд
oDrawPage.add(oTextShape)
'Пишем текст
oTextShape.setString("А вот и наш текст")
End Sub
