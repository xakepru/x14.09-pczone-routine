Doc = ThisComponent 'Кладем в переменную текущий документ
c = Doc.DrawPage.count 'Считаем количество графических примитивов i= c
Do While i >= 1
   'Не забываем, что у прогеров первая цифра — ноль
   drawObject = Doc.DrawPage(i - 1)
   if drawobject.ShapeType =
   "com.sun.star.drawing.LineShape" then
   'В случае, если объект — линия, удаляем его
   Doc.DrawPage.remove(drawObject)
endif
i=i- 1 Loop
