# FinAuto
Finance Automation 


Sub InsertObjects()

Dim fileName As String
Dim fileExt As String
Dim obj As Object

'Insert first object
fileName = "object1"
fileExt = ".jpg"
Set obj = ActiveSheet.Shapes.AddPicture(fileName & fileExt, msoFalse, msoTrue, 10, 10, 100, 100)

'Insert second object
fileName = "object2"
fileExt = ".docx"
Set obj = ActiveSheet.Shapes.AddObject(fileName & fileExt, "", msoTrue, msoFalse, 10, 120, 100, 100)

'Insert third object
fileName = "object3"
fileExt = ".png"
Set obj = ActiveSheet.Shapes.AddPicture(fileName & fileExt, msoFalse, msoTrue, 10, 230, 100, 100)

End Sub
