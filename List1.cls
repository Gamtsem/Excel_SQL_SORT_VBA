VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim Numbers As Variant
Dim FinalRow As Long
Dim Del As Variant
Dim DeleteRow As Long
Dim i As Long
Dim k As Long
Dim g, h As Long
Dim j As Long


Private Sub CommandButton1_Click()
Application.Calculation = xlCalculationManual
    ApplicationScreenUpdating = False
         file = GetFileName(, "C:\", "")
        
            Cells(1, 1).Value = file
   
    
    Application.Calculation = xlCalculationAutomatic
ApplicationScreenUpdating = True

End Sub

Private Sub CommandButton2_Click()
Dim c As Object

Worksheets("Запрос1").Columns("A:IH").ClearContents
Call request1



For j = 1 To 239 Step 4
Worksheets("Запрос1").Cells(2, j + 3).Cut Destination:=Worksheets("Запрос1").Cells(1, j + 2)
    Worksheets("Запрос1").Cells(1, j).Value = "Время"
    Worksheets("Запрос1").Cells(1, j + 1).Value = "Время от начала испытаний, с"
    Next j


    For i = 2 To 240 Step 4

            FinalRow = Worksheets("Запрос1").Cells(Rows.Count, i).End(xlUp).Row
            Numbers = Worksheets("Запрос1").Range(Worksheets("Запрос1").Cells(2, i + 1), Worksheets("Запрос1").Cells(FinalRow, i + 1))
            Worksheets("Запрос1").Range(Worksheets("Запрос1").Cells(2, i + 1), Worksheets("Запрос1").Cells(FinalRow, i + 1)) = Numbers
            DeleteRow = Worksheets("Запрос1").Cells(Rows.Count, i + 1).End(xlUp).Row
            Worksheets("Запрос1").Range(Worksheets("Запрос1").Cells(2, i + 2), Worksheets("Запрос1").Cells(FinalRow, i + 2)) = Empty

For k = 2 To FinalRow
    
    Worksheets("Запрос1").Cells(k, i - 1) = Left(Worksheets("Запрос1").Cells(k, i - 1), 19)
    Worksheets("Запрос1").Cells(k, i - 1) = LTrim(Worksheets("Запрос1").Cells(k, i - 1))
    Worksheets("Запрос1").Cells(k, i - 1).Formula = TimeValue(Worksheets("Запрос1").Cells(k, i - 1)) + DateValue(Worksheets("Запрос1").Cells(k, i - 1))
    Worksheets("Запрос1").Cells(k, i).Formula = (Worksheets("Запрос1").Cells(k, i - 1) - Worksheets("Главное меню").Range("e6")) * 24 * 60 * 60
    Worksheets("Запрос1").Cells(k, i - 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    Worksheets("Запрос1").Cells(k, i).NumberFormat = "0.00"
    Worksheets("Запрос1").Cells(k, i + 1).NumberFormat = "0.00"
       
    Next k

  Next i
  
Worksheets("Запрос1").Columns("A:IH").EntireColumn.AutoFit
End Sub



