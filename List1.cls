VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����1"
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

Worksheets("������1").Columns("A:IH").ClearContents
Call request1



For j = 1 To 239 Step 4
Worksheets("������1").Cells(2, j + 3).Cut Destination:=Worksheets("������1").Cells(1, j + 2)
    Worksheets("������1").Cells(1, j).Value = "�����"
    Worksheets("������1").Cells(1, j + 1).Value = "����� �� ������ ���������, �"
    Next j


    For i = 2 To 240 Step 4

            FinalRow = Worksheets("������1").Cells(Rows.Count, i).End(xlUp).Row
            Numbers = Worksheets("������1").Range(Worksheets("������1").Cells(2, i + 1), Worksheets("������1").Cells(FinalRow, i + 1))
            Worksheets("������1").Range(Worksheets("������1").Cells(2, i + 1), Worksheets("������1").Cells(FinalRow, i + 1)) = Numbers
            DeleteRow = Worksheets("������1").Cells(Rows.Count, i + 1).End(xlUp).Row
            Worksheets("������1").Range(Worksheets("������1").Cells(2, i + 2), Worksheets("������1").Cells(FinalRow, i + 2)) = Empty

For k = 2 To FinalRow
    
    Worksheets("������1").Cells(k, i - 1) = Left(Worksheets("������1").Cells(k, i - 1), 19)
    Worksheets("������1").Cells(k, i - 1) = LTrim(Worksheets("������1").Cells(k, i - 1))
    Worksheets("������1").Cells(k, i - 1).Formula = TimeValue(Worksheets("������1").Cells(k, i - 1)) + DateValue(Worksheets("������1").Cells(k, i - 1))
    Worksheets("������1").Cells(k, i).Formula = (Worksheets("������1").Cells(k, i - 1) - Worksheets("������� ����").Range("e6")) * 24 * 60 * 60
    Worksheets("������1").Cells(k, i - 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    Worksheets("������1").Cells(k, i).NumberFormat = "0.00"
    Worksheets("������1").Cells(k, i + 1).NumberFormat = "0.00"
       
    Next k

  Next i
  
Worksheets("������1").Columns("A:IH").EntireColumn.AutoFit
End Sub



