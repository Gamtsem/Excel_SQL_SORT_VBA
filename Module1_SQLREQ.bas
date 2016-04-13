Attribute VB_Name = "Module1"
Function GetFileName(Optional ByVal Title As String = "Выберите файл для обработки", _
Optional ByVal InitialPath, _
Optional ByVal MyFilter As String = "Все файлы (*.*),") As String
' функция выводит диалоговое окно выбора папки с заголовком Title,
' начиная обзор диска с папки InitialPath
' возвращает полный путь к выбранной папке, или пустую строку в случае отказа от выбора
If Not IsMissing(InitialPath) Then
On Error Resume Next: ChDrive Left(InitialPath, 1)
ChDir InitialPath ' выбираем стартовую папку
End If
res = Application.GetOpenFilename(MyFilter, , Title, "Открыть") ' вывод диалогового окна
GetFileName = IIf(VarType(res) = vbBoolean, "", res) ' пустая строка при отказе от выбора
End Function

'Sub ПримерИспользования_GetFileName()
'ИмяФайла = GetFileName("Заголовок окна", ThisWorkbook.Path) ' запрашиваем имя файла
' ===================== другие варианты вызова функции =====================
' текстовые файлы, стартовая папка не указана
' ИмяФайла = GetFileName("Выберите текстовый файл", , "Текстовые файлы (*.txt),")
' файлы любого типа из папки "C:\Windows"
' ИмяФайла = GetFileName(, "C:\Windows", "")
' ============================================================­==============

'If ИмяФайла = "" Then Exit Sub ' выход, если пользователь отказался от выбора файла
'MsgBox "Выбран файл: " & ИмяФайла, vbInformation
'End Sub

Sub request1()

Dim r As Long


'переменная хранящая результат запроса
 Dim tbl As Recordset

'строка запроса SQL
Dim SQLr As Variant
Dim val As Variant

'переменная хранящая ссылку на подключенную БД
Dim dbs As Database

'подключаемся к mdb
Set dbs = DAO.OpenDatabase(Worksheets("Главное меню").Cells(1, 1).Value)

For r = 6 To 65

 'составляем строку SQL запроса
 val = Worksheets("Главное меню").Cells(r, 2).Value
 SQLr = val
 'отправляем запрос открытой БД
 'результат в виде таблицы сохранен в tbl
 Set tbl = dbs.OpenRecordset(SQLr)
 
 'вставляем результат в лист начиная с ячейки A1
 Worksheets("Запрос1").Cells(2, 4 * r - 23).CopyFromRecordset tbl
 
 Next r
'Закрываем временную таблицу
  tbl.Close
  
'Очищаем память. Если этого не сделать, то таблица
'так и останется висеть в оперативке.
  Set tbl = Nothing
  
'Закрываем базу
dbs.Close
  Set dbs = Nothing
End Sub


