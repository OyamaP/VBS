Option Explicit
' =====================================
' Config
' =====================================
' use workbook
Const workSheetName = "work"
Const exportFileName = "workbook.xlsx"

' Main Sheet
Const leftSheetName = "left"
Const leftUniqueColumnName = "name"
Const leftColumnName = "other-id"

' Sub Sheet
Const joinSheetName = "join"
Const joinUniqueColumnName = "remarks"
Const joinColumnName = "id"
 ' joinsheet
Dim searchColumnNames : searchColumnNames = Array("other-name","remarks")

' close book (true or false)
Const isEndClose = false

' =====================================
' Run
' =====================================
' Create Excel App
Dim xlApp : Set xlApp = WScript.CreateObject("Excel.Application")
xlApp.Visible = true
' Create Workbook
Dim workBook : Set workBook = xlApp.Workbooks.add()
' Set join sheet & left sheet
Dim leftSheet : Set leftSheet = workBook.Sheets(1)
Dim joinSheet : Set joinSheet = workBook.Sheets.Add()
leftSheet.name = leftSheetName
joinSheet.name = joinSheetName

' Get Current Directory Files
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim curDir : curDir = objShell.CurrentDirectory
Dim objFileSys : Set objFileSys = CreateObject("Scripting.FileSystemObject")
Dim objFolder : Set objFolder = objFileSys.GetFolder(curDir + "\source")
Dim objFile
' Define key:value
Dim dic : Set dic = WScript.CreateObject("Scripting.Dictionary")
Dim dicKey
dic.Add leftUniqueColumnName, leftSheetName
dic.Add joinUniqueColumnName, joinSheetName
' Common
Dim i, row, column

' Copy Cells
For Each objFile In objFolder.Files
    ' only xlsx
    If(objFileSys.GetExtensionName(objFile.Path) = "xlsx") Then
        Dim sourceBook : Set sourceBook = xlApp.Workbooks.Open(objFile.Path)
        Dim sourceSheet : Set sourceSheet = sourceBook.Sheets(1)
        Dim sourceSheetName : sourceSheetName = SwitchSheetFromColumn(sourceSheet, dic)
        WScript.Echo "sheet name is " & sourceSheetName

        For row = 1 To GetMaxLastRow(sourceSheet, 1)
            For column = 1 To GetLastColumn(sourceSheet, 1)
                workBook.Sheets(sourceSheetName).Cells(row, column) = sourceSheet.Cells(row, column)
            Next
        Next
        sourceBook.Close
    End If
Next

' Copy leftSheet as worksheet
leftSheet.Copy ,workBook.WorkSheets(workBook.WorkSheets.Count)
Dim workSheet : Set workSheet = workBook.ActiveSheet
workSheet.name = workSheetName

' left join
Dim lastColumn : lastColumn = GetLastColumn(workSheet, 1)
WScript.Echo "Start left join"
For i = 0  To UBound(searchColumnNames)
    workSheet.Cells(1, lastColumn + i + 1) = searchColumnNames(i)
    For row = 2 To GetMaxLastRow(workSheet, 1)
        workSheet.Cells(row, lastColumn + i + 1) = IndexMatch(joinSheet, joinColumnName, workSheet.Cells(row, FindColumnNumber(workSheet, leftColumnName)), searchColumnNames(i))
    Next
Next
WScript.Echo "End left join"


' Save
workBook.SaveAs(curDir + "/" + exportFileName)
if(isEndClose) Then
    xlApp.Quit
End If


' =====================================
' Functions
' =====================================

' /**************************
' * @param sheet as WorkSheet
' * @param index as string
' * @param word as string
' * @param match as string
' * @return string
' ***************************
Function IndexMatch(sheet, index, word, match)
    On Error Resume Next
    WScript.Echo "sheet = " & sheet.name & ", index = " & index & ", word = " & word & ", match = " & match
    Dim column, alphabet, row, matchColumn, result
    column = FindColumnNumber(sheet, index)
    alphabet = ColumnNumberToAlphabet(sheet, column)
    row = FindRowNumber(sheet, alphabet, word)
    matchColumn = FindColumnNumber(sheet, match)
    result = sheet.Cells(row, matchColumn)
    If Err.Number <> 0 Then
        WScript.Echo "Error:IndexMatch()"
    End If
    IndexMatch = result
End Function

' /**************************
' * @param sheet as WorkSheet
' * @param column as string
' * @param name as string
' * @return int
' ***************************
Function FindRowNumber(sheet, column, name)
    On Error Resume Next
    Dim number : number = sheet.Range(column & ":" & column).Find(name).Row
    If Err.Number <> 0 Then
        WScript.Echo "Error:FindRowNumber()"
    End If
    FindRowNumber = number
End Function

' /**************************
' * @param sheet as WorkSheet
' * @param name as string
' * @return int
' ***************************
Function FindColumnNumber(sheet, name)
    On Error Resume Next
    Dim number : number = sheet.Range("1:1").Find(name).Column
    If Err.Number <> 0 Then
        WScript.Echo "Error: FindColumnNumber()"
    End If
    FindColumnNumber = number
End Function

' /**************************
' * @param sheet as WorkSheet
' * @param number as int
' * @return string
' ***************************
Function ColumnNumberToAlphabet(sheet, number)
    Dim buf
    buf = sheet.Cells(1, number).Address(True, False)
    ColumnNumberToAlphabet = Left(buf, InStr(buf, "$") - 1)
End Function

' /**************************
' * @param sheet as WorkSheet
' * @param column as int
' * @return int
' ***************************
Const xlUp = -4162
Function GetLastRow(sheet, column)
    GetLastRow = sheet.Cells(sheet.Rows.Count, column).End(xlUp).Row
End Function

' /**************************
' * @param sheet as WorkSheet
' * @param row as int
' * @return int
' ***************************
Const xlToLeft = -4159
Function GetLastColumn(sheet, row)
    GetLastColumn = sheet.Cells(row, sheet.Columns.Count).End(xlToLeft).Column
End Function

' /**************************
' * @param sheet as WorkSheet
' * @param row as int
' * @return int
' ***************************
Function GetMaxLastRow(sheet, row)
    Dim maxRow, column, lastRow
    maxRow = 0
    For column = 1 To GetLastColumn(sheet, row)
        lastRow = GetLastRow(sheet, column)
        If maxRow < lastRow Then
            maxRow = lastRow
        End If
    Next
    GetMaxLastRow = maxRow
End Function

' /**************************
' * @param sheet as WorkSheet
' * @param dic as Dictionary
' * @return string
' ***************************
Function SwitchSheetFromColumn(sheet, dic)
    Dim dicKey, isNotFound
    For Each dicKey In dic.Keys
        isNotFound = isEmpty(FindColumnNumber(sheet, dicKey))
        WScript.Echo dicKey & " => " & dic(dicKey) & ", isNotFound => " & isNotFound
        If isNotFound Then
        Else
            SwitchSheetFromColumn = dic(dicKey)
        End If
    Next
End Function
