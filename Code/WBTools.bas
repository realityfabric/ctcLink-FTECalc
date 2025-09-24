Attribute VB_Name = "WBTools"
Option Explicit

'@Folder("Library")
'@IgnoreModule ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation, UseMeaningfulName
Public Function FindColumnByName(ByVal ws As Worksheet, ByVal columnName As String, Optional ByVal HeaderRow As Long = 1) As Long
    Dim columnNumber As Variant
    For Each columnNumber In ws.Range("A" & HeaderRow & ":ZZ" & HeaderRow)
        If columnNumber.Value = columnName Then
            FindColumnByName = columnNumber.Column
            Exit For
        Else
            FindColumnByName = -1
        End If
    Next columnNumber
End Function

' Returns column number if column name is found, otherwise returns -1
Public Function FindColumnLikeName(ByVal ws As Worksheet, ByVal columnName As String, Optional ByVal HeaderRow As Long = 1) As Long
    Dim cell As Variant
    For Each cell In ws.Range("A" & HeaderRow & ":ZZ" & HeaderRow)
        If cell.Value Like columnName Then
            FindColumnLikeName = cell.Column
            Exit Function
        Else
            FindColumnLikeName = -1
        End If
    Next cell
End Function

Public Function FindLastRowInSheet(ByVal ws As Worksheet) As Long
    ' based on https://stackoverflow.com/a/11169920
    Dim LastRow As Long

    With ws
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            LastRow = .Cells.Find(What:="*", _
                                  After:=.Range("A1"), _
                                  Lookat:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False).Row
        Else
            LastRow = 1
        End If
    End With

    FindLastRowInSheet = LastRow

End Function

Public Function getCellValueAt(ByVal ws As Worksheet, ByVal col As Long, ByVal Row As Long) As Variant
    Dim rg As Range
    Dim colLetter As String

    colLetter = GetColumnLetterByNumber(col)
    Set rg = ws.Range(colLetter & Row)

    getCellValueAt = rg.Value
End Function

Public Sub setCellValueAt(ByVal ws As Worksheet, ByVal col As Long, ByVal Row As Long, ByVal val As Variant)
    Dim rg As Range
    Dim colLetter As String

    colLetter = GetColumnLetterByNumber(col)
    Set rg = ws.Range(colLetter & Row)

    rg.Value = val
End Sub

Public Function GetColumnLetterByNumber(ByVal columnNumber As Long) As String
    ' Define array of columns
    Dim colArr(1 To 70) As String
    colArr(1) = "A"
    colArr(2) = "B"
    colArr(3) = "C"
    colArr(4) = "D"
    colArr(5) = "E"
    colArr(6) = "F"
    colArr(7) = "G"
    colArr(8) = "H"
    colArr(9) = "I"
    colArr(10) = "J"
    colArr(11) = "K"
    colArr(12) = "L"
    colArr(13) = "M"
    colArr(14) = "N"
    colArr(15) = "O"
    colArr(16) = "P"
    colArr(17) = "Q"
    colArr(18) = "R"
    colArr(19) = "S"
    colArr(20) = "T"
    colArr(21) = "U"
    colArr(22) = "V"
    colArr(23) = "W"
    colArr(24) = "X"
    colArr(25) = "Y"
    colArr(26) = "Z"
    colArr(27) = "AA"
    colArr(28) = "AB"
    colArr(29) = "AC"
    colArr(30) = "AD"
    colArr(31) = "AE"
    colArr(32) = "AF"
    colArr(33) = "AG"
    colArr(34) = "AH"
    colArr(35) = "AI"
    colArr(36) = "AJ"
    colArr(37) = "AK"
    colArr(38) = "AL"
    colArr(39) = "AM"
    colArr(40) = "AN"
    colArr(41) = "AO"
    colArr(42) = "AP"
    colArr(43) = "AQ"
    colArr(44) = "AR"
    colArr(45) = "AS"
    colArr(46) = "AT"
    colArr(47) = "AU"
    colArr(48) = "AV"
    colArr(49) = "AW"
    colArr(50) = "AX"
    colArr(51) = "AY"
    colArr(52) = "AZ"
    colArr(53) = "BA"
    colArr(54) = "BB"
    colArr(55) = "BC"
    colArr(56) = "BD"
    colArr(57) = "BE"
    colArr(58) = "BF"
    colArr(59) = "BG"
    colArr(60) = "BH"
    colArr(61) = "BI"
    colArr(62) = "BJ"
    colArr(63) = "BK"
    colArr(64) = "BL"
    colArr(65) = "BM"
    colArr(66) = "BN"
    colArr(67) = "BO"
    colArr(68) = "BP"
    colArr(69) = "BQ"
    colArr(70) = "BR"

    GetColumnLetterByNumber = colArr(columnNumber)
End Function

Public Function GetSheet(ByVal sheetName As String, Optional ByRef wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim ws As Worksheet
    Dim Sheet As Worksheet

    Set Sheet = Nothing
    For Each ws In wb.Sheets
        If sheetName = ws.Name Then
            Set Sheet = ws
            Set GetSheet = ws
            Exit Function
        End If
    Next ws

    If Sheet Is Nothing Then
        Set Sheet = wb.Sheets.Add(After:=wb.Sheets.Item(wb.Sheets.Count))
        Sheet.Name = sheetName
        Set GetSheet = Sheet
        Exit Function
    End If

    Set GetSheet = Sheet
End Function

Public Function GetSheetLike(ByVal sheetName As String, Optional ByRef wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim ws As Worksheet
    Dim Sheet As Worksheet

    Set Sheet = Nothing
    For Each ws In wb.Sheets
        If LCase$(ws.Name) Like LCase$(sheetName) Then
            Set Sheet = ws
            Set GetSheetLike = ws
            Exit Function
        End If
    Next ws

    Set GetSheetLike = Sheet
End Function


