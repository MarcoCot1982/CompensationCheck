'+-----------------------------------------------------------------------------------------------+
'| Author: Marco Cot         DAS:A669714                                                         |
'| Program which allows to check compensations before month closing.                             |
'| version: 1.0 [20230612]                                                                       |
'+-----------------------------------------------------------------------------------------------+
Sub DateConversion()

Application.ScreenUpdating = False

'Preparation of new data
Sheets("Input").Select
Range("Table223[FECHA]").ClearContents
Range("Table223[Landing]").ClearContents
Range("Table223[Compensado]").Copy
Range("Table223[Landing]").PasteSpecial xlPasteValues

'Cleans wrongly formatted data (innecessary words)
RemoveLettersFromRange

'Removes accidental initial and final spaces
RemoveSpaces

'Replaces wrong date dividers
  Range("Table223[LANDING]").Replace ".", "/"
  Range("Table223[LANDING]").Replace "-", "/"
  Range("Table223[LANDING]").Replace ",", "/"
  Range("Table223[LANDING]").Replace "\", "/"
  
'Changes all data to dates
Range("Table223[LANDING]").Select
    Selection.TextToColumns Destination:=Range("BL2"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
    Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
    Other:=False, FieldInfo:=Array(1, xlDMYFormat)
    
Range("Table223[LANDING]").Copy
Range("Table223[FECHA]").PasteSpecial xlPasteValuesAndNumberFormats
    
Range("BI2").Select
Application.CutCopyMode = False

CreateTable

RemoveEmptyRows

End Sub
'+-----------------------------------------------------------------------------------------------+
Sub RemoveLettersFromRange()
    Dim rng As Range
    Dim cell As Range
    Dim cleanValue As String
    Dim i As Long
    
    'Set the range you want to remove letters from
    Set rng = Range("Table223[LANDING]")
    
    'Loop through each cell in the range
    For Each cell In rng
        cleanValue = ""
        
        'Loop through each character in the cell's value
        For i = 1 To Len(cell.Value)
            'Check if the character is not alphabetic
            If Not IsLetter(Mid(cell.Value, i, 1)) Then
                cleanValue = cleanValue & Mid(cell.Value, i, 1)
            End If
        Next i
        
        'Replace the cell value with the cleaned value
        cell.Value = cleanValue
    Next cell
End Sub
'+-----------------------------------------------------------------------------------------------+
Function IsLetter(character As String) As Boolean
    IsLetter = (character Like "[A-Za-záãéíóúÁÉÍÓÚ()!,:_]")
End Function
'+-----------------------------------------------------------------------------------------------+
Sub CreateTable()

Application.ScreenUpdating = False

'Cleans and sorts data to copy and paste in the email message
Sheets("TABLA EMAIL").Select
Range("TablaMail").Select

Range("TablaMail").ClearContents

Sheets("Input").Select
Range("Table223[Pers.No.]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Pers.No.]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Employee/Appl.Name]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Employee/Appl.Name]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Description]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Description]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Short text]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Short text]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Date]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Date]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Number]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Number]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Estado]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Estado]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Manager]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Manager]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[Main service]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[Main service]").PasteSpecial xlPasteValuesAndNumberFormats

Sheets("Input").Select
Range("Table223[COMENTARIO BOS]").Copy
Sheets("TABLA EMAIL").Select
Range("TablaMail[COMENTARIO BOS]").PasteSpecial xlPasteValuesAndNumberFormats

Range("TablaMail").Select

End Sub
'+-----------------------------------------------------------------------------------------------+
Sub RemoveSpaces()

    Dim rng As Range
    Dim cell As Range
    
    Set rng = Range("Table223[FECHA]")
    
    'Loop through each cell in the range
    For Each cell In rng
        'Trim the value and remove any leading or trailing spaces
        Dim trimmedValue As String
        trimmedValue = Trim(cell.Value)
        
        'Check if the trimmed value is a valid date
        If IsDate(trimmedValue) Then
            'If valid, update the cell value with the parsed date
            cell.Value = CDate(trimmedValue)
        Else
            'If not valid, leave the cell value unchanged or handle the error as needed
        End If
    Next cell
End Sub
'+-----------------------------------------------------------------------------------------------+
Sub RemoveEmptyRows()
    
    Dim RowCounter As Double
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim sortColumn As Range
    
    
    Sheets("TABLA EMAIL").Select
    'Set the worksheet
    Set ws = ThisWorkbook.Worksheets("TABLA EMAIL")
    
    'Set the table
    Set tbl = ws.ListObjects("TablaMail")
    
    'Set the sort column range
    Set sortColumn = tbl.ListColumns("COMENTARIO BOS").Range
    
    'Sort the table
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortColumn, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With
    
    RowCounter = 1048576 - WorksheetFunction.CountIf(Range("J:J"), "")

     RowCounter = 1048577 - WorksheetFunction.CountIf(Range("J:J"), "")
     Range(RowCounter & ":5000").Delete
        Sheets("TABLA EMAIL").Select
    'Set the worksheet
    Set ws = ThisWorkbook.Worksheets("TABLA EMAIL")
    
    'Set the table
    Set tbl = ws.ListObjects("TablaMail")
    
    'Set the sort column range
    Set sortColumn = tbl.ListColumns("Manager").Range
    
    'Sort the table
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortColumn, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

     RowCounter = 1048577 - WorksheetFunction.CountIf(Range("J:J"), "")
     Range(RowCounter & ":5000").Delete
        Sheets("TABLA EMAIL").Select
    'Set the worksheet
    Set ws = ThisWorkbook.Worksheets("TABLA EMAIL")
    
    'Set the table
    Set tbl = ws.ListObjects("TablaMail")
    
    'Set the sort column range
    Set sortColumn = tbl.ListColumns("Main service").Range
    
    'Sort the table
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortColumn, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    Range("TablaMail").Select
    Range("D:D").NumberFormat = "text"
End Sub
