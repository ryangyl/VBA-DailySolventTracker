Attribute VB_Name = "Module1"
Sub Macro2()
    Sheets("Daily_report").Activate
'
' Macro2 Macro
'

'   Copy Paste to new sheet
    Columns("A:G").Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Paste
    Range("A16").Select
    Columns("A:A").ColumnWidth = 42.14
    Columns("B:B").ColumnWidth = 8.43
    Columns("C:C").ColumnWidth = 14.71
    Columns("E:E").ColumnWidth = 1
    Columns("E:E").ColumnWidth = 8.86
    Columns("E:E").ColumnWidth = 12.86
    Columns("F:F").ColumnWidth = 15.57
    Range("A1:G1").Select
    Application.CutCopyMode = False
   
    Rows("1:1").RowHeight = 33.75
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
        
'   Insert date
    Range("A1").Select
    Selection.FormulaArray = "=TODAY()"
    Range("A1").Select
    Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
    Selection.Font.Bold = True

'   Insert Actual Qty
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Actual Qty"
   
' Lookup - Prompt once and store the file reference
Dim msg As String
Range("C4").Select
MsgBox ("Please Choose File Corresponding to This Week")
msg = Application.GetOpenFilename(Title:="Choose File For This Week " & Date)

' Ensure msg contains a valid path
If msg <> "False" Then
    ' Reference the chosen file
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]=R[-1]C[-1],"""",INDEX('" & msg & "'!C1:C26,MATCH(RC[-1],'" & msg & "'!C3,0),MATCH(R1C1,'" & msg & "'!R4,0)))"
    
    ' Autofill the formula

    Selection.AutoFill Destination:=Range("C4:C100"), Type:=xlFillDefault
    Range("C4:C100").Select
    Calculate
Else
    MsgBox "No file selected.", vbExclamation
    Exit Sub
End If
    Application.CutCopyMode = False
    dt = Cells(1, 1).Value
    ActiveSheet.Name = Format(dt, "dd mmm")
    
'   Insert discrepancy
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "Discrepancy"
    Range("F4").Select
    Selection.FormulaArray = _
        "=IF(RC[-4]=R[-1]C[-4],"""",SUMIF(C2,RC[-4],C4)-SUMIF(C2,RC[-4],C3))"
    Selection.AutoFill Destination:=Range("F4:F100"), Type:=xlFillDefault
    Range("F4:F100").Select
    
'   Others
    Range("B4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 1).Select
    Selection.ClearContents
    
    Range("E4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 1).Select
    Selection.ClearContents
    
    Range("A3:I3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A4").Select
    ActiveWindow.FreezePanes = True
    
'   Fixed value
    Range("A1").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A4").Select
    
    Calculate
    
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
MsgBox ("Please click on Calculate Discrepancy")
' Button
    Set NewSheet = Sheets(Sheets.Count)
    Dim btn As Object
    Dim btn2 As Object
    Set btn = NewSheet.Buttons.Add(825, 10, 80, 30)
    With btn
        .OnAction = "CalculateDiscrepancyPerDay"
        .Caption = "Calculate Discrepancy"
        .Font.Bold = True '
    End With
  
   
End Sub

Sub CalculateDiscrepancyPerDay()

    Dim lastSheet As Worksheet, delSheet As Worksheet
    Dim lastSheetName As String, delSheetName As String
    Dim currentSheetIndex As Integer
    If Range("J3").Value = "Discrepancy/Day" Then Exit Sub
    currentSheetIndex = Sheets(ActiveSheet.Name).Index
    Set lastSheet = Sheets(currentSheetIndex - 1)
    lastSheetName = lastSheet.Name

    ' Add new columns in columns J and K
    Columns("J:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Set titles for the new columns
    Range("J3").Value = "Discrepancy/Day"
    Range("K3").Value = "Upper Limit"

    ' Calculate Discrepancy/Day (Column J)
    Range("J4:J" & lastSheet.Cells(lastSheet.Rows.Count, "B").End(xlUp).Row).FormulaR1C1 = _
        "=IF(RC[-8]=R[-1]C[-8], """", IF(RC[-4]<>"""", RC[-4]-VLOOKUP(RC[-8], '" & lastSheetName & "'!C2:C6, 5, FALSE)))"

    ' Set format of data in column J to 'General'
    Range("J4:J" & lastSheet.Cells(lastSheet.Rows.Count, "B").End(xlUp).Row).NumberFormat = "General"

    ' Set reference for 'DEL No.' sheet
    delSheetName = "DEL No."
    Set delSheet = Sheets(delSheetName)

    ' Calculate Upper Limit values in column K
    Range("K4:K" & lastSheet.Cells(lastSheet.Rows.Count, "B").End(xlUp).Row).FormulaR1C1 = _
        "=IF(RC[-9]=R[-1]C[-9], """", VLOOKUP(RC[-9], '" & delSheetName & "'!C2:C3, 2, FALSE))"

    ' Set format of data in column K to 'General'
    Range("K4:K" & lastSheet.Cells(lastSheet.Rows.Count, "B").End(xlUp).Row).NumberFormat = "General"

    ' Apply conditional formatting to highlight cells in column K where the value exceeds column J
    With Range("J4:J" & lastSheet.Cells(lastSheet.Rows.Count, "B").End(xlUp).Row)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=RC[1]"
        .FormatConditions(1).Interior.Color = RGB(400, 0, 0) ' Red color
        
    End With
    ActiveSheet.Buttons.Delete
    ImportPowerBi
End Sub

Sub ImportPowerBi()
    Dim wb1 As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb1 = Workbooks.Open("\\siwdsntv002\SG_PSC_SG1_PL_08_Control_WHse\Daily Tank Reading\powerbidata.xlsx")
    Set ws1 = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    If wb1.Sheets.Count > 1 Then
    wb1.Sheets(1).Delete
    End If
    ws1.Copy wb1.Sheets(1)
    wb1.Sheets(1).Name = "Data"
    wb1.Sheets(1).Activate
    ActiveSheet.Buttons.Delete
    wb1.Save
    wb1.Close
End Sub

