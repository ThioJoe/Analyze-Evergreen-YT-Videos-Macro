Sub Analyze_Evergreen_Videos()
'
' Purpose: Determines the current monthly performance of videos of different age ranges.
' 
' 
' Usage: Takes in a YouTube Analytics exported csv report. 
'        > From YouTube Studio Analytics go to 'Advanced Mode'
'        > Select a time range of 28 Days or the previous month. 
'        > Ensure the 'Views' metric/column is enabled (should be by default)
'        > At the top left click the Download icon
'        > Select 'Comma Separated Values' - Extract the zip file
'        > Open the "Table Data" file and run the macro
'
    Dim lastRow As Long

    ' Part 1: Determine Videos in Desired Ranges
    
    ' Take in lower and upper limits via user input box
    Dim lowerViewLimit As Integer
    Dim higherViewLimit As Integer
    lowerViewLimit = InputBox("Enter lower monthly view threshold")
    higherViewLimit = InputBox("Enter higher monthly view threshold")
    '   Test Examples
    'lowerViewLimit = 1000
    'higherViewLimit = 3000
    
    ' -- Dynamically delete uneeded columns that don't have certain words in title
    ' -- Keeps video title, publish date, and number of views
    columnsToKeep = Array("title", "publish", "views")
    Range_Delete_Unwanted_Fields (columnsToKeep)
    
    ' -- Manually removing columns
    'Range("A:A,E:I").Select
    'Selection.Delete Shift:=xlToLeft
    
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("E1").Select
    ' Calcualtes Date Difference
    ActiveCell.FormulaR1C1 = "=DATEDIF(RC[-3], TODAY(),""m"")"
    
    ' Adding law row test
    Range("E1").Select
    lastRow = Range("C" & Rows.Count).End(xlUp).Row
    Selection.AutoFill Destination:=Range("E1:E" & lastRow), Type:=xlFillDefault
    
    Columns("A:E").Select
    Range("E1").Activate
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "E1:E503"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
        
    With ActiveSheet.Sort
        .SetRange Range("A1:E503")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Columns("E:E").Select
    Application.FindFormat.NumberFormat = "0"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]<6, RC[-1]>36), ""YESDELETE"", ""NOKEEP"")"
    Range("F1").Select
    
    ' Added
    lastRow = Range("C" & Rows.Count).End(xlUp).Row
    Selection.AutoFill Destination:=Range("F1:F" & lastRow), Type:=xlFillDefault
    
    Range("A1:F503").AutoFilter Field:=6, Criteria1:="YESDELETE"
    Application.DisplayAlerts = False
    Range("A1:F503").SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    
    On Error Resume Next
    Worksheet.ShowAllData
    
    
    ' Part Two, Now process data in Range
    
    lastRow = Range("C" & Rows.Count).End(xlUp).Row

    Range("F1").Select
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Range("F1").Select
    ActiveCell.FormulaR1C1 = _
        "=CHOOSE(((RC[-1]>4)+(RC[-1]>12)+(RC[-1]>24)),""AGE_YOUNG"",""AGE_MID"",""AGE_OLD"")"
    Range("F1").Select
    Selection.AutoFill Destination:=Range("F1:F" & lastRow)
    Range("G1").Select
    
    ' Below values is actually one huge long string, so need to break and add in variables
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]>=" & higherViewLimit & ", CONCAT(RC[-1], ""_HIGHER""), IF(RC[-4]>=" & lowerViewLimit & ",CONCAT(RC[-1],""_LOWER""), IF(RC[-4]<" & lowerViewLimit & ", """", ""Error"")))"
    Range("G1").Select
    Selection.AutoFill Destination:=Range("G1:G" & lastRow)
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "Age"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "# Videos"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = CStr(lowerViewLimit) & "/month"
    Columns("N:N").Select
    Columns("M:M").ColumnWidth = 11.71
    Range("N3").Select
    ActiveCell.FormulaR1C1 = CStr(higherViewLimit) & "/month"
    Range("N4").Select
    Columns("N:N").ColumnWidth = 12.57
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "6-12 Mo"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "13-24 Mo"
    Range("K6").Select
    ActiveCell.FormulaR1C1 = "25-36 Mo"
    Range("K7").Select
    Columns("K:K").ColumnWidth = 9.57
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],""AGE_YOUNG"")"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],""AGE_MID"")"
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],""AGE_OLD"")"
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],""AGE_YOUNG_HIGHER"")"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],""AGE_MID_HIGHER"")"
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],""AGE_OLD_HIGHER"")"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF(C[-6],""AGE_YOUNG_HIGHER"")+COUNTIF(C[-6],""AGE_YOUNG_LOWER"")"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF(C[-6],""AGE_MID_HIGHER"")+COUNTIF(C[-6],""AGE_MID_LOWER"")"
    Range("M6").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIF(C[-6],""AGE_OLD_HIGHER"")+COUNTIF(C[-6],""AGE_OLD_LOWER"")"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = CStr(lowerViewLimit) & "/month %"
    Range("P4").Select
    Columns("P:P").ColumnWidth = 14.14
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = CStr(higherViewLimit) & "/month %"
    Range("Q4").Select
    Columns("Q:Q").ColumnWidth = 15.43
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]/RC[-4]"
    Range("P4").Select
    Selection.AutoFill Destination:=Range("P4:P6"), Type:=xlFillDefault
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]/RC[-5]"
    Range("Q4").Select
    Selection.AutoFill Destination:=Range("Q4:Q6"), Type:=xlFillDefault
    Range("P4:Q6").Select
    Selection.NumberFormat = "0%"
End Sub


Sub Range_Delete_Unwanted_Fields(aList)
    'Dim aList As Variant
    'aList = Array("publish", "view", "title")
    Dim ws As Worksheet
    Dim rSrc As Range, rTrg As Range, rCll As Range
    Dim vItem As Variant, sAdrs As String

    Set ws = ActiveSheet

    Rem Set Source Range (Header)
    With ws
        Set rSrc = .Cells(1).Resize(1, .Cells(1, .Columns.Count).End(xlToLeft).Column)
        rSrc.EntireColumn.Hidden = False
    End With

    Rem Set Target Range (Fields in Array List)
    For Each vItem In aList
        With rSrc

            Rem Clear 1st Found Cell Address
            sAdrs = vbNullString

            Rem Set 1st Found Cell
            Set rCll = .Cells.Find( _
                What:=vItem, After:=.Cells(.Cells.Count), _
                LookIn:=xlFormulas, LookAt:=xlPart, _
                SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)

            Rem Validate 1st Found Cell
            If Not (rCll Is Nothing) Then

                Rem Get 1st Found Cell Address
                sAdrs = rCll.Address

                Rem Add Found Cell To Target Range
                If rTrg Is Nothing Then
                    Set rTrg = rCll
                Else
                    Set rTrg = Union(rTrg, rCll)
                End If

                Rem Find Other Cells
                Do
                    Set rCll = .Cells.FindNext(After:=rCll)

                    Rem Validate Next Cell against 1st Cell
                    If rCll.Address = sAdrs Then Exit Do

                    Rem Add Next Cell To Target Range
                    Set rTrg = Union(rTrg, rCll)

                Loop Until rCll.Address = sAdrs

    End If: End With: Next

    Rem Validate Target Range
    If Not rTrg Is Nothing Then
        Rem Delete Columns Not in Target Range Only if Headers were found!
        rTrg.EntireColumn.Hidden = True
        rSrc.SpecialCells(xlCellTypeVisible).EntireColumn.Delete
        rTrg.EntireColumn.Hidden = False
    End If

    Application.Goto ws.Cells(1), 1

End Sub
