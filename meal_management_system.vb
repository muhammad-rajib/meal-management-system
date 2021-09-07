Sub meal_management_system()
    
    'Activate Monthly Sheet
    Sheets("Settings").Activate
    
    'Count total members
    Dim total_persons As Integer
    Dim x As Integer
    
    total_persons = 1
    x = 7
    
    Do While Cells(x, 4).Value <> ""
        total_persons = total_persons + 1
        x = x + 1
    Loop
    
    If total_persons = 1 Then
        MsgBox "No Member included in member list! Please try Again."
        Exit Sub
    End If
    
    'Collecting month name
    Dim month As String
    month = Range("g7").Value
    sheet_name = month & " - " & Year(Now)
    
    'Create a Summary Sheet
    Dim sht As Worksheet
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets.Add

    On Error Resume Next                                    'Prevent Excel from stopping on an error but just goes to next line
    ws.Name = sheet_name

    If Err.Number = 1004 Then
        MsgBox "Worksheet with this name already exists. Delete exist sheet for create new one with this name."
        Application.DisplayAlerts = False                   'Prevent confirmation popup on sheet deletion
        ws.Delete
        Application.DisplayAlerts = True                    'Turn alerts back on
        On Error GoTo 0                                     'Stop excel from skipping errors
        Exit Sub                                            'Terminate sub after a failed attempt to add sheet
    End If

    On Error GoTo 0
    Sheets(sheet_name).Move after:=Sheets("Settings")

    'Activate Monthly Sheet
    Sheets(sheet_name).Activate
    
    Worksheets(sheet_name).Cells.Interior.ColorIndex = 14
    
    'Worksheet Settings
    '========================================
    Worksheets(sheet_name).Cells.EntireColumn.ColumnWidth = 14
    Worksheets(sheet_name).Cells.HorizontalAlignment = xlCenter
    Worksheets(sheet_name).Cells.VerticalAlignment = xlCenter
    Worksheets(sheet_name).Cells.EntireColumn.Font.Name = "Comic Sans MS"
    Worksheets(sheet_name).Cells.EntireColumn.Font.Size = 14
    Rows("2:3").RowHeight = 30
    
    'SET TITLE
    '==================================================
    Range("c2:k2").Cells.MergeCells = True
    Range("c2:k2").Value = "Meal Management System"
    Range("c2:k2").Interior.ColorIndex = 44

    With Range("c2:k2").Font
        .Bold = True
        .Size = 20
    End With

    With Range("c2:k2").Borders
        .Color = vbGrey
        .Weight = 3
    End With

    'SET TOTAL AMOUNT
    '==================================================
    Range("b4:c4").Cells.MergeCells = True
    Range("b5:c5").Cells.MergeCells = True

    Range("b4").Value = "Total Amount(Tk)"
    Range("b4").Interior.ColorIndex = 37
    Range("b5").Interior.ColorIndex = 15

    With Range("b4:c5").Borders
        .Color = vbGrey
        .Weight = 3
    End With

    'SET TOTAL EXPENSE
    '==================================================
    Range("e4:f4").Cells.MergeCells = True
    Range("e5:f5").Cells.MergeCells = True

    Range("e4").Value = "Total Expense(Tk)"
    Range("e4").Interior.ColorIndex = 37
    Range("e5").Interior.ColorIndex = 15

    With Range("e4:f5").Borders
        .Color = vbGrey
        .Weight = 3
    End With

    'SET TOTAL MEALS
    '==================================================
    Range("h4:i4").Cells.MergeCells = True
    Range("h5:i5").Cells.MergeCells = True

    Range("h4").Value = "Total Meals"
    Range("h4").Interior.ColorIndex = 37
    Range("h5").Interior.ColorIndex = 15

    With Range("h4:i5").Borders
        .Color = vbGrey
        .Weight = 3
    End With

    'SET CURRENT MEAL RATE
    '==================================================
    Range("k4:l4").Cells.MergeCells = True
    Range("k5:l5").Cells.MergeCells = True

    Range("k4").Value = "Current Meal Rate"
    Range("k4").Interior.ColorIndex = 37
    Range("k5").Interior.ColorIndex = 15

    With Range("k4:l5").Borders
        .Color = vbGrey
        .Weight = 3
    End With
    
    'Formatting of Rows 4 & 5
    Rows(4).EntireRow.RowHeight = 25
    Rows(5).EntireRow.RowHeight = 30
    Rows(4).EntireRow.Font.Size = 16
    Rows(5).EntireRow.Font.Size = 18
    
    'Start Individual Summary Report
    '=========================================================================================
    Sheets(sheet_name).Activate

    'Individual row
    Range(Cells(9, 2), Cells(9, total_persons + 1)).Cells.MergeCells = True
    Range("b9").Value = "Individual Summary"
    Range("b9").Interior.ColorIndex = 45

    With Range("b9").Font
        .Bold = True
        .Size = 18
    End With

    With Range(Cells(9, 2), Cells(9, total_persons + 1)).Borders
        .Color = vbGrey
        .Weight = 3
    End With
    '==================================================

    'Total, Uses & Balance rows
    With Range(Cells(10, 2), Cells(12, total_persons + 1)).Borders
        .Color = vbGrey
        .Weight = 3
    End With

    Range("b10:b12").Interior.ColorIndex = 38
    Range("b10").Value = "Total"
    Range("b11").Value = "Uses"
    Range("b12").Value = "Balance"

    With Range("b10:b12").Font
        .Bold = True
        .Size = 14
    End With

    Range(Cells(10, 3), Cells(12, total_persons + 1)).Interior.ColorIndex = 35
    '===================================================
    
    'Date - Members - Bazar rows
    Dim i As Integer
    Dim j As Integer

    j = 7
    For i = 2 To total_persons + 1
        If i = 2 Then
            Cells(13, i).Value = "Date"
            Cells(13, i).Interior.ColorIndex = 37
            Cells(13, i).Borders.Color = vbGrey
            Cells(13, i).Borders.Weight = 3
        Else
            Cells(13, i).Value = Sheets("Settings").Cells(j, 4).Value
            Cells(13, i).Interior.ColorIndex = 36
            Cells(13, i).Borders.Color = vbGrey
            Cells(13, i).Borders.Weight = 3
            j = j + 1
        End If
    Next

    Cells(13, total_persons + 2).Value = "Bazar(Tk)"
    Cells(13, total_persons + 3).Value = "Person"

    Range(Cells(13, total_persons + 2), Cells(13, total_persons + 3)).Interior.ColorIndex = 37
    With Range(Cells(13, total_persons + 2), Cells(13, total_persons + 3)).Borders
        .Color = vbGrey
        .Weight = 3
    End With
    '===================================================
    
    'finding month number
    Dim month_num As Integer
    Dim day_total_persons As Integer
    Dim date_str As String
    Dim year_str As String

    year_str = Year(Now)
    date_row = 14

    If month = "January" Then
        month_num = "01"
    ElseIf month = "February" Then month_num = "02"
    ElseIf month = "March" Then month_num = "03"
    ElseIf month = "April" Then month_num = "04"
    ElseIf month = "May" Then month_num = 5
    ElseIf month = "June" Then month_num = 6
    ElseIf month = "July" Then month_num = 7
    ElseIf month = "August" Then month_num = 8
    ElseIf month = "September" Then month_num = 9
    ElseIf month = "October" Then month_num = 10
    ElseIf month = "November" Then month_num = 11
    Else
        month_num = 12
    End If

    'Set date in Individual Summary report
    For day_total_persons = 1 To 31
        date_str = day_total_persons & "/" & month_num & "/" & year_str
        Cells(date_row, 2).Value = date_str
        Cells(date_row, 2).Interior.ColorIndex = 36
        Cells(date_row, 2).Borders.Color = vbGrey
        Cells(date_row, 2).Borders.Weight = 3
        date_row = date_row + 1
    Next
    Range("b:b").EntireColumn.AutoFit

    'Making borders for members & bazars cells
    With Range(Cells(14, 3), Cells(44, total_persons + 3)).Borders
        .Color = vbGrey
        .Weight = 2
    End With

    Range(Cells(14, 3), Cells(44, total_persons + 1)).Interior.ColorIndex = 34
    Range(Cells(14, 3), Cells(44, total_persons + 3)).Font.Size = 14
    Range(Cells(14, total_persons + 2), Cells(44, total_persons + 3)).Interior.ColorIndex = 19
    'End Individual Summary Report
    '=========================================================================================


    'Start Individual Meal Report
    '=========================================================================================
    Range(Cells(9, total_persons + 5), Cells(9, total_persons + 4 + total_persons)).Cells.MergeCells = True

    Cells(9, total_persons + 5).Value = "Meals"
    Cells(9, total_persons + 5).Interior.ColorIndex = 45

    With Cells(9, total_persons + 5).Font
        .Bold = True
        .Size = 18
    End With

    With Range(Cells(9, total_persons + 5), Cells(9, total_persons + 4 + total_persons)).Borders
        .Color = vbGrey
        .Weight = 3
    End With

    With Range(Cells(10, total_persons + 5), Cells(9, total_persons + 4 + total_persons)).Borders
        .Color = vbGrey
        .Weight = 3
    End With

    Cells(10, total_persons + 5).Value = "Total Meal"
    Cells(10, total_persons + 5).Interior.ColorIndex = 42

    With Cells(10, total_persons + 5).Font
        .Bold = True
        .Size = 15
    End With

    'borders of members cells in meal report
    Range(Cells(10, total_persons + 6), Cells(10, total_persons + 4 + total_persons)).Interior.ColorIndex = 15

    'Copy Members names row with date column
    Range(Cells(13, 2), Cells(13, total_persons + 1)).Copy
    Cells(11, total_persons + 5).PasteSpecial
    'Cells(11, total_persons + 5).Font.Bold = True
    Application.CutCopyMode = False

    'Copy Dates column
    Range(Cells(14, 2), Cells(44, 2)).Copy
    Cells(12, total_persons + 5).PasteSpecial
    Application.CutCopyMode = False
    Columns(total_persons + 5).EntireColumn.AutoFit

    'Making borders of members cells in Meals report
    With Range(Cells(12, total_persons + 6), Cells(42, total_persons + 4 + total_persons)).Borders
        .Color = vbGrey
        .Weight = 2
    End With

    Range(Cells(12, total_persons + 6), Cells(42, total_persons + 4 + total_persons)).Interior.ColorIndex = 15
    'End Individual Meal Report
    '=========================================================================================
    
        'Adding Formulas
    '=========================================
    Dim total_amount As String
    Dim total_expense As String
    Dim total_meals As String
    Dim meal_rate As String
    Dim person_total_amount As String
    Dim person_total_uses As String
    Dim balance As String
    Dim person_total_meal As String
    Dim column As Integer
    Dim person_total_meal_column As Integer
    
    total_persons = total_persons - 1
    
    Range("c11", Cells(11, total_persons + 2)).NumberFormat = "0.00"
    Range("c12", Cells(12, total_persons + 2)).NumberFormat = "0.00"
    
    'Excel column array
    Dim column_name(35) As String
    column_name(1) = "a"
    column_name(2) = "b"
    column_name(3) = "c"
    column_name(4) = "d"
    column_name(5) = "e"
    column_name(6) = "f"
    column_name(7) = "g"
    column_name(8) = "h"
    column_name(9) = "i"
    column_name(10) = "j"
    column_name(11) = "k"
    column_name(12) = "l"
    column_name(13) = "m"
    column_name(14) = "n"
    column_name(15) = "o"
    column_name(16) = "p"
    column_name(17) = "q"
    column_name(18) = "r"
    column_name(19) = "s"
    column_name(20) = "t"
    column_name(21) = "u"
    column_name(22) = "v"
    column_name(23) = "w"
    column_name(24) = "x"
    column_name(25) = "y"
    column_name(26) = "z"
    column_name(27) = "aa"
    column_name(28) = "ab"
    column_name(29) = "ac"
    column_name(30) = "ad"
    column_name(31) = "ae"
    column_name(32) = "af"
    column_name(33) = "ag"
        
    'Formula for total amount
    total_amount = "=sum(c10" & ":" & column_name(total_persons + 2) & 10 & ")"
    Range("b5").Formula = total_amount

    'Formula for total expense
    total_expense = "=sum(" & column_name(total_persons + 3) & "14" & ":" & column_name(total_persons + 3) & "44" & ")"
    Range("e5").Formula = total_expense

    'Formula for total meals
    total_meals = "=sum(" & column_name(total_persons + 7) & 10 & ":" & column_name(total_persons + 15) & 10 & ")"
    Range("h5").Formula = total_meals

    'Formula for current meal rate
    meal_rate = "=e5/h5"
    Range("k5").Formula = meal_rate
    Range("k5").NumberFormat = "0.00"

    'Find total amount for each person
    column = 3
    For x = 1 To total_persons
       person_total_amount = "=SUM(" & column_name(column) & "14" & ":" & column_name(column) & "44" & ")"
       Range(column_name(column) & 10).Formula = person_total_amount
       column = column + 1
    Next
     
    'Find total uses of each person
    person_total_meal_column = total_persons + 7
    column = 3
     
    For x = 1 To total_persons
       person_total_uses = "=" & column_name(person_total_meal_column) & 10 & "*" & "k5"
       Range(column_name(column) & 11).Formula = person_total_uses
       column = column + 1
       person_total_meal_column = person_total_meal_column + 1
    Next
     
    'Find balance of each person
    column = 3
    For x = 1 To total_persons
       balance = "=" & column_name(column) & 10 & "-" & column_name(column) & 11
       Range(column_name(column) & 12).Formula = balance
       column = column + 1
    Next
    
    'Find total meals for each person
    column = total_persons + 7
    For x = 1 To total_persons
       person_total_meal = "=SUM(" & column_name(column) & "12" & ":" & column_name(column) & "42" & ")"
       Range(column_name(column) & 10).Formula = person_total_meal
       column = column + 1
    Next
   'End Formaula Setup
   '==================================================================================

    
End Sub

