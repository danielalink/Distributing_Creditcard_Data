Attribute VB_Name = "Module5"
Option Explicit
Sub clean_data()

Dim i, j As Integer
Dim Title As String
Dim RowCnt As Long

    ''''''''''''''''''''''''''''''''''''''''''''''''
    '               Select columns                 '
    ''''''''''''''''''''''''''''''''''''''''''''''''
    For j = 1 To 10
        For i = 1 To 12
            Title = Cells(1, i).Value
            If Title = "Account Number " Or Title = " Merchant Zip " Or Title = " Reference Number " Or Title = " Debit/Credit Flag " Or Title = " SICMCC Code " Then
                Columns(i).Delete
            End If
        Next i
    Next j
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    '             Sort by card number              '
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Range("A:G").Sort Key1:=Range("G1"), Order1:=xlAscending, Header:=xlYes
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    '               Last 4-digits                  '
    ''''''''''''''''''''''''''''''''''''''''''''''''
    RowCnt = ActiveSheet.UsedRange.Rows.Count
    For i = 2 To RowCnt
        Cells(i, 7).Value = Right(Cells(i, 7).Value, 4)
    Next i
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    '              Save as .xlsx file              '
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ActiveSheet.Name = "report"
    ChDir "C:\Users\ckkim\Downloads"
    ActiveWorkbook.SaveAs Filename:="C:\Users\ckkim\Downloads\report.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Sub move_Each_Group_To_Sheets()

    Dim colName As String                                   'ColName of the column that would be filtered
    Dim rngAll As Range
    Dim colsCnt As Integer
    Dim varTemp
    Dim strName As String                                   'Name of Sheets
    Dim s As Long                                           'Number of Sheets
    
    Application.ScreenUpdating = False                      'Pause screen update
    
    colName = "G"                                           'ColName of the column that would be filtered
       
    Set rngAll = ActiveSheet.UsedRange
    If rngAll.Rows.Count < 2 Then                           'When there is only header
        MsgBox "There is no Data.", 64, "Data Error"        'Show message
        Exit Sub                                            'End Macro
    End If
    
    colsCnt = rngAll.Columns.Count                          'Total number of columns

    rngAll.Columns(colName).AdvancedFilter Action:=2, _
        CopyToRange:=Cells(1, colsCnt + 1), Unique:=1       'Pull unique names
        
    Columns(colsCnt + 1).SpecialCells(2).Offset(1). _
    Sort Cells(2, colsCnt + 1), 1                           'Sort.Ascend
    varTemp = Range(Cells(2, colsCnt + 1), _
    Cells(Rows.Count, colsCnt + 1).End(3))                  'Temp = Array of unique names
    
    If rngAll.Rows.Count = 2 Then                           'When there is only one data
        strName = varTemp
        Call move_data(colsCnt, rngAll, strName)            'Call Sub move_data()
    Else
    
        For s = 1 To UBound(varTemp, 1)                     'Repeat * (number of sheets)
            strName = varTemp(s, 1)                         'Designate Names
            Call move_data(colsCnt, rngAll, strName)        'Call Sub move_data()
        Next s
    End If
 

    Columns(colsCnt + 1).Delete                             'Delete Temp
    sorting_Sheets                                          'Run sub sorting_Sheets()
    
    Set rngAll = Nothing                                    'Empty Memory
    
    MsgBox "Macro has run successfully."                    'Show Message
End Sub

'-------------------------------------------------------------

' Adding Sheets and Moving Data by Using Advanced Filter

'-------------------------------------------------------------

Sub move_data(colsCnt As Integer, rngAll As Range, strName As String)

    Dim rngT As Range                                      '(T)arget of place where data is copied
    Dim sht As Worksheet
    
        On Error Resume Next                               'Ignore Error
        Set sht = Sheets(strName)                          'strName is name of a sheet
        
        If Err <> 0 Then                                   'When error because there is no sheet
            Sheets.Add after:=Sheets(Sheets.Count)         'Add sheet at the end
            Sheets(Sheets.Count).Name = strName            'Change the name of the sheet
            Sheets(1).Activate                             'Activate the data sheet(first sheet)
            Worksheets.FillAcrossSheets Range("A1", _
            Cells(1, Columns.Count).End(1)(1, 0))          'Copy name of sheets in every sheets
        End If
        
        With Sheets(strName)
            Set rngT = .Cells(Rows.Count, "A").End(3)(2)   '(T)arget of place where data is copied
            Cells(2, colsCnt + 1) = strName                'Filter by unique filter names

            Cells(2, colsCnt + 1) = "'=" & strName         'Add '= to bring only exact same data

            rngAll.AdvancedFilter Action:=2, CriteriaRange:=Cells(1, colsCnt + 1).Resize(2), _
            CopyToRange:=rngT, Unique:=0                   'Copy data according to the sheet's name
            rngT.EntireRow.Delete                          'Remove duplicated sheets
 
            .Columns.AutoFit                               'Autofit Column Width
            .UsedRange.Sort .Range("A2"), 1, Header:=True  'Sort every sheets
        End With

        Set sht = Nothing                                  'Reset variable
        Set rngT = Nothing                                 'Reset variable
    

End Sub

'-------------------------------------------------------------

' Sort Sheets as Ascending except Data Sheet

'-------------------------------------------------------------
Sub sorting_Sheets()

    Dim i As Long
    Dim j As Long
 
        For i = 2 To Sheets.Count
            For j = i + 1 To Sheets.Count
                If UCase(Sheets(j).Name) < UCase(Sheets(i).Name) Then
                    Sheets(j).Move Before:=Sheets(i)
                End If
            Next j
        Next i
 

End Sub

Sub export()

    Dim i, wbRcnt, RCnt, col As Integer
    Dim wb As Workbook
    Dim month, Reset As String
    
    month = InputBox("Month? [01~12]")
    Reset = InputBox("Reset? [Type RESET]")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    '           RESET = Clear every coulmns        '
    '         Else = Except user written data      '
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If Reset = "RESET" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Uh.xlsx", , , , "abi")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Pro_B.xlsx", , , , "h24br!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Pro_B.xlsx", , , , "h24br!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
            
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Yang.xlsx", , , , "Accessbio1")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RA.xlsx", , , , "g74ac#")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Finance.xlsx", , , , "abc123")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BD.xlsx", , , , "j4p4!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True

            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Purchase.xlsx", , , , "fp91d#")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BMO.xlsx", , , , "d4k82$")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\QAQC.xlsx", , , , "a2bw5@")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND_Baek.xlsx", , , , "malaria")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
            
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\CEO.xlsx", , , , "j4p4!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:I" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
    
    
    Else
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Uh.xlsx", , , , "abi")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Pro_B.xlsx", , , , "h24br!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
            
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Pro_B.xlsx", , , , "h24br!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Yang.xlsx", , , , "Accessbio1")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RA.xlsx", , , , "g74ac#")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Finance.xlsx", , , , "abc123")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BD.xlsx", , , , "j4p4!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True

            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Purchase.xlsx", , , , "fp91d#")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BMO.xlsx", , , , "d4k82$")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\QAQC.xlsx", , , , "a2bw5@")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
        
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND_Baek.xlsx", , , , "malaria")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
            
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\CEO.xlsx", , , , "j4p4!")
                ActiveSheet.Unprotect
                wb.Sheets("Raw Data").Range("A2:E" & 100).ClearContents
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowUsingPivotTables:=True
            wb.Close savechanges:=True
    End If
    
    
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Open and write data to corresponding card numbers '
    '       Last parameter of Open is the password      '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 2 To Worksheets.Count
        If Workbooks("report.xlsx").Sheets(i).Name = "107" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Uh.xlsx", , , , "abi")
            With Workbooks("report.xlsx").Sheets("107")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "1478" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Pro_B.xlsx", , , , "h24br!")
            With Workbooks("report.xlsx").Sheets("1478")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "9224" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Pro_B.xlsx", , , , "h24br!")
            With Workbooks("report.xlsx").Sheets("9224")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "1750" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Yang.xlsx", , , , "Accessbio1")
            With Workbooks("report.xlsx").Sheets("1750")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "3350" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RA.xlsx", , , , "g74ac#")
            With Workbooks("report.xlsx").Sheets("3350")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "1188" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RA.xlsx", , , , "g74ac#")
            With Workbooks("report.xlsx").Sheets("1188")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "7806" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Finance.xlsx", , , , "abc123")
            With Workbooks("report.xlsx").Sheets("7806")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "5934" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
            With Workbooks("report.xlsx").Sheets("5934")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "9943" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
            With Workbooks("report.xlsx").Sheets("9943")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
                
        If Workbooks("report.xlsx").Sheets(i).Name = "503" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND_Baek.xlsx", , , , "malaria")
            With Workbooks("report.xlsx").Sheets("503")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "6867" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BD.xlsx", , , , "j4p4!")
            With Workbooks("report.xlsx").Sheets("6867")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If

        If Workbooks("report.xlsx").Sheets(i).Name = "8830" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Purchase.xlsx", , , , "fp91d#")
            With Workbooks("report.xlsx").Sheets("8830")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If

        
        If Workbooks("report.xlsx").Sheets(i).Name = "8689" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BMO.xlsx", , , , "d4k82$")
            With Workbooks("report.xlsx").Sheets("8689")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "9914" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\QAQC.xlsx", , , , "a2bw5@")
            With Workbooks("report.xlsx").Sheets("9914")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "9930" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
            With Workbooks("report.xlsx").Sheets("9930")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "81" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
            With Workbooks("report.xlsx").Sheets("81")
                ActiveSheet.Unprotect
                RCnt = .UsedRange.Rows.Count
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
        
        If Workbooks("report.xlsx").Sheets(i).Name = "9948" Then
            Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\CEO.xlsx", , , , "j4p4!")
            With Workbooks("report.xlsx").Sheets("9948")
                ActiveSheet.Unprotect
                wbRcnt = wb.Worksheets("Raw Data").Cells(1048576, 1).End(xlUp).End(xlUp).Row + 1
                wb.Sheets("Raw Data").Range("B" & wbRcnt & ":E" & RCnt + wbRcnt - 2).Value = .Range("A2:D" & RCnt).Value
                wb.Sheets("Raw Data").Range("A" & wbRcnt & ":A" & RCnt + wbRcnt - 2).Value = .Range("G2:G" & RCnt).Value
                ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                    False, AllowUsingPivotTables:=True
            End With
            wb.Close savechanges:=True
        End If
    
    Next i

End Sub

Sub concatenate()
    Dim i, RCnt, RTotal As Integer
    Dim wb As Workbook
    Dim month, Reset As String
    month = InputBox("Month? [01~12]")
    RTotal = 2
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           Combine every workbooks into one         '
    '   Form.xlsx is an empty workbook only with header  '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Uh.xlsx", , , , "abi")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Pro_B.xlsx", , , , "h24br!")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Sales_Yang.xlsx", , , , "Accessbio1")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RA.xlsx", , , , "g74ac#")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Finance.xlsx", , , , "abc123")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BD.xlsx", , , , "j4p4!")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\Purchase.xlsx", , , , "fp91d#")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\BMO.xlsx", , , , "d4k82$")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\QAQC.xlsx", , , , "a2bw5@")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND.xlsx", , , , "yr53s$")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\RND_Baek.xlsx", , , , "malaria")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
    
    Set wb = Workbooks.Open("\\FILESERVER\Data File\File Server\01.Common\Level 4\Creditcard Transactions\2019\" & month & "\CEO.xlsx", , , , "j4p4!")
    With Workbooks("Form.xlsx").Sheets("Raw Data")

        For i = 2 To 100
            If wb.Sheets("Raw Data").Cells(i, 1) <> wb.Sheets("Raw Data").Cells(i + 1, 1) Then RCnt = i - 1
        Next
        .Range("A" & RTotal & ":I" & RTotal + RCnt - 1).Value = wb.Sheets("Raw Data").Range("A2:I" & RCnt + 1).Value
        RTotal = RTotal + RCnt
    End With
    wb.Close savechanges:=False
End Sub
