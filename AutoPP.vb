Sub AutoPP()
 
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim rptExportType As String
    
    Dim Datee As String, AlohaDatedSubDir As String, saveDate As String, day As String
    Dim CurrentDayRow As Integer, lastSundayDate As String, sundayDateArray() As String
    
    Dim fiveAMRow As Integer, twoPMRow As Integer, tenPMRow As Integer, fourAMRow
    Dim netsalesCol As Integer, laborDollarsCol As Integer
    
    Dim sales1Col As Integer, sales2Col As Integer, sales3Col As Integer, ColNum As Integer
    Dim Labor1Col As Integer, Labor2Col As Integer, Labor3Col As Integer
    
    Dim sales1 As Double, sales2 As Double, sales3 As Double
    Dim Labor1 As Double, Labor2 As Double, Labor3 As Double
    Dim grossSales1 As Double, grossSales2 As Double, grossSales3 As Double
    Dim Netsales As Double, Tax As Double
    
    Dim wbSource As Workbook, wbDSR As Workbook, wbDest As Workbook, wbMaster As Workbook
    Dim wsFrom As Worksheet, wsDSR As Worksheet, wsTo As Worksheet
    
    Set wbDSR = ActiveWorkbook
    Set wsDSR = ActiveSheet
    
    'Set wbDest = Workbooks.Open("C:\Users\JWooten\Desktop\MASTER PAYROLL.xls")
    Set wbDest = Workbooks.Open("C:\Documents and Settings\Administrator\Desktop\#7 PAYROLL.xls")
    Set wsTo = wbDest.Sheets(1)
    
    'yesterday's date/day
    Datee = wsDSR.Cells(2, 11).Value
    day = wsDSR.Cells(4, 11).Value
    
    If Datee = vbNullString Then
        MsgBox "Enter the date on your worksheet."
        Exit Sub
     End If
      If day = vbNullString Then
        MsgBox "Enter the day on your worksheet."
        Exit Sub
     End If
       
    AlohaDatedSubDir = GenAlohaDate(Datee)
    'end date formatting
    
    wsTo.Unprotect
    
     With wsTo
     CurrentDayRow = .Range("C2:E9").Find(What:="", LookIn:=xlValues).Row
    End With
 
    'check for next week
    'save a copy of this wb and erase contents
     If wsTo.Cells(CurrentDayRow, 2) = "next week" Then
        lastSundayDate = wsTo.Cells(2, 1).Value
        sundayDateArray = Split(lastSundayDate, "/")
        saveDate = sundayDateArray(0) & "-" & sundayDateArray(1) & "-" & sundayDateArray(2)
        'wbDest.SaveCopyAs ("C:\Users\JWooten\Desktop\PAYROLL " & saveDate & ".xls")
        wbDest.SaveCopyAs ("C:\Documents and Settings\Administrator\Desktop\payroll percentage\PAYROLL " & saveDate & ".xls")
        wsTo.Cells(2, 1) = Datee
        wsTo.Range("C2:D8").ClearContents
        wsTo.Range("G2:H8").ClearContents
        wsTo.Range("K2:L8").ClearContents
        CurrentDayRow = 2
     End If
     
    With wsTo
     sales1Col = .Range("A:Z").Find(What:="7-3 Sales", LookIn:=xlValues).Column
     sales2Col = .Range("A:Z").Find(What:="3-11 Sales", LookIn:=xlValues).Column
     sales3Col = .Range("A:Z").Find(What:="11-7 Sales", LookIn:=xlValues).Column
     Labor1Col = .Range("A:Z").Find(What:="7-3 Payroll", LookIn:=xlValues).Column
     Labor2Col = .Range("A:Z").Find(What:="3-11 Payroll", LookIn:=xlValues).Column
     Labor3Col = .Range("A:Z").Find(What:="11-7 Payroll", LookIn:=xlValues).Column
    End With
    
    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /RH /LOAD ""DEFAULT.SLS.SET""", windowStyle, waitOnReturn 'print

    Set wbSource = Workbooks.Open("C:\Aloha\RptExport\Default.hrs.csv")
    'Set wbSource = Workbooks.Open("C:\Users\JWooten\Desktop\Default.hrs.csv")
    'Search for the row of target data (sales)
    Set wsFrom = wbSource.Sheets(1)

    With wsFrom
        fourAMRow = .Range("A:D").Find(What:="4:00 AM", LookIn:=xlValues).Row
        fiveAMRow = .Range("A:D").Find(What:="5:00 AM", LookIn:=xlValues).Row
        twoPMRow = .Range("A:D").Find(What:="12:00 PM", LookIn:=xlValues).Row + 2
        tenPMRow = .Range("A:D").Find(What:="10:00 PM", LookIn:=xlValues).Row
        netsalesCol = .Range("A:D").Find(What:="Net sales", LookIn:=xlValues).Column
        laborDollarsCol = .Range("A:P").Find(What:="Labor $", LookIn:=xlValues).Column
    End With
    
    sales1 = Application.Sum(Range(Cells(fiveAMRow, netsalesCol), Cells(twoPMRow, netsalesCol)))
    sales2 = Application.Sum(Range(Cells(twoPMRow + 1, netsalesCol), Cells(tenPMRow, netsalesCol)))
    sales3 = Application.Sum(Range(Cells(tenPMRow + 1, netsalesCol), Cells(fourAMRow, netsalesCol)))
    Labor1 = Application.Sum(Range(Cells(fiveAMRow, laborDollarsCol), Cells(twoPMRow, laborDollarsCol)))
    Labor2 = Application.Sum(Range(Cells(twoPMRow + 1, laborDollarsCol), Cells(tenPMRow, laborDollarsCol)))
    Labor3 = Application.Sum(Range(Cells(tenPMRow + 1, laborDollarsCol), Cells(fourAMRow, laborDollarsCol)))
    
    wsTo.Unprotect
    
    wsTo.Cells(CurrentDayRow, sales1Col) = sales1
    wsTo.Cells(CurrentDayRow, sales2Col) = sales2
    wsTo.Cells(CurrentDayRow, sales3Col) = sales3
    
    wsTo.Cells(CurrentDayRow, Labor1Col) = Labor1
    wsTo.Cells(CurrentDayRow, Labor2Col) = Labor2
    wsTo.Cells(CurrentDayRow, Labor3Col) = Labor3
    
    wbDest.PrintOut
    wsTo.Protect
    
    wbDest.Close (True)
    Workbooks("Default.hrs.csv").Close SaveChanges:=False
    
    'Set wbSource = Workbooks.Open("C:\Users\Jwooten\Desktop\Default.sls.csv")
    Set wbSource = Workbooks.Open("C:\Aloha\RptExport\Default.sls.csv")
    Set wsFrom = wbSource.Sheets(1)
    
    With wsFrom
     RowNumTax = .Range("A20:H24").Find(What:="TAX", LookIn:=xlValues).Row
     ColNum = .Range("A1:J15").Find(What:="Gross sales", LookIn:=xlValues).Column
    End With
    
    Tax = wsFrom.Cells(RowNumTax, ColNum + 1).Value
    
    Netsales = sales1 + sales2 + sales3
    
    Shift1Percent = sales1 / Netsales
    Shift2Percent = sales2 / Netsales
    Shift3Percent = sales3 / Netsales
    
    grossSales1 = sales1 + (Shift1Percent * Tax)
    grossSales2 = sales2 + (Shift2Percent * Tax)
    grossSales3 = sales3 + (Shift3Percent * Tax)
    'end calculate tax
    
    wsDSR.Unprotect
    
    wsDSR.Cells(6, 11) = grossSales1
    wsDSR.Cells(7, 11) = grossSales2
    wsDSR.Cells(8, 11) = grossSales3
    wsDSR.Cells(9, 11) = (grossSales1 + grossSales2 + grossSales3)
    wsDSR.Range("K6:K9").NumberFormat = ".##"
    
    wsDSR.PrintOut
    wsDSR.Protect
    
    Workbooks("Default.sls.csv").Close SaveChanges:=False
End Sub








