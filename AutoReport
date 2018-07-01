Sub AutoReport()

    If MsgBox("Do you want to run your report? Please run EOD first if you have not already.", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim rptExportType As String
    
    Dim Datee As String, DateArray() As String, AlohaDatedSubDir As String
    
    Dim RowNumTax As Integer, totalSales As Integer, sales As Integer
    Dim RowNumCustCount1 As Integer, RowNumCustCount2 As Integer, RowNumCustCount3 As Integer
    Dim RowNumTakeOut As Integer
    
    Dim RowNumLabor As Integer, ColNumLabor As Integer
    Dim RowNumDeposit As Integer, RowNumOS As Integer, RowNumCC As Integer
    Dim ColNum As Integer, RowNumGCR As Integer, RowNumGCA As Integer
    
    Dim RowNumCaterPaid As Integer, RowNumCaterPend As Integer, RowNumEZCater As Integer
    
    Dim DeclTips As Double, CCTips As Double, ColNumCCTips As Integer, ColNumDeclTips As Integer
    
    Dim Netsales As Double, Tax As Double, searchRange As Range

    Dim wbSource As Workbook, wbDest As Workbook
    Dim wsFrom As Worksheet, wsTo As Worksheet
    
    Set wbDest = ActiveWorkbook
    Set wsTo = wbDest.ActiveSheet
    wsTo.Unprotect
    
    Datee = wsTo.Cells(2, 11).Value
    AlohaDatedSubDir = GenAlohaDate(Datee)
    If AlohaDatedSubDir = vbNullString Then
        MsgBox "Enter the date"
        Exit Sub
    End If

    If Dir("C:\Aloha\" + AlohaDatedSubDir, vbDirectory) = vbNullString Then
        MsgBox "Check the date on your excel worksheet. If it is correct, make sure that End of Day has been run"
    Else

    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /RC /LOAD ""DEFAULT.SLS.SET""", windowStyle, waitOnReturn 'print sales
    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /XC /LOAD ""DEFAULT.SLS.SET""", windowStyle, waitOnReturn 'export sales
    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /RH /LOAD ""DEFAULT.SLS.SET""", windowStyle, waitOnReturn 'print hourly labor
    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /XA /LOAD ""ZDO NOT USE.LBR.SET""", windowStyle, waitOnReturn 'export labor
    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /RA /LOAD ""ZDO NOT USE.LBR.SET""", windowStyle, waitOnReturn 'print labor
    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /RV", windowStyle, waitOnReturn 'print void
    wsh.Run "CMD.EXE /S /C" & "ALOHA\BIN\RPT.EXE /DATE " & AlohaDatedSubDir & " /XH /LOAD ""DEFAULT.SLS.SET""", windowStyle, waitOnReturn 'export hourly sales/labor
    
    
    'Open DSR
    Set wbSource = Workbooks.Open("C:\Aloha\RptExport\Default.sls.csv")
    'Set wbSource = Workbooks.Open("C:\Users\JWooten\Desktop\Default.sls.csv")
    'Search for the row of target data (sales)
    Set wsFrom = wbSource.Sheets(1)
     
    With wsFrom
     ColNum = .Range("A1:J15").Find(What:="Gross Sales", LookIn:=xlValues).Column
     RowNumTotalSales = .Range("A1:J15").Find(What:="(less Voids Comps Promos Surch.", LookIn:=xlValues).Row
     RowNumTax = .Range("A20:Z24").Find(What:="TAX", LookIn:=xlValues).Row
     RowNumCustCount1 = .Range("A33:Z43").Find(What:="7-3 Shift", LookIn:=xlValues).Row
     RowNumCustCount2 = .Range("A33:Z43").Find(What:="3-11 Shift", LookIn:=xlValues).Row
     RowNumCustCount3 = .Range("A33:Z43").Find(What:="11-7 Shift", LookIn:=xlValues).Row
     RowNumSales = .Range("A220:Z260").Find(What:="Totals", LookIn:=xlValues).Row
     RowNumTakeOut = .Range("A70:Z90").Find(What:="To Go", LookIn:=xlValues).Row
     RowNumDeposit = .Range("A:Z").Find(What:="Deposit (calculated)", LookIn:=xlValues).Row
     RowNumOS = .Range("A:Z").Find(What:="Deposit O/S", LookIn:=xlValues).Row
     RowNumCC = .Range("A161:Z210").Find(What:="Charge", LookIn:=xlValues).Row
     RowNumGCA = .Range("A200:Z270").Find(What:="Gift Card", LookIn:=xlValues).Row
     RowNumGCR = .Range("A:Z").Find(What:="GIFT CRT", LookIn:=xlValues).Row
    End With
    
    'copy target data to specified cell and lock said cell
    wsTo.Range("F1").Locked = False
    wsTo.Cells(1, 2) = wsFrom.Cells(RowNumTotalSales, ColNum + 1).Value
    wsTo.Range("A1").Locked = True
    wsTo.Cells(1, 6) = wsFrom.Cells(RowNumTotalSales, ColNum + 1).Value
    wsTo.Range("F1").Locked = True
    wsTo.Cells(2, 6) = wsFrom.Cells(RowNumTax, ColNum + 1).Value
    wsTo.Range("F2").Locked = True
    wsTo.Cells(35, 8) = wsFrom.Cells(RowNumCustCount1, ColNum + 1).Value
    wsTo.Range("H35").Locked = True
    wsTo.Cells(36, 8) = wsFrom.Cells(RowNumCustCount2, ColNum + 1).Value
    wsTo.Range("H36").Locked = True
    wsTo.Cells(37, 8) = wsFrom.Cells(RowNumCustCount3, ColNum + 1).Value
    wsTo.Range("H37").Locked = True
    wsTo.Cells(37, 2) = wsFrom.Cells(RowNumTakeOut, ColNum + 1).Value
    wsTo.Range("B37").Locked = True
    wsTo.Cells(10, 11) = wsFrom.Cells(RowNumCC, ColNum + 1).Value
    wsTo.Range("K10").Locked = True
    wsTo.Cells(7, 2) = wsFrom.Cells(RowNumGCR, ColNum + 1).Value
    wsTo.Range("B7").Locked = True
    wsTo.Cells(8, 2) = wsFrom.Cells(RowNumGCA, ColNum + 4).Value
    wsTo.Range("B8").Locked = True
    wsTo.Range("B40").NumberFormat = ".##"
    
    'set deposit to (calculated deposit - tax) and set O/U formula
    wsTo.Cells(10, 2) = (wsFrom.Cells(RowNumDeposit, ColNum + 1).Value + wsFrom.Cells(RowNumOS, ColNum + 1).Value - wsFrom.Cells(RowNumTax, ColNum + 1).Value)
    wsTo.Cells(9, 2) = "=(B1-B2-B3-B4-B5-B6-B7+B8-B10)*-1"
    wsTo.Range("B9").Locked = True
    wsTo.Range("B10").Locked = False
   
   Workbooks("Default.sls.csv").Close SaveChanges:=False
   
   'open labor report
    Set wbSource = Workbooks.Open("C:\Aloha\RptExport\Default 1.lbr.csv")
    'Set wbSource = Workbooks.Open("C:\Users\JWooten\Desktop\Daily.lbr.csv")
   
   'find row number of target data (labor)
    With wbSource.Sheets(1)
     RowNumLabor = .Range("A:C").Find(What:="TOTALS:", LookIn:=xlValues).Row + 1
     ColNumLabor = .Range("A1:Z12").Find(What:="Pay", LookIn:=xlValues).Column
     ColNumCCTips = .Range("A1:Z12").Find(What:="CC", LookIn:=xlValues).Column
     ColNumDeclTips = .Range("A1:Z12").Find(What:="Decl", LookIn:=xlValues).Column
    End With
    
    Set wsFrom = wbSource.Sheets(1)
    CCTips = wsFrom.Cells(RowNumLabor, ColNumCCTips).Value
    DeclTips = wsFrom.Cells(RowNumLabor, ColNumDeclTips).Value
    
    'copy target data to specified cell and lock said cell
    wsTo.Cells(14, 6) = wsFrom.Cells(RowNumLabor, ColNumLabor).Value
    wsTo.Range("F14").Locked = True
    wsTo.Cells(34, 11) = CCTips
    wsTo.Range("K34").Locked = True
    wsTo.Cells(35, 11) = (DeclTips - CCTips)
    wsTo.Range("K35").Locked = True
    
    wsTo.Cells(1, 12) = "Automated"
    
    Workbooks("Default 1.lbr.csv").Close SaveChanges:=False
    'end labor copy
    
    MsgBox ("Report Completed." & vbNewLine & "Please Enter:" & vbNewLine & "1) Credit Card sales by TERMINAL" & vbNewLine & "2) Check main deposit" & vbNewLine & "3) Paid outs" & vbNewLine & "4) Invoices" & vbNewLine & "5) Sales by shift")
    
    wsTo.Protect
    
    If MsgBox("Do you want to run your hourly sales and labor?", vbYesNo + vbQuestion) = vbYes Then
        Call AutoPP
    End If
    
    Call SendEmail
    
    End If

End Sub



Public Function GenAlohaDate(ByVal Datee As String) As String
    'generate date in aloha subdir format

     If Datee = vbNullString Then
        Exit Function
     End If
     DateArray = Split(Datee, "/")

     If Len(DateArray(0)) < 2 Then
        DateArray(0) = "0" + DateArray(0)
        End If
    
    If Len(DateArray(1)) < 2 Then
        DateArray(1) = "0" + DateArray(1)
        End If
        
    GenAlohaDate = DateArray(2) + DateArray(0) + DateArray(1)
    'end date formatting
End Function







