Attribute VB_Name = "ModuleParticular"
Option Explicit

'���������� ���������
Global lngItemID As Long
Global blnPrintHour As Boolean
Global blnPrintBalance As Boolean
Global blnRoundBuys As Boolean
Global blnRoundSales As Boolean
Global bytRoundCents As Byte
Global strTransportReason As String
Global strTransportWay As String
Global strLoadingSite As String
Global strDestinationSite As String
Global blnCheckTaxNo As Boolean
Global curExtraChargesVATPercent As Currency
Global intSalesInvoiceLines As Integer
Global blnCheckEAFDSS As Boolean
Global strEAFDSS As String
Global datClosedPeriod As Date
Global strSender As String
Global strServer As String
Global strUserName As String
Global strPassword As String

Global curGrandTotal() As Currency

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Const TH32CS_SNAPPROCESS As Long = &H2


Function InitReport(myPrinterType, myEAFDSSString, myInvoiceHeight)

    Dim intTopMargin As Integer
        
    If myPrinterType = "1" Then
        Print #1, Chr(27); Chr(64)
        Print #1, Chr(27); Chr(67); Chr(myInvoiceHeight);
    End If
    
    If myEAFDSSString <> "" Then
        Print #1, myEAFDSSString
    End If
    
End Function


Function IsProcessRunning(strProcess)
    
    Dim processInfo As PROCESSENTRY32
    Dim hSnapshot As Long
    Dim success As Long
    Dim retval As Long
    Dim exeName As String
    
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    
    processInfo.dwSize = Len(processInfo)
    success = Process32First(hSnapshot, processInfo)
    
    If hSnapshot <> -1 Then
        Do While success <> 0
            exeName = Left(processInfo.szExeFile, InStr(processInfo.szExeFile, vbNullChar) - 1)
            processInfo.dwSize = Len(processInfo)
            success = Process32Next(hSnapshot, processInfo)
            If UCase(exeName) = UCase(strProcess) Then
                IsProcessRunning = True
                Exit Do
            End If
        Loop
        retval = CloseHandle(hSnapshot)
    End If

End Function

Function ShowPersonLedger(myPersonID, myPersonDescription, myWindowTitle, myTable, myOppositeTable, myRefersTo)

    With PersonsLedger
        .txtPersonID.text = myPersonID
        .txtPersonDescription.text = myPersonDescription
        .lblTitle.Caption = myWindowTitle
        .txtTable.text = myTable
        .txtOppositeTable.text = myOppositeTable
        .txtRefersTo.text = myRefersTo
        .Tag = "True"
        DisableFields .txtPersonDescription, .cmdIndex(0)
        .Show 1
    End With
    
End Function

Function FillArray(strArrayName, ParamArray myColumns() As Variant)

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myColumns())
        strArrayName(intLoop) = strArrayName(intLoop) + myColumns(intLoop)
    Next intLoop
    
End Function

Function DoRunningTotal(strArrayName, ParamArray Columns() As Variant)

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(Columns)
        If Columns(intLoop) <> "" Then
            strArrayName(intLoop) = strArrayName(intLoop) + Columns(intLoop)
        End If
    Next intLoop
    
End Function

Function CalculateGrandTotals(ParamArray myFields() As Variant)

    Dim intLoop As Integer
    
    For intLoop = 0 To UBound(myFields)
        curGrandTotal(intLoop) = curGrandTotal(intLoop) + myFields(intLoop)
    Next intLoop
    
End Function

Function AddGridRowWithTotals(myGrid As iGrid, myOnlyQty, myMessageColumn, myMessage, mySums, myColumnCount, myHowManyBlankLinesBefore, myHowManyBlankLinesAfter, ParamArray myColumns() As Variant)

    Dim intLoop As Integer
    Dim lngRow As Long
    
    If myHowManyBlankLinesBefore > 0 Then
        myGrid.AddRow , , , , , , , myHowManyBlankLinesBefore
    End If
    
    lngRow = myGrid.RowCount
    
    myGrid.CellValue(lngRow, myMessageColumn) = myMessage
    
    For intLoop = 0 To myColumnCount
        myGrid.CellValue(lngRow, myColumns(intLoop)) = mySums(intLoop)
    Next intLoop
    
    If myHowManyBlankLinesAfter > 0 Then
        myGrid.AddRow , , , , , , , myHowManyBlankLinesAfter
    End If
    
End Function

Function CalculateDebitCreditAndBalance(myDebitOrCredit, myPerson, myInvoiceGrossAmount, myCodeCustomers, myCodeSuppliers, myCodeInventoryQtyOrAmount, myPaymentWayCreditID, myRefersTo)

    CalculateDebitCreditAndBalance = 0
    
    If myPerson <> "Items" Then
        
        '������
        If myDebitOrCredit = "Debit" Then
            '������ �� �������
            If myRefersTo = 1 And myCodeSuppliers = "+" And myPaymentWayCreditID = 0 Then
                CalculateDebitCreditAndBalance = myInvoiceGrossAmount
            End If
            '�������� - ������ ������ � ����������� - ������ ���������
            If myRefersTo = 2 And myCodeCustomers = "+" Or (myRefersTo = 3 And myCodeSuppliers = "-") Then
                CalculateDebitCreditAndBalance = myInvoiceGrossAmount
            End If
            '�������� - ������ ������  - �� ����� ������� � ����������� - ������ ��������� - �� ����� �������
            If (myRefersTo = 2 And myCodeCustomers = "-") Or (myRefersTo = 3 And myCodeSuppliers = "+") Then
                CalculateDebitCreditAndBalance = -myInvoiceGrossAmount
            End If
            '���������
            Exit Function
        End If
        
        '�������
        If myDebitOrCredit = "Credit" Then
            '�������� �� �������
            If myRefersTo = 2 And myCodeCustomers = "+" And myPaymentWayCreditID = 0 Then
                CalculateDebitCreditAndBalance = myInvoiceGrossAmount
            End If
            '������ - ������ ������ � ������� - ������ ���������
            If (myRefersTo = 1 And myCodeSuppliers = "+") Or (myRefersTo = 4 And myCodeCustomers = "-") Then
                CalculateDebitCreditAndBalance = myInvoiceGrossAmount
            End If
            '������ - ������ ������ - �� ����� ������� � ������� - ������ ��������� - �� ����� �������
            If (myRefersTo = 1 And myCodeSuppliers = "-") Or (myRefersTo = 4 And myCodeCustomers = "+") Then
                CalculateDebitCreditAndBalance = -myInvoiceGrossAmount
            End If
            '���������
            Exit Function
        End If
        
    End If
    
    If myPerson = "Items" Then
    
        If myDebitOrCredit = "Debit" Then
            If (myCodeInventoryQtyOrAmount = "+") Then
                CalculateDebitCreditAndBalance = myInvoiceGrossAmount
            End If
        End If
        
        If myDebitOrCredit = "Credit" Then
            If (myCodeInventoryQtyOrAmount = "-") Then
                CalculateDebitCreditAndBalance = myInvoiceGrossAmount
            End If
        End If
    
    End If

End Function

Function AddOneToTheLastRecord()

    Dim strSQL As String
    Dim rsInvoices As Recordset
    
    strSQL = "SELECT InvoiceTrnID FROM Invoices ORDER BY InvoiceTrnID"
    Set rsInvoices = CommonDB.OpenRecordset(strSQL)
    
    With rsInvoices
        .MoveLast
        AddOneToTheLastRecord = IIf(.EOF, 1, !invoiceTrnID + 1)
    End With
    
    rsInvoices.Close
    Set rsInvoices = Nothing

End Function

Function CalculatePersonBalance(tmpTable, tmpCode)
    
    Dim strTable As String
    Dim curPreviousBalance As Currency
    Dim rstInvoices As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    TempQuery.SQL = "PARAMETERS intPerson Integer; " _
    & "SELECT InvoicePersonID, CodeID, InvoicePaymentWay, InvoiceGross, Codes.[Code" & tmpTable & "] AS Column, Codes.[CodeRefersTo], PaymentWays.[PaymentWayCredit] " _
    & "FROM ((Invoices " _
    & "INNER JOIN " & tmpTable & " ON Invoices.InvoicePerson = " & tmpTable & ".ID) " _
    & "INNER JOIN Codes ON Invoices.InvoiceInvoiceID= Codes.CodeID) " _
    & "INNER JOIN PaymentWays ON Invoices.InvoicePaymentWay = PaymentWays.[PaymentWayID] " _
    & "WHERE InvoicePerson = [intPerson] AND PaymentWayCredit = True AND (Codes.[Code" & tmpTable & "] = '+' or Codes.[Code" & tmpTable & "] = '-')"
    TempQuery![intPerson] = Val(tmpCode)
    
    Set rstInvoices = TempQuery.OpenRecordset()
    With rstInvoices
        Do While Not .EOF
            If ![CodeRefersTo] = 1 Or ![CodeRefersTo] = 3 Then
                If ![Column] = "+" Then
                    curPreviousBalance = curPreviousBalance + ![InvoiceGross]
                Else
                    curPreviousBalance = curPreviousBalance - ![InvoiceGross]
                End If
            End If
            If ![CodeRefersTo] = 0 Or ![CodeRefersTo] = 2 Then
                If ![Column] = "+" Then
                    curPreviousBalance = curPreviousBalance - ![InvoiceGross]
                Else
                    curPreviousBalance = curPreviousBalance + ![InvoiceGross]
                End If
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    CalculatePersonBalance = curPreviousBalance
    
    Exit Function
    
End Function

Function CalculateItemBalance(tmpCode)

    Dim intBalance As Integer
    Dim rstTransactions As Recordset
    
    TempQuery.SQL = "PARAMETERS intItemCode Integer; " _
        & "SELECT ItemID, Qty, Codes.CodeInventoryQty, Codes.CodeRefers, Invoices.InvoiceTrnID " _
        & "FROM (InvoicesTrn " _
        & "INNER JOIN Invoices ON Invoices.InvoiceTrnID = InvoicesTrn.InvoiceID) " _
        & "INNER JOIN Codes ON Invoices.CodeID = Codes.CodeID " _
        & "WHERE ItemID = intItemCode"
    TempQuery![intItemCode] = Val(tmpCode)

    Set rstTransactions = TempQuery.OpenRecordset()
    
    With rstTransactions
        
        Do Until .EOF
            If ![CodeInventoryQty] = "+" Then
                intBalance = intBalance + ![Qty]
            Else
                If ![CodeInventoryQty] = "-" Then
                    intBalance = intBalance - ![Qty]
                End If
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    CalculateItemBalance = Format(intBalance, "#,##0")

End Function

Function CheckForItemMatch(txtCategoryID, txtManufacturerID, txtItemID, txtItemShortDescription, txtTable, strField, lngOrder, bytShowInList, blnShowList) As Recordset

    Dim blnCriteria As Boolean
    
    Dim intIndex As Integer
    Dim strSQL As String
    Dim strThisParameter As String
    Dim strThisQuery As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strParameters As String
    Dim strParFields As String
    Dim strOrder As String
    
    blnCriteria = False
    
    Set TempQuery = CommonDB.CreateQueryDef("")
        
    strSQL = "SELECT ItemID, CategoryID, CategoryShortDescription, CategoryDescription, CategoryCheckBalance, CategoryItemDescriptionRequired, ItemDescription,  ManufacturerDescription, ItemVAT, ItemBalance, ItemActive  " _
        & "FROM ((" & txtTable & " " _
        & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) " _
        & "INNER JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID) "
        
    strOrder = " ORDER BY ManufacturerDescription, ItemDescription, CategoryDescription"
        
    If txtCategoryID <> "" Then
        blnCriteria = True
        strThisParameter = "lngCategoryID Long"
        strThisQuery = "Items.[ItemCategoryID] = lngCategoryID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCategoryID)
    End If
    
    If txtManufacturerID <> "" Then
        blnCriteria = True
        strThisParameter = "lngManufacturerID Long"
        strThisQuery = "Items.[ItemManufacturerID] = lngManufacturerID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtManufacturerID)
    End If
    
    If Left(txtItemShortDescription, 1) <> "*" And Len(txtItemShortDescription) > 0 Then
        blnCriteria = True
        strThisParameter = "strItemShortDescription String"
        strThisQuery = "Left(Items![ItemQuickDescription],Len(strItemShortDescription)) = strItemShortDescription"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = txtItemShortDescription
    End If
    
    If Left(txtItemShortDescription, 1) = "*" Then
        blnCriteria = True
        strThisParameter = "strItemShortDescription String"
        strThisQuery = "InStr(Items!ItemQuickDescription, " & "'" & Right(txtItemShortDescription, Len(txtItemShortDescription) - 1) & "'" & ") "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = txtItemShortDescription
    End If
    
    If blnCriteria Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strOrder
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    Else
        TempQuery.SQL = strSQL & strOrder
    End If
    
    Set CheckForItemMatch = TempQuery.OpenRecordset()
    
    Exit Function
        
UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    Return

End Function

Function CheckForInvoiceExist(tmpStatus, tmpDate, tmpPersonID, tmpInvoiceCodeID, tmpInvoiceNo)

    Dim rstTransactions As Recordset
    Dim intRecordCount As Integer
    
    CheckForInvoiceExist = False
    
    TempQuery.SQL = "PARAMETERS datInvoiceIssueDate Date, lngInvoicePersonID Long, lngInvoiceCodeID Long, lngInvoiceNo Long; " _
    & "SELECT * FROM Invoices " _
    & "WHERE InvoiceIssueDate = datInvoiceIssueDate AND InvoicePersonID = lngInvoicePersonID AND InvoiceCodeID = lngInvoiceCodeID AND InvoiceNo = lngInvoiceNo"
    
    TempQuery!datInvoiceIssueDate = CDate(tmpDate)
    TempQuery!lngInvoicePersonID = Val(tmpPersonID)
    TempQuery!lngInvoiceCodeID = Val(tmpInvoiceCodeID)
    TempQuery!lngInvoiceNo = Val(tmpInvoiceNo)
    
    Set rstTransactions = TempQuery.OpenRecordset()
    
    intRecordCount = IIf(tmpStatus, 0, 1)
    
    If Not rstTransactions.EOF Then
        rstTransactions.MoveLast
        If rstTransactions.RecordCount > intRecordCount Then
            CheckForInvoiceExist = True
        End If
    End If

End Function

Sub PrintAsciiFile(strAsciiFile, strPrinterPort)

    Open App.Path & "\PrintCommand.Bat" For Output As #1
    
    Print #1, "print /d:" & strPrinterPort & " " & strAsciiFile
    Close #1
    
    Shell App.Path & "\PrintCommand.Bat"
    
End Sub

Function LoadSettings()
    
    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim intUpper As Integer
    
    Dim TableSettings As TableDef
    
    Dim rsParameters As Recordset
    
    Set TableSettings = dBaseTables("Settings")
    Set rsParameters = TableSettings.OpenRecordset()
    
    With rsParameters
        .MoveFirst
        '�������
        arrCompanyData(1) = !Line01
        arrCompanyData(2) = !Line02
        arrCompanyData(3) = !Line03
        arrCompanyData(4) = !Line04
        arrCompanyData(5) = !Line05
        arrCompanyData(6) = !Line06
        '��������
        arrCompanyData(7) = !Line07
        arrCompanyData(8) = !Line08
        arrCompanyData(9) = !Line09
        arrCompanyData(10) = !Line10
        blnPreviewReports = !PreviewReportsID
        '��������
        blnRoundSales = !RoundSalesID
        bytRoundCents = !RoundSalesCents
        curExtraChargesVATPercent = !ExtraChargesVATPercent
        strTransportReason = !TransportReason
        strTransportWay = !TransportWay
        strLoadingSite = !LoadingSite
        strDestinationSite = !DestinationSite
        blnPrintHour = IIf(!PrintHourID = 1, True, False)
        blnPrintBalance = IIf(!PrintBalanceID = 1, True, False)
        intSalesInvoiceLines = !SalesInvoiceLines
        '��������������
        blnCheckTaxNo = IIf(!TaxNoCheckID = 1, True, False)
        '������
        blnCheckEAFDSS = IIf(!EAFDSSCheckID = 1, True, False)
        strEAFDSS = !EAFDSSProcessName
        '��������� ��������
        datClosedPeriod = !ClosedPeriod
        'Email
        strSender = !EmailSender
        strServer = !EmailServer
        strUserName = !EmailUserName
        strPassword = !EmailPassword
        '�������
        strBankAccountNo = !BankAccountNo
        '�����
        .Close
    End With
    
    LoadSettings = True
    
    Exit Function
    
ErrTrap:
    LoadSettings = False
    DisplayErrorMessage True, Err.Description
    
End Function

Public Function FullNumber(tmpOldNumber)
    
    'Local ����������
    Dim intLoop As Byte
    Dim aArray(9, 10) As String
    Dim strTotalGross As String
    Dim strSubNumber As String
    Dim tmpDecNumber As String
    Dim strFullNumber As String
    Dim strDecNumber As String
    Dim bytArrayIndex As Byte
    Dim tmpIntNumber As Long
    Dim tmpNumber As String
    Dim aFullNumber(9) As String
    
    '������� �����
    bytArrayIndex = 1
   
    aArray(1, 1) = " "
    aArray(1, 2) = "������ "
    aArray(1, 3) = "�������� "
    aArray(1, 4) = "��������� "
    aArray(1, 5) = "���������� "
    aArray(1, 6) = "���������� "
    aArray(1, 7) = "�������� "
    aArray(1, 8) = "��������� "
    aArray(1, 9) = "��������� "
    aArray(1, 10) = "���������� "
    
    aArray(2, 1) = " "
    aArray(2, 2) = "���� "
    aArray(2, 3) = "������ "
    aArray(2, 4) = "������� "
    aArray(2, 5) = "������� "
    aArray(2, 6) = "������� "
    aArray(2, 7) = "������ "
    aArray(2, 8) = "��������� "
    aArray(2, 9) = "������� "
    aArray(2, 10) = "�������� "
    
    aArray(3, 1) = " "
    aArray(3, 2) = "��� "
    aArray(3, 3) = "��� "
    aArray(3, 4) = "���� "
    aArray(3, 5) = "������� "
    aArray(3, 6) = "����� "
    aArray(3, 7) = "��� "
    aArray(3, 8) = "���� "
    aArray(3, 9) = "���� "
    aArray(3, 10) = "����� "
    
    aArray(4, 1) = " "
    aArray(4, 2) = "������ "
    aArray(4, 3) = "��������� "
    aArray(4, 4) = "���������� "
    aArray(4, 5) = "����������� "
    aArray(4, 6) = "����������� "
    aArray(4, 7) = "��������� "
    aArray(4, 8) = "���������� "
    aArray(4, 9) = "���������� "
    aArray(4, 10) = "����������� "
    
    aArray(5, 1) = " "
    aArray(5, 2) = "���� "
    aArray(5, 3) = "������"
    aArray(5, 4) = "������� "
    aArray(5, 5) = "������� "
    aArray(5, 6) = "������� "
    aArray(5, 7) = "������ "
    aArray(5, 8) = "��������� "
    aArray(5, 9) = "������� "
    aArray(5, 10) = "�������� "
    
    aArray(6, 1) = " "
    aArray(6, 2) = "��� "
    aArray(6, 3) = "��� "
    aArray(6, 4) = "����� "
    aArray(6, 5) = "�������� "
    aArray(6, 6) = "����� "
    aArray(6, 7) = "��� "
    aArray(6, 8) = "���� "
    aArray(6, 9) = "���� "
    aArray(6, 10) = "����� "
    
    aArray(7, 1) = " "
    aArray(7, 2) = "������ "
    aArray(7, 3) = "�������� "
    aArray(7, 4) = "��������� "
    aArray(7, 5) = "���������� "
    aArray(7, 6) = "���������� "
    aArray(7, 7) = "�������� "
    aArray(7, 8) = "��������� "
    aArray(7, 9) = "���������"
    aArray(7, 10) = "���������� "
    
    aArray(8, 1) = " "
    aArray(8, 2) = "���� "
    aArray(8, 3) = "������"
    aArray(8, 4) = "������� "
    aArray(8, 5) = "������� "
    aArray(8, 6) = "������� "
    aArray(8, 7) = "������ "
    aArray(8, 8) = "��������� "
    aArray(8, 9) = "������� "
    aArray(8, 10) = "�������� "
    
    aArray(9, 1) = " "
    aArray(9, 2) = "��� "
    aArray(9, 3) = "��� "
    aArray(9, 4) = "���� "
    aArray(9, 5) = "������� "
    aArray(9, 6) = "����� "
    aArray(9, 7) = "��� "
    aArray(9, 8) = "���� "
    aArray(9, 9) = "���� "
    aArray(9, 10) = "����� "
    
    For intLoop = 1 To 14
        If Mid(tmpOldNumber, intLoop, 1) <> "." Then
            tmpNumber = tmpNumber + Mid(tmpOldNumber, intLoop, 1)
        End If
    Next intLoop
    
    tmpIntNumber = Int(Val(tmpNumber))
    
    For intLoop = 1 To 9 - Len(Trim(tmpIntNumber))
        strTotalGross = strTotalGross + "0"
    Next intLoop
    strTotalGross = strTotalGross + Trim(tmpNumber)

    For intLoop = 1 To 9
        strSubNumber = Mid(strTotalGross, intLoop, 1)
        aFullNumber(intLoop) = aArray(bytArrayIndex, Val(strSubNumber) + 1)
        bytArrayIndex = bytArrayIndex + 1
    Next intLoop
    
    '�����������
    If aFullNumber(1) <> " " Or aFullNumber(2) <> " " Or aFullNumber(3) <> " " Then
        If aFullNumber(2) = "���� " Then
            If aFullNumber(3) = "��� " Then
                aFullNumber(2) = ""
                aFullNumber(3) = "������ "
            End If
            If aFullNumber(3) = "��� " Then
                aFullNumber(2) = ""
                aFullNumber(3) = "������ "
            End If
        End If
    End If
    
    '��������
    If aFullNumber(4) <> " " Or aFullNumber(5) <> " " Or aFullNumber(6) <> " " Then
        If aFullNumber(5) = "���� " Then
            If aFullNumber(6) = "��� " Then
                aFullNumber(5) = ""
                aFullNumber(6) = "������ "
            End If
            If aFullNumber(6) = "��� " Then
                aFullNumber(5) = ""
                aFullNumber(6) = "������ "
            End If
        End If
    End If
    
    '�����������
    If aFullNumber(7) <> " " Or aFullNumber(8) <> " " Or aFullNumber(9) <> " " Then
        If aFullNumber(8) = "���� " Then
            If aFullNumber(9) = "��� " Then
                aFullNumber(8) = ""
                aFullNumber(9) = "������ "
            End If
            If aFullNumber(9) = "��� " Then
                aFullNumber(8) = ""
                aFullNumber(9) = "������ "
            End If
        End If
    End If
    
    '�����������
    If aFullNumber(1) <> " " Or aFullNumber(2) <> " " Or aFullNumber(3) <> " " Then
        If aFullNumber(1) = "������ " And aFullNumber(2) = " " And aFullNumber(3) = " " Then
            aFullNumber(1) = "����� "
        End If
        If aFullNumber(1) = " " And aFullNumber(2) = " " And aFullNumber(3) = "��� " Then
            aFullNumber(3) = aFullNumber(3) + "����������� "
        Else
            aFullNumber(3) = aFullNumber(3) + "����������� "
        End If
    End If
    
    '��������
    If aFullNumber(4) <> " " Or aFullNumber(5) <> " " Or aFullNumber(6) <> " " Then
        If aFullNumber(4) = "������ " And aFullNumber(5) = " " And aFullNumber(6) = " " Then
            aFullNumber(4) = "����� "
        End If
        If aFullNumber(4) = " " And aFullNumber(5) = " " And aFullNumber(6) = "��� " Then
            aFullNumber(6) = "����� "
        End If
        If aFullNumber(6) <> "����� " Then
            aFullNumber(6) = aFullNumber(6) + "�������� "
        End If
    End If
    
    '�����������
    If aFullNumber(7) = "������ " And aFullNumber(8) = " " And aFullNumber(9) = " " Then
        aFullNumber(7) = "����� "
    End If
    
    For intLoop = 1 To 9
        If Trim(aFullNumber(intLoop)) <> "" Then
            strFullNumber = strFullNumber + aFullNumber(intLoop)
        End If
    Next intLoop
    
    If strFullNumber = "" Then strFullNumber = "����� "
    strFullNumber = strFullNumber + "���� "
    
    bytArrayIndex = 8
    tmpDecNumber = Mid(strTotalGross, 11, 2)
     
    If tmpDecNumber = "00" Then
        FullNumber = strFullNumber
        Exit Function
    End If
        
    strFullNumber = IIf(strFullNumber <> "����� ����", strFullNumber + "��� ", "")
    
    For intLoop = 1 To 2
        strSubNumber = Mid(tmpDecNumber, intLoop, 1)
        aFullNumber(intLoop) = aArray(bytArrayIndex, Val(strSubNumber) + 1)
        bytArrayIndex = bytArrayIndex + 1
    Next intLoop
    
    If aFullNumber(1) <> " " Or aFullNumber(2) <> " " Then
        If aFullNumber(1) = "���� " Then
            If aFullNumber(2) = "��� " Then
                aFullNumber(1) = " "
                aFullNumber(2) = "������ "
            End If
            If aFullNumber(2) = "��� " Then
                aFullNumber(1) = " "
                aFullNumber(2) = "������ "
            End If
        End If
    End If
    
    For intLoop = 1 To 2
        If Len(Trim(aFullNumber(intLoop))) <> 0 Then
            strFullNumber = strFullNumber + aFullNumber(intLoop)
        End If
    Next intLoop
    
    If tmpDecNumber = "01" Then
        strFullNumber = strFullNumber + "����� "
    Else
        strFullNumber = strFullNumber + "����� "
    End If
            
    FullNumber = strFullNumber

End Function

Function WindowTitle(title)

    Select Case title
        Case "������"
            WindowTitle = "���������� ������"
        Case "��������"
            WindowTitle = "���������� ��������"
        Case "�����������", "�������� �����������"
            WindowTitle = "���������� �������� �����������"
        Case "�������", "�������� �������"
            WindowTitle = "���������� �������� �������"
        Case "����", "�������� �����"
            WindowTitle = "���������� �������� �����"
    End Select

End Function

