VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptInvoiceA 
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   Icon            =   "rptInvoiceA.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   33443
   _ExtentY        =   18045
   SectionData     =   "rptInvoiceA.dsx":1CCA
End
Attribute VB_Name = "rptInvoiceA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpRecordset As Recordset
Dim blnError As Boolean
Dim strInvoiceRemarks As String

Private Sub ActiveReport_DataInitialize()

    Set tmpRecordset = SeekInvoiceData
    
    If tmpRecordset.RecordCount = 0 Then
        GoTo ErrTrap
    End If
    
    Fields.RemoveAll
    
    Fields.Add ("CompanyTitle")
    Fields.Add ("CompanyData")
    
    Fields.Add "CodeDescription"
    Fields.Add "CodeBatch"
    Fields.Add "InvoiceNo"
    Fields.Add "InvoiceIssueDate"
    
    Fields.Add "InvoiceRemarks"
    Fields.Add "InvoiceTransportReason"
    Fields.Add "InvoiceTransportWay"
    Fields.Add "InvoiceLoadingSite"
    Fields.Add "InvoiceDestinationSite"
    Fields.Add "InvoicePlates"
        
    Fields.Add "ID"
    Fields.Add "Description"
    Fields.Add "Profession"
    Fields.Add "Address"
    Fields.Add "TaxNo"
    Fields.Add "Phones"
    Fields.Add "TaxOfficeDescription"
    
    Fields.Add "ItemDescription"
    Fields.Add "ManufacturerDescription"
    Fields.Add "ItemUnitOfMeasurement"
    Fields.Add "Qty"
    Fields.Add "UnitPrice"
    Fields.Add "TotalNetPreDiscount"
    Fields.Add "DiscPercent"
    Fields.Add "DiscAmount"
    Fields.Add "TotalNetPostDiscount"
    Fields.Add "VATPercent"
    Fields.Add "VATAmount"
    Fields.Add "TotalGross"
    
    Fields.Add "PerVATNetAmount"
    Fields.Add "PerVATPercent"
    Fields.Add "PerVATAmount"
    
    Fields.Add "InvoiceRestAmount"
    Fields.Add "InvoiceVATAmount"
    Fields.Add "InvoiceGrossAmount"
    
    Fields.Add "PaymentWayDescription"
    Fields.Add "BankAccountNumber"
    
    Fields.Add "NumberInWords"
    
    Exit Sub
    
ErrTrap:
    blnError = True
    DisplayErrorMessage True, strMessages(8)
    Unload Me
        
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    On Error GoTo ErrTrap
    
    If tmpRecordset.EOF Then
        EOF = True
        Exit Sub
    Else
        strInvoiceRemarks = tmpRecordset!InvoiceRemarks
    End If
    
    With tmpRecordset
    
        Fields("CompanyTitle") = arrCompanyData(1)
        Fields("CompanyData") = arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6) & Chr(13)
        
        Fields("CodeDescription") = !CodeDescription
        Fields("CodeBatch") = IIf(!CodeBatch <> "", "сеияа " & !CodeBatch, "")
        Fields("InvoiceNo") = "мО " & !InvoiceNo
        Fields("InvoiceIssueDate") = !InvoiceIssueDate
        
        Fields("InvoiceTransportReason") = !InvoiceTransportReason
        Fields("InvoiceTransportWay") = !InvoiceTransportWay
        Fields("InvoiceLoadingSite") = !InvoiceLoadingSite
        Fields("InvoiceDestinationSite") = !InvoiceDestinationSite
        Fields("InvoicePlates") = !InvoicePlates
        
        Fields("Description") = !Description
        Fields("Profession") = !Profession
        Fields("Address") = !Address + " " + !City
        Fields("TaxNo") = !TaxNo
        Fields("Phones") = !Phones
        Fields("TaxOfficeDescription") = !TaxOfficeDescription
        
        Fields("ItemDescription") = !ItemDescription
        Fields("ManufacturerDescription") = !ManufacturerDescription
        Fields("ItemUnitOfMeasurement") = "TEM"
        Fields("Qty") = !Qty
        Fields("UnitPrice") = !UnitPrice
        Fields("TotalNetPreDiscount") = !TotalNetPreDiscount
        Fields("DiscPercent") = !DiscPercent
        Fields("DiscAmount") = !DiscAmount
        Fields("TotalNetPostDiscount") = !TotalNetPostDiscount
        Fields("VATPercent") = !VATPercent
        Fields("VATAmount") = !VATAmount
        Fields("TotalGross") = !TotalGross
        
        Fields("PerVATNetAmount") = !InvoiceRestAmount
        Fields("PerVATPercent") = "24"
        Fields("PerVATAmount") = !InvoiceVATAmount

        Fields("InvoiceRestAmount") = !InvoiceRestAmount
        Fields("InvoiceVATAmount") = !InvoiceVATAmount
        Fields("InvoiceGrossAmount") = !InvoiceGrossAmount
        
        Fields("PaymentWayDescription") = !PaymentWayDescription
        Fields("BankAccountNumber") = strBankAccountNo
        
        Fields("NumberInWords") = FullNumber(Format(Fields("InvoiceGrossAmount"), "#,##0.00")) + "   "
        
    End With
    
    EOF = False
    
    tmpRecordset.MoveNext
    
    Exit Sub
    
ErrTrap:
    If Err.Number = 6 Then
        Resume Next
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
End Sub

Private Function SeekInvoiceData()

    Dim strSQL As String
    
    Dim rstRecordset As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strSQL = "SELECT " _
        & "Description, Profession, Address, City, TaxNo, Phones, " _
        & "CodeDescription, CodeBatch, " _
        & "TaxOfficeDescription, " _
        & "PaymentWayDescription, " _
        & "ItemDescription, " _
        & "ManufacturerDescription, " _
        & "InvoiceIssueDate, InvoiceNo, InvoiceNet, InvoiceAmountDiscount, InvoiceRestAmount, InvoiceVATAmount, InvoiceGrossAmount, InvoiceTransportReason, InvoiceTransportWay, InvoiceLoadingSite, InvoiceDestinationSite, InvoicePlates, InvoiceRemarks, " _
        & "Qty , UnitPrice, TotalNetPreDiscount, DiscPercent, DiscAmount, TotalNetPostDiscount, VATPercent, VATAmount, TotalGross " _
        & "FROM (((((((InvoicesTrn " _
        & "INNER JOIN Invoices ON InvoicesTrn.InvoiceTrnID = Invoices.InvoiceTrnID) " _
        & "INNER JOIN Items ON InvoicesTrn.ItemID = Items.ItemID) " _
        & "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN PaymentWays ON Invoices.InvoicePaymentWayID = PaymentWays.PaymentWayID) " _
        & "INNER JOIN TaxOffices ON Customers.TaxOfficeID = TaxOffices.TaxOfficeID) " _
        & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) " _
        & "WHERE InvoicesTrn.InvoiceTrnID = " & CLng(Me.Tag)
        
    TempQuery.SQL = strSQL
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    Set SeekInvoiceData = rstRecordset
    
End Function

Private Function CalculateTotalPages()

    Dim curNumber As Currency
    Dim curResult As Currency
    Dim intInteger As Integer
    Dim intPages As Integer
    
    tmpRecordset.MoveLast
    curNumber = tmpRecordset.RecordCount / 8
    tmpRecordset.MoveFirst
    curResult = curNumber - Int(curNumber)
    
    intInteger = Int(curNumber)
    
    If curResult <> 0 Then
        intPages = Int(curNumber) + 1
    Else
        intPages = intInteger
    End If
    
    CalculateTotalPages = IIf(intPages = 0, 1, intPages)
    
End Function

Private Sub PageFooter_Format()

    If Not tmpRecordset.EOF Then
        ToggleFieldVisibility False, PaymentWayDescription, BankAccountNumber, PerVATNetAmount, PerVATPercent, PerVATAmount, InvoiceRestAmount, InvoiceVATAmount, InvoiceGrossAmount, NumberInWords
        lblRemarks.Caption = "то паяастатийо сумевифетаи..."
    Else
        ToggleFieldVisibility True, PaymentWayDescription, BankAccountNumber, PerVATNetAmount, PerVATPercent, PerVATAmount, InvoiceRestAmount, InvoiceVATAmount, InvoiceGrossAmount, NumberInWords
        lblRemarks.Caption = strInvoiceRemarks
    End If

End Sub

