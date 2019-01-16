VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptInvoice 
   ClientHeight    =   15180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   24960
   Icon            =   "rptInvoice.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   44027
   _ExtentY        =   26776
   SectionData     =   "rptInvoice.dsx":1CCA
End
Attribute VB_Name = "rptInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpRecordset As Recordset

Private Sub ActiveReport_DataInitialize()

    Set tmpRecordset = SeekInvoiceData
    
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
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If tmpRecordset.EOF Then
        EOF = True
        Exit Sub
    End If
    
    With tmpRecordset
    
        Fields("CompanyTitle") = arrCompanyData(1)
        Fields("CompanyData") = arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6) & Chr(13)
        
        Fields("CodeDescription") = !CodeDescription
        Fields("CodeBatch") = !CodeBatch
        Fields("InvoiceNo") = !InvoiceNo
        Fields("InvoiceIssueDate") = !InvoiceIssueDate
        
        Fields("InvoiceRemarks") = !InvoiceRemarks
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
        
    End With
    
    EOF = False
    
    tmpRecordset.MoveNext
    
End Sub

Private Function SeekInvoiceData()

    Dim strSQL As String
    
    Dim rstRecordset As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strSQL = "SELECT " _
        & "Description, Profession, Address, City, TaxNo, Phones, " _
        & "CodeDescription, CodeBatch, " _
        & "TaxOfficeDescription, " _
        & "ItemDescription, " _
        & "ManufacturerDescription, " _
        & "InvoiceIssueDate, InvoiceNo, InvoiceNet, InvoiceAmountDiscount, InvoiceRestAmount, InvoiceVATAmount, InvoiceGrossAmount, InvoiceTransportReason, InvoiceTransportWay, InvoiceLoadingSite, InvoiceDestinationSite, InvoicePlates, InvoiceRemarks, " _
        & "Qty , UnitPrice, TotalNetPreDiscount, DiscPercent, DiscAmount, TotalNetPostDiscount, VATPercent, VATAmount, TotalGross " _
        & "FROM ((((((InvoicesTrn " _
        & "INNER JOIN Invoices ON InvoicesTrn.InvoiceTrnID = Invoices.InvoiceTrnID) " _
        & "INNER JOIN Items ON InvoicesTrn.ItemID = Items.ItemID) " _
        & "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN TaxOffices ON Customers.TaxOfficeID = TaxOffices.TaxOfficeID) " _
        & "INNER JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID) " _
        & "WHERE InvoicesTrn.InvoiceTrnID = 46535"
        
    TempQuery.SQL = strSQL
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    Set SeekInvoiceData = rstRecordset

End Function


