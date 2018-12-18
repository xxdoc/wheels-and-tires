VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form UtilsTablesCheck 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12300
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   300
      TabIndex        =   3
      Top             =   7650
      Width           =   4665
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Συνέχεια"
         ButtonStyle     =   4
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ShowFocusRect   =   0   'False
         BackColor       =   8438015
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Ακυρο"
         ButtonStyle     =   4
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ShowFocusRect   =   0   'False
         BackColor       =   8438015
      End
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   3075
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         Caption         =   "Κλείσιμο"
         ButtonStyle     =   4
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         ShowFocusRect   =   0   'False
         BackColor       =   8421631
      End
   End
   Begin VB.Frame frmProgress 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1290
      Left            =   4875
      TabIndex        =   0
      Top             =   4650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   465
         Left            =   150
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   820
         Picture         =   "UtilsTablesCheck.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "UtilsTablesCheck.frx":001C
         BarPictureMode  =   0
         BackPictureMode =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "1 από 99999"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   900
         Width           =   3765
      End
      Begin VB.Label lblMaster 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Τίτλος"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   10950
      Top             =   3000
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   1575
      Top             =   8325
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ελεγχος αρχείων"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   720
      Left            =   225
      TabIndex        =   4
      Top             =   75
      Width           =   3870
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   0
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "UtilsTablesCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnProcessing As Boolean
Dim blnErrorFound As Boolean
Dim blnWarningFound As Boolean
Dim blnProcessAborted As Boolean

Private Function CheckBanks()

    CheckTable "Banks", "SELECT Banks.BankID, Banks.BankDescription FROM Banks LEFT JOIN Checks ON Banks.BankID = Checks.CheckBankID WHERE Checks.CheckBankID Is Null", "η τράπεζα δεν χρησιμοποιείται. Μπορεί να διαγραφεί", "Warning"

End Function

Private Function CheckCategories()

    CheckTable "Categories", "SELECT Categories.CategoryID, Categories.CategoryDescription FROM Categories LEFT JOIN Items ON Categories.CategoryID = Items.ItemCategoryID WHERE Items.ItemCategoryID Is Null", "η κατηγορία δεν χρησιμοποιείται. Μπορεί να διαγραφεί", "Warning"
    
End Function

Private Function CheckChecks()

    CheckTable "Checks", "SELECT CheckID, CheckExpireDate FROM Checks WHERE Year(CheckExpireDate) < 2000", "η ημερομηνία είναι λάθος", False
    CheckTable "Checks", "SELECT CheckID, CheckAmount FROM Checks WHERE CheckAmount < 0", "το ποσό είναι λάθος", False
    CheckTable "Checks", "SELECT Checks.CheckID, Checks.CheckBankID FROM Checks LEFT JOIN Banks ON Checks.CheckBankID = Banks.BankID WHERE Banks.BankID Is Null", "η τράπεζα δεν βρέθηκε", "Error"
    CheckTable "Checks", "SELECT CheckID, CheckIssuedByID FROM Checks LEFT JOIN Customers ON Checks.CheckIssuedByID = Customers.ID WHERE (CheckRefersToID = 3 AND Checks.CheckIssuedByID <> 0) AND Customers.ID Is Null", "ο πελάτης δεν βρέθηκε", "Error"
    CheckTable "Checks", "SELECT CheckID, CheckIssuedByID FROM Checks LEFT JOIN Suppliers ON Checks.CheckIssuedByID = Suppliers.ID WHERE (CheckRefersToID = 4 AND Checks.CheckIssuedByID <> 0) AND Suppliers.ID Is Null", "ο προμηθευτής δεν βρέθηκε", "Error"
    
End Function

Private Function CheckCustomers()

    CheckTable "Customers", "SELECT Customers.ID, Customers.TaxOfficeID FROM Customers LEFT JOIN TaxOffices ON Customers.TaxOfficeID = TaxOffices.TaxOfficeID WHERE TaxOffices.TaxOfficeID Is Null", "η οικονομκή υπηρεσία δεν βρέθηκε", "Error"
    CheckTable "Customers", "SELECT Customers.ID, Customers.CountryID FROM Customers LEFT JOIN Countries ON Customers.CountryID = Countries.CountryID WHERE Countries.CountryID Is Null", "η χώρα δεν βρέθηκε", "Error"
    CheckTable "Customers", "SELECT ID, Description FROM Customers WHERE NOT EXISTS (SELECT InvoicePersonID FROM Invoices WHERE Invoices.InvoicePersonID = Customers.ID AND (Invoices.InvoiceRefersToID = 2 OR Invoices.InvoiceRefersToID = 4))", "ο πελάτης δεν έχει κινηθεί και μπορεί να διαγραφεί", "Warning"

End Function

Private Function CheckInvoices()

    CheckTable "Invoices", "SELECT InvoiceID, InvoiceIssueDate FROM Invoices WHERE Year(InvoiceIssueDate) < 2000", "η ημερομηνία είναι λάθος", "Error"
    CheckTable "Invoices", "SELECT InvoiceID, InvoiceNo FROM Invoices WHERE InvoiceNo Is Null", "το νούμερο παραστατικού είναι λάθος", "Error"
    CheckTable "Invoices", "SELECT InvoiceID, InvoiceCodeID FROM Invoices LEFT JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID WHERE Codes.CodeID Is Null", "το παραστατικό δεν βρέθηκε", "Error"
    CheckTable "Invoices", "SELECT InvoiceID, InvoicePaymentWayID FROM Invoices LEFT JOIN PaymentWays ON Invoices.InvoicePaymentWayID = PaymentWays.PaymentWayID WHERE PaymentWays.PaymentWayID Is Null", "ο τρόπος πληρωμής δεν βρέθηκε", "Error"
    CheckTable "Invoices", "SELECT InvoiceID, InvoicePersonID FROM Invoices LEFT JOIN Suppliers ON Invoices.InvoicePersonID = Suppliers.ID WHERE ((InvoiceRefersToID = 1 OR InvoiceRefersToID = 3) AND Invoices.InvoicePersonID <> 0) AND Suppliers.ID Is Null", "ο προμηθευτής δεν βρέθηκε", "Error"
    CheckTable "Invoices", "SELECT InvoiceID, InvoicePersonID FROM Invoices LEFT JOIN Customers ON Invoices.InvoicePersonID = Customers.ID WHERE ((InvoiceRefersToID = 2 OR InvoiceRefersToID = 4) AND Invoices.InvoicePersonID <> 0) AND Customers.ID Is Null", "ο πελάτης δεν βρέθηκε", "Error"
    CheckTable "Invoices", "SELECT InvoiceID, InvoiceDeliveryPointID FROM Invoices LEFT JOIN DeliveryPoints ON Invoices.InvoiceDeliveryPointID = DeliveryPoints.DeliveryPointID WHERE DeliveryPoints.DeliveryPointID Is Null", "το σημείο παραλαβής δεν βρέθηκε", "Error"

End Function

Private Function CheckInvoicesTrn()

    CheckTable "InvoicesTrn", "SELECT ID, InvoicesTrn.ItemID FROM InvoicesTrn LEFT JOIN Items ON InvoicesTrn.ItemID = Items.ItemID WHERE Items.ItemID Is Null", "το είδος δεν βρέθηκε", "Error"
    CheckTable "InvoicesTrn", "SELECT ID, InvoicesTrn.InvoiceTrnID FROM InvoicesTrn LEFT JOIN Invoices ON InvoicesTrn.InvoiceTrnID = Invoices.InvoiceTrnID WHERE Invoices.InvoiceTrnID Is Null", "το βασικό παραστατικό δεν βρέθηκε", "Error"
    
End Function

Private Function CheckItems()

    CheckTable "Items", "SELECT ItemID, ItemCategoryID FROM Items LEFT JOIN Categories ON Items.ItemCategoryID = Categories.CategoryID WHERE Categories.CategoryID Is Null", "η κατηγορία δεν βρέθηκε", "Error"
    CheckTable "Items", "SELECT ItemID, ItemManufacturerID FROM Items LEFT JOIN Manufacturers ON Items.ItemManufacturerID = Manufacturers.ManufacturerID WHERE Manufacturers.ManufacturerID Is Null", "ο κατασκευαστής δεν βρέθηκε", "Error"
    CheckTable "Items", "SELECT ItemID, ItemDescription FROM Items WHERE NOT EXISTS (SELECT ItemID FROM InvoicesTrn WHERE InvoicesTrn.ItemID = Items.ItemID)", "το είδος δεν έχει κινηθεί και μπορεί να διαγραφεί", "Warning"

End Function

Private Function CheckSuppliers()

    CheckTable "Suppliers", "SELECT Suppliers.ID, Suppliers.TaxOfficeID FROM Suppliers LEFT JOIN TaxOffices ON Suppliers.TaxOfficeID = TaxOffices.TaxOfficeID WHERE TaxOffices.TaxOfficeID Is Null", "η οικονομκή υπηρεσία δεν βρέθηκε", "Error"
    CheckTable "Suppliers", "SELECT Suppliers.ID, Suppliers.CountryID FROM Suppliers LEFT JOIN Countries ON Suppliers.CountryID = Countries.CountryID WHERE Countries.CountryID Is Null", "η χώρα δεν βρέθηκε", "Error"
    CheckTable "Suppliers", "SELECT ID, Description FROM Suppliers WHERE NOT EXISTS (SELECT InvoicePersonID FROM Invoices WHERE Invoices.InvoicePersonID = Suppliers.ID AND (Invoices.InvoiceRefersToID = 1 OR Invoices.InvoiceRefersToID = 3))", "ο προμηθευτής δεν έχει κινηθεί και μπορεί να διαγραφεί", "Warning"

End Function

Private Function CheckTable(myTable, mySQL, myMessage, myErrorOrWarning)

    Dim lngRecordCount As Long
    Dim rstRecordset As Recordset
    
    Set rstRecordset = CommonDB.OpenRecordset(mySQL)
    
    Do While Not rstRecordset.EOF
        UpdateLogFile "Table " & myTable & " ID = " & rstRecordset.Fields(0).Value & " " & rstRecordset.Fields(1).Value & " " & myMessage
        If myErrorOrWarning = "Error" And Not blnErrorFound Then blnErrorFound = True
        If myErrorOrWarning = "Warning" And Not blnWarningFound Then blnWarningFound = True
        rstRecordset.MoveNext
        DoEvents
        If Not blnProcessing Then blnErrorFound = False: Exit Do
    Loop
    
    frmProgress.Visible = False

End Function

Private Function CheckTablesForNullValues()

    Dim strSQL As String
    Dim intLoop As Integer
    Dim intFieldCount As Integer
    Dim intNoOfTables As Integer
    Dim strTables
    Dim rstTemp As Recordset
    Dim lngRecordCount As Long
    
    strTables = Array("Banks", "Categories", "Checks", "Codes", "Countries", "Customers", "DeliveryPoints", "Inventory", "Invoices", "InvoicesTrn", "Items", "Manufacturers", "Options", "PaymentWays", "Settings", "Suppliers", "TaxOffices", "VATStates", "YesOrNo")

    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Ελέγχω τον κάθε πίνακα για κενά πεδία
    For intLoop = 0 To UBound(strTables)
        TempQuery.SQL = "SELECT * FROM " & strTables(intLoop)
        Set rstTemp = TempQuery.OpenRecordset()
        Do While Not rstTemp.EOF
            For intFieldCount = 0 To rstTemp.Fields.Count - 1
                If IsNull(rstTemp.Fields(intFieldCount)) Then
                    UpdateLogFile "Table: " & strTables(intLoop) & " Field: " & rstTemp.Fields(intFieldCount).Name & " | Rec ID: " & rstTemp.Fields(0).Value & " Field is NULL"
                    blnErrorFound = True
                End If
            Next intFieldCount
            rstTemp.MoveNext
            DoEvents
            If Not blnProcessing Then blnErrorFound = False: Exit For
        Loop
    Next intLoop
    
    'Τέλος
    frmProgress.Visible = False
    
End Function

Private Function StartProcess()

    On Error GoTo ErrTrap
    
    blnErrorFound = False
    blnWarningFound = False
    blnProcessing = True
    blnProcessAborted = False
    
    UpdateButtons Me, 2, 0, 1, 0
    
    If blnProcessing Then CheckTablesForNullValues
    If blnProcessing Then CheckBanks
    If blnProcessing Then CheckCategories
    If blnProcessing Then CheckChecks
    If blnProcessing Then CheckCustomers
    If blnProcessing Then CheckInvoices
    If blnProcessing Then CheckInvoicesTrn
    If blnProcessing Then CheckItems
    If blnProcessing Then CheckSuppliers
    
    blnProcessing = False
    
    UpdateButtons Me, 2, 1, 0, 1
    
    If Not blnProcessAborted Then
        If blnErrorFound Then
            DisplayMessage 35, 3, 1, ""
        Else
            If blnWarningFound Then
                DisplayMessage 38, 3, 1, ""
            Else
                DisplayMessage 34, 1, 1, ""
            End If
        End If
    End If
    
    Exit Function
    
ErrTrap:
    UpdateButtons Me, 2, 1, 0, 1
    blnProcessing = False
    frmProgress.Visible = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            StartProcess
        Case 1
            AbortProcedure False
        Case 2
            AbortProcedure True
    End Select
    
End Sub

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then
        'Διακοπή;
        If DisplayMessage(33, 2, 2, "") Then
            blnProcessing = False
            blnProcessAborted = True
        End If
    End If

    If blnStatus Then
        Unload Me
    End If

End Function

Private Sub Form_Activate()
                                                                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        Me.Refresh
    End If
            
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)

    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyEscape
            If cmdButton(1).Enabled Then cmdButton_Click 1: Exit Function
            If cmdButton(2).Enabled Then cmdButton_Click 2
    End Select

End Function

Private Sub Form_Load()

    PositionControls Me, False: ColorizeControls Me
    UpdateButtons Me, 2, 1, 0, 1

End Sub

