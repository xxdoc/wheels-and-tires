VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form UtilsUpdateItemQty 
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
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   3150
      TabIndex        =   5
      Top             =   4050
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "UtilsUpdateItemQty.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "UtilsUpdateItemQty.frx":001C
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   7
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   300
      TabIndex        =   0
      Top             =   7650
      Width           =   4665
      Begin GurhanButtonOCX.GurhanButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
      Caption         =   "Ενημέρωση ποσοτήτων"
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
      TabIndex        =   1
      Top             =   75
      Width           =   5445
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
Attribute VB_Name = "UtilsUpdateItemQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnProcessing As Boolean
Dim blnErrorFound As Boolean
Dim blnProcessAborted As Boolean

Private Function UpdateQties()

    On Error GoTo ErrTrap
    
    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    'Recordsets
    Dim rsItems As Recordset
    Dim rstRecordset As Recordset

    'Local μεταβλητές
    Dim lngItemID As Long
    Dim lngQty As Long
    
    'Αρχικές τιμές
    lngQty = 0
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Είδη
    Set rsItems = CommonDB.OpenRecordset("items")
    rsItems.Index = "ID"
    
    
    'Προσωρινά
    blnProcessing = True
    
    'Εγγραφές
    TempQuery.SQL = "SELECT Items.ItemID, Items.ItemDescription, InvoicesTrn.Qty, Codes.CodeInventoryQty, InvoicesTrn.InvoiceTrnID " _
                                    & " FROM ((Items " _
                                    & " INNER JOIN InvoicesTrn ON Items.ItemID = InvoicesTrn.ItemID) " _
                                    & " INNER JOIN Invoices ON InvoicesTrn.InvoiceTrnID = Invoices.InvoiceTrnID) " _
                                    & " INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID WHERE (((Codes.CodeInventoryQty) = '+' Or (Codes.CodeInventoryQty) = '-')) ORDER BY Items.ItemID, InvoicesTrn.InvoiceTrnID"
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω είδη, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnErrorFound = False: UpdateQties = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Αρχικές τιμές
    lngItemID = rstRecordset!ItemID
    
    Do While Not rstRecordset.EOF
        While lngItemID = rstRecordset!ItemID
            lngQty = IIf(rstRecordset!CodeInventoryQty = "+", lngQty + rstRecordset!Qty, lngQty - rstRecordset!Qty)
            rstRecordset.MoveNext
            UpdateProgressBar Me
            If rstRecordset.EOF Then Exit Do
        Wend
        GoSub UpdateItemWithQtyBalance
        lngQty = 0
        lngItemID = rstRecordset!ItemID
    Loop
    
    GoSub UpdateItemWithQtyBalance
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
    Else
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    rstRecordset.Close
    frmProgress.Visible = False
    
    Exit Function
    
UpdateItemWithQtyBalance:
    rsItems.Seek "=", lngItemID
    If Not rsItems.NoMatch Then
        If rsItems!ItemBalance <> lngQty Then
            rsItems.Edit
            rsItems!ItemBalance = lngQty
            rsItems.Update
            UpdateLogFile "Table Items ID = " & lngItemID & " Έγινε ενημέρωση ποσότητας"
        End If
    End If
    
    Return
    
UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return
    
ErrTrap:
    ClearFields frmProgress
    blnProcessAborted = True
    cmdButton(1).Caption = "Νέα αναζήτηση"
    DisplayErrorMessage True, Err.Description

End Function

Private Function StartProcess()

    On Error GoTo ErrTrap
    
    blnErrorFound = False
    blnProcessing = True
    blnProcessAborted = False
    
    UpdateButtons Me, 2, 0, 1, 0
    
    If blnProcessing Then UpdateQties
    
    blnProcessing = False
    
    UpdateButtons Me, 2, 1, 0, 1
    
    If Not blnProcessAborted Then
        DisplayMessage 10, 1, 1, ""
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
        If DisplayMessage(33, 2, 2, "", "") Then
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

