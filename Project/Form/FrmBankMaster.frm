VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmBankMaster 
   BorderStyle     =   0  'None
   Caption         =   "BANK MASTER"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtRemark 
      DataField       =   "Oprice"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox TxtMgr 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox TxtRef 
      DataField       =   "Oprice"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Width           =   3495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxBSlab 
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   5
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox TxtBrName 
      DataField       =   "Oprice"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox TxtBName 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox TxtBankCode 
      DataField       =   "Model"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Reference"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Manager Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Branch Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Bank Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmBankMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BFlg As String
Dim BCN As New ADODB.Connection
Dim BSTR As String
Dim rs As New ADODB.Recordset
Dim DT As Date


Private Sub CmdAdd_Click()
  Call BlankText(Me)
  Flg = "ADD"
  TxtBName.SetFocus
End Sub
Private Sub CmdCancel_Click()
 Call BlankText(Me)
 FlxBSlab.Clear
 FlxBSlab.Rows = 2
 
 Call HEADING
  
End Sub
Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub CmdSave_Click()

If Flg = "ADD" Then
    BSTR = "INSERT INTO bankMASTER VALUES ('" & TxtBankCode.Text & "','" & TxtBName.Text & "',  '" & TxtBrName.Text & "',  '" & TxtMgr.Text & "', '" & TxtRef.Text & "',' " & TxtRemark.Text & "','" & UN & "','" & DT & "')"
    BCN = GETCON
    BCN.Open
    BCN.BeginTrans
    BCN.Execute BSTR
    
    For I = 1 To FlxBSlab.Rows - 1
           If FlxBSlab.TextMatrix(I, 1) <> "" And FlxBSlab.TextMatrix(I, 2) <> "" And FlxBSlab.TextMatrix(I, 3) <> "" And FlxBSlab.TextMatrix(I, 4) <> "" Then
            BSTR = "INSERT INTO bankslabmaster VALUES ('" & TxtBankCode.Text & "'," & FlxBSlab.TextMatrix(I, 0) & ", " & FlxBSlab.TextMatrix(I, 1) & "," & FlxBSlab.TextMatrix(I, 2) & "," & FlxBSlab.TextMatrix(I, 3) & "," & FlxBSlab.TextMatrix(I, 4) & ")"
            BCN.Execute BSTR
           End If
    Next
    BCN.CommitTrans
    MsgBox "RECORD SAVE SUCCESSSFULLY "
    Flg = ""
End If

End Sub

Private Sub FlxBSlab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        FlxBSlab.TextMatrix(FlxBSlab.Row, FlxBSlab.Col) = ""
    Else
        If KeyAscii = 13 And FlxBSlab.Col = 4 Then

            FlxBSlab.Rows = FlxBSlab.Rows + 1
            FlxBSlab.TextMatrix(FlxBSlab.Row + 1, 0) = Val(FlxBSlab.TextMatrix(FlxBSlab.Row, 0)) + 1
        '   FlxBSlab.RowPosition
            FlxBSlab.Row = FlxBSlab.Row + 1
            FlxBSlab.Col = 1
             
        Else
            FlxBSlab = FlxBSlab + Chr(KeyAscii)
        End If
     End If
End Sub

Private Sub Form_Load()

 Me.Top = 250
 Me.Left = 200
 Call HEADING
DT = Date
End Sub
Private Sub HEADING()
FlxBSlab.TextMatrix(0, 0) = "SrNo"
FlxBSlab.TextMatrix(0, 1) = "From Slab"
FlxBSlab.TextMatrix(0, 2) = "To Slab"
FlxBSlab.TextMatrix(0, 3) = "Amount"
FlxBSlab.TextMatrix(0, 4) = "Percentage"

End Sub

Private Sub TxtBankCode_(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub TxtBName_LostFocus()
 If TxtBName.Text <> "" And Flg = "ADD" Then
  Set rs = New ADODB.Recordset
        rs.Open "SELECT COUNT(bCODE) FROM bankMASTER WHERE LEFT(bCODE,3) = '" & Left(TxtBName.Text, 3) & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = True And rs.BOF = True Then
            TxtBankCode.Text = UCase(Left(TxtBName.Text, 3)) & Format(1, "0000")
        Else
            TxtBankCode.Text = UCase(Left(TxtBName.Text, 3)) & Format(rs(0) + 1, "0000")
        End If
        rs.Close
    End If
End Sub

Private Sub TxtBrName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtBrName_LostFocus()
FlxBSlab.TextMatrix(1, 0) = 1
End Sub

Private Sub TxtBname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtMgr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
