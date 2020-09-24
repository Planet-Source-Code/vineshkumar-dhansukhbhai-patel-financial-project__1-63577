VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmVehMaster 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   435
      Left            =   13440
      TabIndex        =   17
      Top             =   8370
      Width           =   585
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   435
      Left            =   14070
      TabIndex        =   16
      Top             =   8370
      Width           =   585
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cl&ose"
      Height          =   345
      Left            =   13590
      TabIndex        =   4
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   12390
      TabIndex        =   3
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   11190
      TabIndex        =   2
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   9990
      TabIndex        =   1
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   345
      Left            =   8790
      TabIndex        =   0
      Top             =   9570
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   -30
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsItem 
      Height          =   4785
      Left            =   1260
      TabIndex        =   10
      Top             =   3480
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   8440
      _Version        =   393216
      Cols            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.TextBox TxtOPrize 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Oprice"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3540
      TabIndex        =   9
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox TxtPrize 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3540
      TabIndex        =   8
      Top             =   2100
      Width           =   3495
   End
   Begin VB.TextBox TxtModel 
      DataField       =   "Model"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3540
      TabIndex        =   7
      Top             =   1470
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3540
      TabIndex        =   15
      Top             =   870
      Width           =   3495
      Begin VB.OptionButton OptFour 
         Caption         =   "Four"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   90
         Width           =   735
      End
      Begin VB.OptionButton OptTwo 
         Caption         =   "Two"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   90
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      DataSource      =   "Adodc1"
      Height          =   7425
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   870
      Width           =   7425
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Price [OutSide Octroi ]:"
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   2820
      Width           =   2985
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Price:"
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Top             =   2145
      Width           =   2985
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Model Name:"
      Height          =   285
      Left            =   480
      TabIndex        =   12
      Top             =   1515
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Vehical Type:"
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   960
      Width           =   2985
   End
End
Attribute VB_Name = "FrmVehMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim RS As New ADODB.Recordset
    Dim Flg As String
    Dim VCN As New ADODB.Connection
    Dim STR As String
    Dim VTYPE As String
    Dim DT As Date
    Dim IM As Integer
    
Private Sub CmdAdd_Click()
    If Flg <> "add" Then
        Call BlankText(Me)
        Call CtrlEnabled(Me)
        Flg = "add"
        OptTwo.Value = True
        OptTwo.SetFocus
    End If
End Sub

Private Sub CmdCancel_Click()
    Call BlankText(Me)
    MsItem.Clear
    Call HEADING
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If OptTwo.Value = True Then
        VTYPE = "TWO"
    Else
        VTYPE = "FOUR"
    End If
    If Flg = "add" Then
        STR = "INSERT INTO VEHICALMASTER VALUES ('" & VTYPE & "','" & TxtModel.Text & "'," & TxtPrize.Text & ",  " & TxtOPrize.Text & ", '" & UN & "','" & DT & "')"
        VCN = GETCON
        VCN.Open
        VCN.BeginTrans
        VCN.Execute STR
        For I = 1 To MsItem.Rows - 1
            If MsItem.TextMatrix(I, 1) <> "" And MsItem.TextMatrix(I, 2) <> "" Then
             STR = "INSERT INTO COLOURMASTER VALUES ('" & TxtModel.Text & "','" & MsItem.TextMatrix(I, 0) & "', '" & MsItem.TextMatrix(I, 1) & "','" & MsItem.TextMatrix(I, 2) & "')"
             VCN.Execute STR
            End If
        Next
        VCN.CommitTrans
        VCN.Close
        MsgBox "RECORD SAVE SUCCESSSFULLY", vbInformation, "FD"
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        Flg = ""
    ElseIf Flg = "edit" Then
        STR = "UPDATE VEHICALMASTER SET TYPE = '" & VTYPE & "',  PRICE = " & TxtPrize.Text & ", OPRICE =  " & TxtOPrize.Text & ", USERID =  '" & UN & "',SYSDATE = '" & DT & "' WHERE MODEL = '" & TxtModel.Text & "'"
        VCN = GETCON
        VCN.Open
        VCN.BeginTrans
        VCN.Execute STR
        For I = 1 To MsItem.Rows - 1
            If MsItem.TextMatrix(I, 1) <> "" And MsItem.TextMatrix(I, 2) <> "" Then
                RS.Open "SELECT COUNT(*) FROM COLOURMASTER WHERE MODEL = '" & TxtModel.Text & "' AND INO = " & MsItem.TextMatrix(I, 0) & "", GETCON, adOpenForwardOnly, adLockReadOnly
                If RS.EOF = False And RS.BOF = False Then
                        STR = "UPDATE COLOURMASTER SET  COLOUR = '" & MsItem.TextMatrix(I, 1) & "', PHOTO = '" & MsItem.TextMatrix(I, 2) & "' WHERE MODEL = '" & TxtModel.Text & "' AND INO = " & MsItem.TextMatrix(I, 0) & ""
                        VCN.Execute STR
                Else
                        STR = "INSERT INTO COLOURMASTER VALUES ('" & TxtModel.Text & "','" & MsItem.TextMatrix(I, 0) & "', '" & MsItem.TextMatrix(I, 1) & "','" & MsItem.TextMatrix(I, 2) & "')"
                        VCN.Execute STR
                End If
                RS.Close
            End If
        Next
        VCN.CommitTrans
        VCN.Close
        MsgBox "RECORD UPDATE SUCCESSSFULLY", vbInformation, "FD"
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        Flg = ""
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo err
    Image1.Picture = LoadPicture(MsItem.TextMatrix(IM, 2))
    IM = IM + 1
    Exit Sub
err:
    MsgBox "YOU ARE IN LAST POSITION", vbInformation, "FD"
    IM = MsItem.Rows - 2
End Sub

Private Sub Command2_Click()
On Error GoTo err
    Image1.Picture = LoadPicture(MsItem.TextMatrix(IM, 2))
    IM = IM - 1
    Exit Sub
err:
    MsgBox "YOU ARE IN FIRST POSITION", vbInformation, "FD"
    IM = 1
End Sub

Private Sub MsItem_GotFocus()
    MsItem.TextMatrix(1, 0) = 1
End Sub

Private Sub MsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        MsItem.TextMatrix(MsItem.Row, MsItem.Col) = ""
    Else
        If KeyAscii = 13 And MsItem.Col = 2 Then
            CD.ShowOpen
            Image1.Picture = LoadPicture(CD.FileName)
            MsItem.TextMatrix(MsItem.Row, 2) = CD.FileName
            MsItem.Rows = MsItem.Rows + 1
            MsItem.TextMatrix(MsItem.Row + 1, 0) = Val(MsItem.TextMatrix(MsItem.Row, 0)) + 1
            MsItem.Row = MsItem.Row + 1
            MsItem.Col = 1
        Else
            MsItem = MsItem + Chr(KeyAscii)
        End If
     End If
End Sub

Private Sub Form_Load()
    Me.Top = 250
    Me.Left = 200
    Call BlankText(Me)
    Call CtrlDisabled(Me)
    Call HEADING
    DT = Date
    IM = 1
End Sub
Private Sub HEADING()
    MsItem.Clear
    MsItem.Rows = 2
    MsItem.TextMatrix(0, 0) = "SrNo"
    MsItem.TextMatrix(0, 1) = "COLOUR"
    MsItem.TextMatrix(0, 2) = "PICTURE"
    MsItem.ColWidth(1) = 1500
    MsItem.ColWidth(2) = 2500
    MsItem.ColAlignment(0) = 3
    MsItem.ColAlignmentFixed(0) = 3
    MsItem.ColAlignmentFixed(1) = 3
    MsItem.ColAlignmentFixed(2) = 3
End Sub

Private Sub TxtModel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtOPrize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtPrize_Change()
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Public Sub VIEWDATA(A As Variant)
Dim RS1 As New ADODB.Recordset
    RS.Open "SELECT * FROM VEHICALMASTER WHERE MODEL = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        If RS(0) = "TWO" Then
            OptTwo.Value = True
        Else
            OptFour.Value = True
        End If
        TxtModel.Text = RS(1) & ""
        TxtPrize.Text = RS(2) & ""
        TxtOPrize.Text = RS(3) & ""
        I = 1
        HEADING
        RS1.Open "SELECT * FROM COLOURMASTER WHERE MODEL = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        Do While Not RS1.EOF
            MsItem.TextMatrix(I, 0) = RS1(1) & ""
            MsItem.TextMatrix(I, 1) = RS1(2) & ""
            MsItem.TextMatrix(I, 2) = RS1(3) & ""
            Image1.Picture = LoadPicture(RS1(3))
            MsItem.Rows = MsItem.Rows + 1
            I = I + 1
            RS1.MoveNext
        Loop
        RS1.Close
    End If
    RS.Close
    Flg = "edit"
    CmdSave.Caption = "&Edit"
    Call CtrlEnabled(Me)
    TxtModel.Enabled = False
    TxtPrize.SetFocus
End Sub
