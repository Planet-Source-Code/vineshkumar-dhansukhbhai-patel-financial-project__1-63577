VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSupMaster 
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
   Begin VB.TextBox TxtCPAdd 
      Height          =   1845
      Left            =   10230
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1338
      Width           =   3615
   End
   Begin VB.TextBox TxtCPName 
      Height          =   360
      Left            =   10230
      TabIndex        =   11
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox TxtCpPh 
      Height          =   360
      Left            =   10230
      TabIndex        =   13
      Top             =   3441
      Width           =   1695
   End
   Begin VB.TextBox TxtRef 
      Height          =   360
      Left            =   10230
      TabIndex        =   14
      Top             =   4027
      Width           =   3735
   End
   Begin VB.TextBox TxtRemark 
      Height          =   360
      Left            =   10230
      TabIndex        =   15
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox TxtName 
      Height          =   360
      Left            =   3000
      TabIndex        =   6
      Top             =   1335
      Width           =   3615
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cl&ose"
      Height          =   345
      Left            =   13620
      TabIndex        =   4
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   12420
      TabIndex        =   3
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   11220
      TabIndex        =   2
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   10020
      TabIndex        =   1
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   345
      Left            =   8820
      TabIndex        =   0
      Top             =   9570
      Width           =   1155
   End
   Begin VB.TextBox TxtCode 
      BackColor       =   &H00404040&
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsItem 
      Bindings        =   "FrmSupMaster.frx":0000
      Height          =   2625
      Left            =   4290
      TabIndex        =   16
      Top             =   6600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4630
      _Version        =   393216
      Cols            =   5
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Y/N"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   5340
      Width           =   615
   End
   Begin VB.TextBox TxtMob 
      Height          =   360
      Left            =   3000
      TabIndex        =   9
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox TxtPh 
      Height          =   360
      Left            =   3000
      TabIndex        =   8
      Top             =   4020
      Width           =   1695
   End
   Begin VB.TextBox TxtAdd 
      Height          =   1845
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1950
      Width           =   3615
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slabe Details"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4140
      TabIndex        =   30
      Top             =   6210
      Width           =   1260
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier/Dealer Details"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   660
      TabIndex        =   29
      Top             =   240
      Width           =   2205
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   2835
      Left            =   4170
      Top             =   6510
      Width           =   7335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   5445
      Left            =   7860
      Top             =   510
      Width           =   6405
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   5445
      Left            =   690
      Top             =   510
      Width           =   6405
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   930
      TabIndex        =   28
      Top             =   1335
      Width           =   1365
   End
   Begin VB.Label Label7 
      Caption         =   "Address:"
      Height          =   255
      Left            =   8670
      TabIndex        =   27
      Top             =   1335
      Width           =   1365
   End
   Begin VB.Label Label12 
      Caption         =   "Is Active:"
      Height          =   255
      Left            =   930
      TabIndex        =   26
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label Label11 
      Caption         =   "Remarks:"
      Height          =   255
      Left            =   8670
      TabIndex        =   25
      Top             =   4740
      Width           =   1365
   End
   Begin VB.Label Label10 
      Caption         =   "Reference:"
      Height          =   255
      Left            =   8670
      TabIndex        =   24
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Label Label9 
      Caption         =   "Mobile No:"
      Height          =   255
      Left            =   930
      TabIndex        =   23
      Top             =   4740
      Width           =   1365
   End
   Begin VB.Label Label8 
      Caption         =   "Phone No:"
      Height          =   255
      Left            =   8670
      TabIndex        =   22
      Top             =   3450
      Width           =   1365
   End
   Begin VB.Label Label6 
      Caption         =   "Name:"
      Height          =   255
      Left            =   8670
      TabIndex        =   21
      Top             =   780
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Persone Details"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   7860
      TabIndex        =   20
      Top             =   240
      Width           =   2265
   End
   Begin VB.Label Label4 
      Caption         =   "Phone No:"
      Height          =   255
      Left            =   930
      TabIndex        =   19
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "Address:"
      Height          =   255
      Left            =   930
      TabIndex        =   18
      Top             =   1950
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   255
      Left            =   930
      TabIndex        =   17
      Top             =   780
      Width           =   1365
   End
End
Attribute VB_Name = "FrmSupMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim MyCn As New ADODB.Connection
Dim SSql As String

Private Sub Check1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check1_LostFocus()
MsItem.TextMatrix(1, 0) = 1
MsItem.Row = 1
MsItem.Col = 1
MsItem.SetFocus
End Sub

Private Sub CmdAdd_Click()
    If Flg <> "add" Then
        Call BlankText(Me)
        Call FLXHEAD
        Flg = "add"
        Call CtrlEnabled(Me)
        TxtCode.Enabled = False
        CmdSave.Caption = "&save"
        TxtName.SetFocus
    End If
End Sub

Private Sub CmdCancel_Click()
Call BlankText(Me)
MsItem.Clear
Call FLXHEAD
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim type1 As String
    Dim DT As Date
    DT = Date
    If MainMDI.Caption = "Supplier Master" Then
        type1 = "Supplier"
    ElseIf MainMDI.Caption = "Dealer Master" Then
        type1 = "Dealer"
    End If
    If Flg = "add" Then
        MyCn = GETCON
        MyCn.Open
        MyCn.BeginTrans
        SSql = "insert into supdealmaster values('" & TxtCode.Text & "','" & TxtName.Text & "','" & TxtAdd.Text & "','" & TxtPh.Text & "','" & TxtCPName.Text & "','" & TxtCPAdd.Text & "','" & TxtCpPh.Text & "','" & TxtMob.Text & "','" & TxtRef.Text & "','" & type1 & "','" & TxtRemark.Text & "'," & Check1.Value & " ,'" & UN & "','" & DT & "' )"
        MyCn.Execute SSql
        For I = 1 To MsItem.Rows - 1
            SSql = "insert into slabmaster values('" & TxtCode.Text & "'," & MsItem.TextMatrix(I, 0) & "," & MsItem.TextMatrix(I, 1) & "," & MsItem.TextMatrix(I, 2) & "," & MsItem.TextMatrix(I, 3) & "," & MsItem.TextMatrix(I, 4) & ")"
            MyCn.Execute SSql
        Next
        MyCn.CommitTrans
        MsgBox "Record Save Sucessfully", vbInformation, "FD"
        MyCn.Close
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        CmdAdd.SetFocus
        Flg = ""
    ElseIf Flg = "edit" Then
        MyCn = GETCON
        MyCn.Open
        MyCn.BeginTrans
        SSql = "update supdealmaster set dname = '" & TxtName.Text & "', daddress='" & TxtAdd.Text & "',dphone = '" & TxtPh.Text & "', cpname = '" & TxtCPName.Text & "', cpaddress = '" & TxtCPAdd.Text & "', cpphone = '" & TxtCpPh.Text & "', cpmobile = '" & TxtMob.Text & "',reference = '" & TxtRef.Text & "',type = '" & type1 & "',remarks = '" & TxtRemark.Text & "',isactive = " & Check1.Value & " ,userid = '" & UN & "', sysdate = '" & DT & "' where dcode ='" & TxtCode.Text & "'"
        MyCn.Execute SSql
        For I = 1 To MsItem.Rows - 1
            RS.Open "select count(*) from slabmaster where code = '" & A & "' and ino = '" & MsItem.TextMatrix(I, 0) & "'", GETCON, adOpenForwardOnly, adLockReadOnly
            If RS.EOF = True And RS.BOF = True Then
                SSql = "insert into slabmaster values('" & TxtCode.Text & "'," & MsItem.TextMatrix(I, 0) & "," & MsItem.TextMatrix(I, 1) & "," & MsItem.TextMatrix(I, 2) & "," & MsItem.TextMatrix(I, 3) & "," & MsItem.TextMatrix(I, 4) & ")"
            Else
                SSql = "update slabmaster set  fs = " & MsItem.TextMatrix(I, 1) & ", ts = " & MsItem.TextMatrix(I, 2) & ", amount =" & MsItem.TextMatrix(I, 3) & ", per = " & MsItem.TextMatrix(I, 4) & ") where code ='" & TxtCode.Text & "' and ino = " & MsItem.TextMatrix(I, 0) & ""
            End If
            MyCn.Execute SSql
        Next
        MyCn.CommitTrans
        MsgBox "Record Update Sucessfully", vbInformation, "FD"
        MyCn.Close
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        CmdAdd.SetFocus
        Flg = ""
    End If
    CmdSave.Caption = "&Save"
End Sub

Private Sub MsItem_KeyDown(KeyCode As Integer, Shift As Integer)

     If KeyCode = 46 Then
        MsItem.TextMatrix(MsItem.Row, MsItem.Col) = ""
    End If
End Sub

Private Sub MsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        MsItem.TextMatrix(MsItem.Row, MsItem.Col) = ""
    Else
        If KeyAscii = 13 And MsItem.Col = 4 Then
            MsItem.Rows = MsItem.Rows + 1
            MsItem.TextMatrix(MsItem.Row + 1, 0) = Val(MsItem.TextMatrix(MsItem.Row, 0)) + 1
            MsItem.Row = MsItem.Row + 1
            MsItem.Col = 1
        Else
            MsItem = MsItem + Chr(KeyAscii)
        End If
     End If
End Sub

Private Sub MsItem_LostFocus()
    CmdSave.SetFocus
End Sub

Private Sub Form_Load()
    Me.Top = 250
    Me.Left = 200
    Call CtrlDisabled(Me)
    Call BlankText(Me)
    Call FLXHEAD
    CmdSave.Caption = "&Save"
End Sub

Private Sub FLXHEAD()
    MsItem.Rows = 2
    MsItem.Cols = 5
    MsItem.TextMatrix(0, 0) = "SrNo"
    MsItem.TextMatrix(0, 1) = "FS"
    MsItem.TextMatrix(0, 2) = "TS"
    MsItem.TextMatrix(0, 3) = "AMOUNT"
    MsItem.TextMatrix(0, 4) = "PER"
End Sub

Private Sub TxtAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtCPAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtCPName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtCpPh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtMob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtName_LostFocus()
  If TxtName.Text <> "" And Flg = "add" Then
        RS.Open "SELECT COUNT(dCODE) FROM supdealMASTER WHERE LEFT(dCODE,3) = '" & Left(TxtName.Text, 3) & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        If RS.EOF = True And RS.BOF = True Then
            TxtCode = UCase(Left(TxtName.Text, 3)) & Format(1, "0000")
        Else
            TxtCode.Text = UCase(Left(TxtName.Text, 3)) & Format(RS(0) + 1, "0000")
        End If
        RS.Close
    End If
End Sub

Private Sub TxtPh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtRemark_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Public Sub VIEWDATA(A As Variant)
Dim RS1 As New ADODB.Recordset
    RS.Open "SELECT * FROM SUPDEALMASTER WHERE DCODE = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        TxtCode.Text = RS(0) & ""
        TxtName.Text = RS(1) & ""
        TxtAdd.Text = RS(2) & ""
        TxtPh.Text = RS(3) & ""
        TxtCPName.Text = RS(4) & ""
        TxtCPAdd.Text = RS(5) & ""
        TxtCpPh.Text = RS(6) & ""
        TxtMob.Text = RS(7) & ""
        TxtRef.Text = RS(8) & ""
        TxtRemark.Text = RS(10) & ""
        Check1.Value = RS(11) & ""
        I = 1
        RS1.Open "SELECT * FROM SLABMASTER WHERE CODE = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        Do While Not RS1.EOF
            MsItem.TextMatrix(I, 0) = RS1(1) & ""
            MsItem.TextMatrix(I, 1) = RS1(2) & ""
            MsItem.TextMatrix(I, 2) = RS1(3) & ""
            MsItem.TextMatrix(I, 3) = RS1(4) & ""
            MsItem.TextMatrix(I, 4) = RS1(5) & ""
            MsItem.Rows = MsItem.Rows + 1
            I = I + 1
            RS1.MoveNext
        Loop
        RS1.Close
    End If
    RS.Close
    Call CtrlEnabled(Me)
    Flg = "edit"
    CmdSave.Caption = "&Edit"
    TxtCode.Enabled = False
    TxtName.SetFocus
End Sub
