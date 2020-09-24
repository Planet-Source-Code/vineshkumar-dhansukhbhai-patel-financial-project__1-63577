VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEmpMaster 
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   345
      Left            =   8820
      TabIndex        =   16
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cl&ose"
      Height          =   345
      Left            =   13620
      TabIndex        =   20
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   12420
      TabIndex        =   19
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   11220
      TabIndex        =   18
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   10020
      TabIndex        =   17
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   345
      Left            =   7620
      TabIndex        =   15
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton CmdPicture 
      Caption         =   "..."
      Height          =   345
      Left            =   13860
      TabIndex        =   21
      Top             =   2160
      Width           =   435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Is Active"
      DataField       =   "IsActive"
      DataSource      =   "Adodc1"
      Height          =   345
      Left            =   9360
      TabIndex        =   14
      Top             =   8490
      Width           =   2805
   End
   Begin MSComCtl2.DTPicker DTP2 
      DataSource      =   "Adodc1"
      Height          =   345
      Left            =   9360
      TabIndex        =   9
      Top             =   6780
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Format          =   22806529
      CurrentDate     =   38681
   End
   Begin MSComCtl2.DTPicker DTP1 
      DataSource      =   "Adodc1"
      Height          =   345
      Left            =   2490
      TabIndex        =   8
      Top             =   6780
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Format          =   22806529
      CurrentDate     =   38681
   End
   Begin VB.TextBox TxtRemarks 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   13
      Top             =   8490
      Width           =   4425
   End
   Begin VB.TextBox TxtSal 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   11
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox TxtIns 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   9360
      TabIndex        =   12
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox TxtRefrence 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   10
      Top             =   7320
      Width           =   4455
   End
   Begin VB.TextBox TxtEmail 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   7
      Top             =   6180
      Width           =   2775
   End
   Begin VB.TextBox TxtMobile 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   6
      Top             =   5610
      Width           =   2775
   End
   Begin VB.TextBox TxtPhone 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   5
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox TxtPAddress 
      DataSource      =   "Adodc1"
      Height          =   2340
      Left            =   9360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2550
      Width           =   4485
   End
   Begin VB.TextBox TxtAddress 
      DataSource      =   "Adodc1"
      Height          =   2340
      Left            =   2490
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2550
      Width           =   4485
   End
   Begin VB.TextBox TxtFName 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   2
      Top             =   2040
      Width           =   4515
   End
   Begin VB.TextBox TxtName 
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2490
      TabIndex        =   1
      Top             =   1500
      Width           =   4515
   End
   Begin VB.TextBox TxtEmpCode 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2490
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.Image Photo 
      Height          =   2385
      Left            =   11070
      Picture         =   "FrmEmpMaster.frx":0000
      Stretch         =   -1  'True
      Top             =   90
      Width           =   2745
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   315
      Left            =   1470
      TabIndex        =   35
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Insentive %"
      Height          =   315
      Left            =   8190
      TabIndex        =   34
      Top             =   7950
      Width           =   1125
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      Height          =   315
      Left            =   990
      TabIndex        =   33
      Top             =   7950
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Refrence"
      Height          =   315
      Left            =   990
      TabIndex        =   32
      Top             =   7365
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Join Date"
      Height          =   315
      Left            =   8340
      TabIndex        =   31
      Top             =   6795
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date"
      Height          =   315
      Left            =   990
      TabIndex        =   30
      Top             =   6795
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Emp Code"
      Height          =   315
      Left            =   990
      TabIndex        =   29
      Top             =   990
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      Height          =   315
      Left            =   990
      TabIndex        =   28
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      Height          =   315
      Left            =   990
      TabIndex        =   27
      Top             =   5070
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Perment Address"
      Height          =   315
      Left            =   7620
      TabIndex        =   26
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   315
      Left            =   990
      TabIndex        =   25
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      Height          =   315
      Left            =   990
      TabIndex        =   24
      Top             =   2055
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      Height          =   315
      Left            =   990
      TabIndex        =   23
      Top             =   1515
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   315
      Left            =   990
      TabIndex        =   22
      Top             =   6225
      Width           =   1455
   End
End
Attribute VB_Name = "FrmEmpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim SSql As String

Private Sub CmdAdd_Click()
    If Flg = "" Then
        Flg = "add"
        DTP1.Value = Date
        DTP2.Value = Date
        Call CtrlEnabled(Me)
        Call BlankText(Me)
        CmdSave.Caption = "&Save"
        TxtEmpCode.Enabled = False
        TxtName.SetFocus
    End If
End Sub

Private Sub CmdCancel_Click()
    Flg = ""
    Call CtrlDisabled(Me)
    Call BlankText(Me)
    Set RS = Nothing
    CmdSave.Caption = "&Save"
    CmdAdd.SetFocus
End Sub

Private Sub CmdClose_Click()
    Unload Me
    Flg = ""
    Set FrmEmpMaster = Nothing
    Set RS = Nothing
End Sub

Private Sub CmdDelete_Click()
    If Flg = "edit" Then
        Dim R As String
        R = MsgBox("ARE YOU SURE, YOU WANT TO DELETE THIS RECORD?", vbInformation + vbYesNo, "FD")
        If R = vbYes Then
            Dim Con1 As New ADODB.Connection
            Con1 = GETCON
            Con1.Open
            Con1.BeginTrans
            SSql = "DELETE FROM EMPMASTER WHERE EMPCODE = '" & TxtEmpCode.Text & "'"
            Con1.Execute SSql
            Con1.CommitTrans
            Con1.Close
            MsgBox "RECORD DELETE SUCESSFULLY", vbInformation, "FD"
            Flg = ""
            Call CtrlDisabled(Me)
            Call BlankText(Me)
            Adodc1.Refresh
            CmdAdd.SetFocus
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If Flg = "" Then
        Flg = "edit"
        CmdSave.Caption = "&Update"
        Call CtrlEnabled(Me)
        TxtName.SetFocus
    End If
End Sub

Private Sub CmdSave_Click()
    Dim Con1 As New ADODB.Connection
    If Flg = "add" Then
        Con1 = GETCON
        Con1.Open
        Con1.BeginTrans
        'SSql = "INSERT INTO EMPMASTER (EMPCODE, EMPNAME, EMPFNAME, ADDRESS, PADDRESS, PHONE, MOBILE, EMAIL, BDATE, DOJ, REFERENCE, SAL, INS, REMARKS, ISACTIVE, USERID, SYSDATE) VALUES ('" & TxtEmpCode.Text & "','" & TxtName.Text & "','" & TxtFName.Text & "','" & TxtAddress.Text & "','" & TxtPAddress.Text & "','" & TxtPhone.Text & "','" & TxtMobile.Text & "','" & TxtEmail.Text & "',#" & Format(DTP1.Value, "MM/DD/YYYY") & "#,#" & Format(DTP2.Value, "MM/DD/YYYY") & "#, '" & TxtRefrence.Text & "','" & TxtSal.Text & "','" & TxtIns.Text & "','" & TxtRemarks.Text & "','" & Check1.Value & "','" & UN & "', #" & Date & "#"
        SSql = "INSERT INTO EmpMaster VALUES('" & TxtEmpCode.Text & "','" & TxtName.Text & "','" & TxtFName.Text & "','" & TxtAddress.Text & "','" & TxtPAddress.Text & "','" & TxtPhone.Text & "','" & TxtMobile.Text & "','" & TxtEmail.Text & "',#" & Format(DTP1.Value, "MM/dd/yyyy") & "#,#" & Format(DTP2.Value, "MM/dd/yyyy") & "#, '" & TxtRefrence.Text & "','" & TxtSal.Text & "','" & TxtIns.Text & "','" & TxtRemarks.Text & "','" & Check1.Value & "','" & UN & "', #" & Format(Date, "MM/dd/yyyy") & "#)"
        Con1.Execute SSql
        Con1.CommitTrans
        Con1.Close
        MsgBox "RECORD SAVE SUCESSFULLY", vbInformation, "FD"
        Flg = ""
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        Adodc1.Refresh
        CmdAdd.SetFocus
    End If
    If Flg = "edit" Then
        Con1 = GETCON
        Con1.Open
        Con1.BeginTrans
        SSql = "UPDATE EMPMASTER SET EMPNAME='" & TxtName.Text & "', EMPFNAME='" & TxtFName.Text & "', ADDRESS='" & TxtAddress.Text & "', PADDRESS='" & TxtPAddress.Text & "', PHONE='" & TxtPhone.Text & "', MOBILE='" & TxtMobile.Text & "', EMAIL='" & TxtEmail.Text & "', BDATE=#" & Format(DTP1.Value, "MM/dd/yyyy") & "#, DOJ=#" & Format(DTP2.Value, "MM/dd/yyyy") & "#, REFERENCE='" & TxtRefrence.Text & "', SAL='" & TxtSal.Text & "', INS='" & TxtIns.Text & "', REMARKS='" & TxtRemarks.Text & "', ISACTIVE='" & Check1.Value & "', USERID='" & UN & "', SYSDATE=#" & Format(Date, "MM/dd/yyyy") & "# WHERE EMPCODE='" & TxtEmpCode.Text & "'"
        Con1.Execute SSql
        Con1.CommitTrans
        Con1.Close
        MsgBox "RECORD UPDATE SUCESSFULLY", vbInformation, "FD"
        Flg = ""
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        Adodc1.Refresh
        CmdSave.Caption = "&Save"
        CmdAdd.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 250
    Me.Left = 200
    Call CtrlDisabled(Me)
    Call BlankText(Me)
End Sub

Private Sub TxtName_LostFocus()
    If TxtName.Text <> "" And Flg = "add" Then
        RS.Open "SELECT COUNT(EMPCODE) FROM EMPMASTER WHERE LEFT(EMPCODE,3) = '" & Left(TxtName.Text, 3) & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        If RS.EOF = True And RS.BOF = True Then
            TxtEmpCode.Text = UCase(Left(TxtName.Text, 3)) & Format(1, "0000")
        Else
            TxtEmpCode.Text = UCase(Left(TxtName.Text, 3)) & Format(RS(0) + 1, "0000")
        End If
        RS.Close
    End If
End Sub

Public Sub VIEWDATA(A As Variant)
    RS.Open "SELECT * FROM EMPMASTER WHERE EMPCODE = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        TxtEmpCode.Text = RS(0)
        TxtName.Text = RS(1)
        TxtFName.Text = RS(2)
        TxtAddress.Text = RS(3)
        TxtPAddress.Text = RS(4)
        TxtPhone.Text = RS(5)
        TxtMobile.Text = RS(6)
        TxtEmail.Text = RS(7)
        DTP1.Value = RS(8)
        DTP2.Value = RS(9)
        TxtRefrence.Text = RS(10)
        TxtSal.Text = RS(11)
        TxtIns.Text = RS(12)
        TxtRemarks.Text = RS(13)
        Check1.Value = RS(14)
    End If
    RS.Close
    Flg = "edit"
    Call CtrlEnabled(Me)
    TxtEmpCode.Enabled = False
    TxtName.Enabled = False
End Sub
