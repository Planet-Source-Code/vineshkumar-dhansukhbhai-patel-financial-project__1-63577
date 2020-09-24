VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MainMDI 
   BackColor       =   &H8000000C&
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   7380
      Top             =   5190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu MnuEM 
         Caption         =   "&Employee Master"
      End
      Begin VB.Menu MnuSM 
         Caption         =   "&Supplier Master"
      End
      Begin VB.Menu MnuDM 
         Caption         =   "&Dealer Master"
      End
      Begin VB.Menu MnuBM 
         Caption         =   "&Bank Master"
      End
      Begin VB.Menu MnuVM 
         Caption         =   "&Vehical Master"
      End
   End
   Begin VB.Menu MnuTrans 
      Caption         =   "&Transaction"
      Begin VB.Menu MnuTWCI 
         Caption         =   "Two Wheeler &Customer Info"
      End
   End
   Begin VB.Menu MnuUTI 
      Caption         =   "&Utility"
      Begin VB.Menu MnuSrch 
         Caption         =   "&Search"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "MainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim Con1 As New ADODB.Connection

Private Sub MDIForm_DblClick()
    Dim R As String
    R = MsgBox("Are you sure, you can change you DB path?", vbInformation + vbYesNo, "OrderPM")
    If R = vbYes Then
        CD.ShowOpen
        Con1.BeginTrans
        SSql = "UPDATE DBPATH SET DPATH = '" & CD.FileName & "'"
        Con1.Execute SSql
        Con1.CommitTrans
        Con1.Close
        MsgBox "Path update sucessfully, Restart your project", vbInformation, "OrderPM"
        End
    End If
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
    Me.Top = 0
    Me.Left = 0
    Con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\FDPath.mdb;Persist Security Info=False"
    RS.Open "SELECT DPATH FROM DBPATH", Con1, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = True And RS.BOF = True Then
        MsgBox "First Paste FDPath.MBD in C Drive and start your project", vbInformation, "FD"
        End
    Else
        'MsgBox RS(0)
        DBPath = RS(0)
    End If
    UN = "Vinesh"
End Sub

Private Sub MnuBM_Click()
    Me.Caption = "Bank Master"
    FrmBankMaster.Show
End Sub

Private Sub MnuDM_Click()
    Me.Caption = "Dealer Master"
    FrmSupMaster.Show
End Sub

Private Sub MnuEM_Click()
    Me.Caption = "Employee Master"
    FrmEmpMaster.Show
End Sub

Private Sub MnuExit_Click()
    Dim R As String
    R = MsgBox("Are you Sure, You want to Queit?", vbInformation + vbYesNo, "FDP")
    If R = vbYes Then
        End
    End If
End Sub

Private Sub MnuSM_Click()
    Me.Caption = "Supplier Master"
    FrmSupMaster.Show
End Sub

Private Sub MnuSrch_Click()
    Select Case Me.Caption
    Case "Employee Master"
        StrQryHelp = "SELECT EmpMaster.EmpCode, EmpMaster.EmpName, EmpMaster.EmpFname, EmpMaster.Address, EmpMaster.PAddress, EmpMaster.Phone, EmpMaster.Mobile, EmpMaster.email, EmpMaster.BDate, EmpMaster.DOJ FROM EmpMaster"
        FrmSearch.Show
    Case "Customer Information Form"
        StrQryHelp = "SELECT CustomerInfo.Code, CustomerInfo.Name, CustomerInfo.Address, CustomerInfo.Phone, CustomerInfo.Mobile, CustomerInfo.Email, CustomerInfo.IDate, CustomerInfo.EmpCode, CustomerInfo.Model, CustomerInfo.Colour FROM CustomerInfo"
        FrmSearch.Show
    Case "Supplier Master"
        StrQryHelp = "SELECT SupDealMaster.DCode, SupDealMaster.DName, SupDealMaster.DAddress, SupDealMaster.DPhone, SupDealMaster.CPMobile FROM SupDealMaster where SupDealMaster.Type = 'Supplier'"
        FrmSearch.Show
    Case "Dealer Master"
        StrQryHelp = "SELECT SupDealMaster.DCode, SupDealMaster.DName, SupDealMaster.DAddress, SupDealMaster.DPhone, SupDealMaster.CPMobile FROM SupDealMaster where SupDealMaster.Type = 'Dealer'"
        FrmSearch.Show
    Case "Vehical Master"
        StrQryHelp = "SELECT VehicalMaster.Model, VehicalMaster.Type, VehicalMaster.Price, VehicalMaster.Oprice FROM VehicalMaster"
        FrmSearch.Show
    Case "Bank Master"
        'StrQryHelp = "SELECT VehicalMaster.Model, VehicalMaster.Type, VehicalMaster.Price, VehicalMaster.Oprice FROM VehicalMaster"
        FrmSearch.Show
    End Select
End Sub

Private Sub MnuTWCI_Click()
    Me.Caption = "Customer Information Form"
    FrmCustomerinfo.Show
End Sub

Private Sub MnuVM_Click()
    Me.Caption = "Vehical Master"
    FrmVehMaster.Show
End Sub
