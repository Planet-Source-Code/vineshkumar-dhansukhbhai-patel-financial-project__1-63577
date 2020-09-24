VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSearch 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MsSearch 
      Height          =   8835
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   15584
      _Version        =   393216
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim J, I As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = 1500
    Me.Top = 1000
    RS.CursorLocation = adUseClient
    RS.Open StrQryHelp, GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        MsSearch.Cols = RS.Fields.Count
        For I = 0 To RS.Fields.Count - 1
            MsSearch.TextMatrix(0, I) = RS.Fields(I).Name
        Next
        J = 1
        Do While Not RS.EOF
            For I = 0 To RS.Fields.Count - 1
                MsSearch.TextMatrix(J, I) = "" & RS.Fields(I).Value
            Next
            J = J + 1
            MsSearch.Rows = MsSearch.Rows + 1
            RS.MoveNext
        Loop
    End If
    RS.Close
End Sub

Private Sub MsSearch_DblClick()
    If MsSearch.TextMatrix(MsSearch.Row, MsSearch.Col) <> "" Then
        Select Case MainMDI.Caption
        Case "Employee Master"
            FrmEmpMaster.VIEWDATA (MsSearch.TextMatrix(MsSearch.Row, 0))
            Unload Me
        Case "Customer Information Form"
            FrmCustomerinfo.VIEWDATA (MsSearch.TextMatrix(MsSearch.Row, 0))
            Unload Me
        Case "Supplier Master"
            FrmSupMaster.VIEWDATA (MsSearch.TextMatrix(MsSearch.Row, 0))
            Unload Me
        Case "Dealer Master"
            FrmSupMaster.VIEWDATA (MsSearch.TextMatrix(MsSearch.Row, 0))
            Unload Me
        Case "Vehical Master"
            FrmVehMaster.VIEWDATA (MsSearch.TextMatrix(MsSearch.Row, 0))
            Unload Me
        End Select
    End If
End Sub

Private Sub MsSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
