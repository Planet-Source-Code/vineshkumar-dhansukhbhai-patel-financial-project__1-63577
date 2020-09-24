Attribute VB_Name = "GenMod"
    Public CON As New ADODB.Connection
    Public DBPath As String
    Public SsqlDB As String
    Public Flg As String
    Public UN As String
    Public StrQryHelp As String
    '''''''''''''Connect with ExcelBook
    Public XLA As New Excel.Application
    Public XLW As New Excel.Workbook
    Public XLS As New Excel.Worksheet
    Public xlapp As Excel.Application
    Public xlapp2 As Excel.Application
    Public wkbWorkBook As Excel.Workbook
    Public wksSheet As Excel.Worksheet
    Public wkb2 As Excel.Workbook
    Public wks2 As Excel.Worksheet
    ''''''''''''''''''''''''''''''''''''''''

Public Function GETCON() As ADODB.Connection
    'On Error GoTo err:
    Dim CONN As New ADODB.Connection
    Dim SD As Long
    SsqlDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False"
    CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False"
    CONN.CommandTimeout = 1200
    Set GETCON = CONN
    Set TCONN = CONN
    Exit Function
err:     R = MsgBox("Data Base Connectivety is Close", vbOKCancel + vbCritical, "Vin Error")
End Function

Public Sub clsConn()
    If CON.State <> 0 Then
        CON.Close
    End If
End Sub

Public Function TextVal(sKey As Integer) As Integer
    On Error Resume Next
    If sKey = 8 Or sKey = 46 Or sKey = 47 Or sKey = 32 Then
        TextVal = sKey
        Exit Function
    End If
    If sKey >= 48 And sKey <= 57 Then
        sKey = 0
    ElseIf sKey < 65 Or sKey > 122 Then
        sKey = 0
    ElseIf sKey > 90 And sKey < 97 Then
        sKey = 0
    End If
    TextVal = sKey
End Function

Public Function NumVal(sKey As Integer) As Integer
    On Error Resume Next
    If sKey = 8 Or sKey = 32 Then
    
    ElseIf sKey < 46 Or sKey > 57 Then
        sKey = 0
    End If
    NumVal = sKey
End Function

Public Function DateVal(D1 As Date) As Date
    If D1 > MaxDate Or D1 < MinDate Then
        MsgBox "Worng Date, Because you select accounting year " & Year(MinDate) & " - " & Year(MaxDate)
        DateVal = MinDate
    Else
        DateVal = D1
    End If
End Function

Public Function FillText(Fname As Form)
    On Error Resume Next
    For Each ctrl In Fname.Controls
        If TypeOf ctrl Is MText Or TypeOf ctrl Is MCombo Or TypeOf ctrl Is CheckBox Then
            ctrl.BackColor = RGB(186, 179, 214)
            ctrl.ForeColor = vbBlack
        End If
    Next
End Function

Public Sub RowColor(M As Object)
    IRow = M.Rows - 1
    If Mrow = 1 Then
        For Jrow = 0 To M.Cols - 1
            M.Row = IRow
            M.Col = Jrow
            M.CellBackColor = &HFFFFC0
        Next Jrow
        Mrow = 0
    ElseIf Mrow = 0 Then
        For Jrow = 0 To M.Cols - 1
            M.Row = IRow
            M.Col = Jrow
            M.CellBackColor = &HFFFFFF
        Next Jrow
        Mrow = 1
    End If
End Sub

'Public Sub SetReportLoc(rpt As Report)
'     Dim crxtable As CRAXDRT.DatabaseTable
'     For Each crxtable In rpt.Database.Tables
'        crxtable.SetLogOnInfo "192.168.124.111", "newsstar", "sa", "kanchan"
'     Next
'End Sub


Public Sub CtrlEnabled(Frm1 As Variant)
On Error Resume Next
    Dim oCtrl As Variant
    For Each oCtrl In Frm1
        If TypeOf oCtrl Is TextBox Or TypeOf oCtrl Is ComboBox Or TypeOf oCtrl Is DTPicker Or TypeOf oCtrl Is CheckBox Or TypeOf oCtrl Is MSFlexGrid Then
           oCtrl.Enabled = True
        End If
    Next
    TxtEmpCode.Enabled = False
End Sub

Public Sub CtrlDisabled(Frm1 As Variant)
    On Error Resume Next
    Dim oCtrl As Variant
    For Each oCtrl In Frm1
        If TypeOf oCtrl Is TextBox Or TypeOf oCtrl Is ComboBox Or TypeOf oCtrl Is DTPicker Or TypeOf oCtrl Is CheckBox Or TypeOf oCtrl Is MSFlexGrid Then
           oCtrl.Enabled = False
        End If
    Next
End Sub

Public Function BlankText(Frm1 As Variant)
    On Error Resume Next
    Dim ctrl As Variant
    For Each ctrl In Frm1.Controls
        If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
            ctrl.Text = ""
        End If
    Next
End Function
