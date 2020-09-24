VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCustomerinfo 
   BorderStyle     =   0  'None
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
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   345
      Left            =   8850
      TabIndex        =   40
      Top             =   9600
      Width           =   1155
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   10050
      TabIndex        =   39
      Top             =   9600
      Width           =   1155
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   11250
      TabIndex        =   38
      Top             =   9600
      Width           =   1155
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   12450
      TabIndex        =   37
      Top             =   9600
      Width           =   1155
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cl&ose"
      Height          =   345
      Left            =   13650
      TabIndex        =   36
      Top             =   9600
      Width           =   1155
   End
   Begin TabDlg.SSTab STab 
      Height          =   9555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   16854
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Scheme"
      TabPicture(0)   =   "FrmCustomerInfo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(3)=   "MsItem"
      Tab(0).Control(4)=   "TxtCAmt"
      Tab(0).Control(5)=   "TxtDBD"
      Tab(0).Control(6)=   "TxtCBD"
      Tab(0).Control(7)=   "TxtTotalEmi"
      Tab(0).Control(8)=   "TxtDPay"
      Tab(0).Control(9)=   "TxtTotalAmt"
      Tab(0).Control(10)=   "TxtNetPF"
      Tab(0).Control(11)=   "TXTAdEmi"
      Tab(0).Control(12)=   "CboColor"
      Tab(0).Control(13)=   "CboModel"
      Tab(0).Control(14)=   "TxtROI"
      Tab(0).Control(15)=   "TxtEMI"
      Tab(0).Control(16)=   "TxtTenor"
      Tab(0).Control(17)=   "TxtLAmt"
      Tab(0).Control(18)=   "TxtPF"
      Tab(0).Control(19)=   "TxtPrice"
      Tab(0).Control(20)=   "LblTotalDP"
      Tab(0).Control(21)=   "Label1"
      Tab(0).Control(22)=   "Image1"
      Tab(0).Control(23)=   "lblLabels(16)"
      Tab(0).Control(24)=   "lblLabels(15)"
      Tab(0).Control(25)=   "lblLabels(14)"
      Tab(0).Control(26)=   "lblLabels(13)"
      Tab(0).Control(27)=   "lblLabels(12)"
      Tab(0).Control(28)=   "lblLabels(11)"
      Tab(0).Control(29)=   "lblLabels(10)"
      Tab(0).Control(30)=   "lblLabels(9)"
      Tab(0).Control(31)=   "lblLabels(8)"
      Tab(0).Control(32)=   "lblLabels(7)"
      Tab(0).Control(33)=   "lblLabels(6)"
      Tab(0).Control(34)=   "lblLabels(5)"
      Tab(0).Control(35)=   "lblLabels(4)"
      Tab(0).Control(36)=   "lblLabels(3)"
      Tab(0).Control(37)=   "lblLabels(2)"
      Tab(0).Control(38)=   "lblLabels(1)"
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Customer Information"
      TabPicture(1)   =   "FrmCustomerInfo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(1)=   "Command5"
      Tab(1).Control(2)=   "Command4"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "CboBankName"
      Tab(1).Control(5)=   "CboEmpName"
      Tab(1).Control(6)=   "TxtNOtes"
      Tab(1).Control(7)=   "TxtEmail"
      Tab(1).Control(8)=   "TxtMobile"
      Tab(1).Control(9)=   "TxtPhone"
      Tab(1).Control(10)=   "TxtAddress"
      Tab(1).Control(11)=   "TxtName"
      Tab(1).Control(12)=   "TxtCode"
      Tab(1).Control(13)=   "DTP1"
      Tab(1).Control(14)=   "CboBankCode"
      Tab(1).Control(15)=   "CboEmpCode"
      Tab(1).Control(16)=   "CboType"
      Tab(1).Control(17)=   "LblBankName"
      Tab(1).Control(18)=   "LblEmpName"
      Tab(1).Control(19)=   "Label11"
      Tab(1).Control(20)=   "Label10"
      Tab(1).Control(21)=   "Label9"
      Tab(1).Control(22)=   "Label8"
      Tab(1).Control(23)=   "Label7"
      Tab(1).Control(24)=   "Label6"
      Tab(1).Control(25)=   "Label5"
      Tab(1).Control(26)=   "Label4"
      Tab(1).Control(27)=   "Label3"
      Tab(1).Control(28)=   "Label2"
      Tab(1).Control(29)=   "lblLabels(0)"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "Document"
      TabPicture(2)   =   "FrmCustomerInfo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check2"
      Tab(2).Control(1)=   "Command6"
      Tab(2).Control(2)=   "TxtCopy"
      Tab(2).Control(3)=   "TxtDoc"
      Tab(2).Control(4)=   "MsDoc"
      Tab(2).Control(5)=   "Label51"
      Tab(2).Control(6)=   "Label50"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Utility"
      TabPicture(3)   =   "FrmCustomerInfo.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label52"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "CboLoginStatus"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame3"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame3 
         Caption         =   "Final Document"
         ForeColor       =   &H00FF0000&
         Height          =   4305
         Left            =   300
         TabIndex        =   122
         Top             =   5070
         Visible         =   0   'False
         Width           =   14265
         Begin VB.CheckBox Check5 
            BackColor       =   &H00404040&
            Caption         =   "Cheque Agreement Done"
            ForeColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   7410
            TabIndex        =   143
            Top             =   330
            Width           =   2715
         End
         Begin VB.TextBox TxtDoc1 
            Height          =   375
            Left            =   240
            TabIndex        =   139
            Top             =   870
            Width           =   10485
         End
         Begin VB.TextBox TxtCopy1 
            Height          =   375
            Left            =   10800
            TabIndex        =   138
            Top             =   870
            Width           =   1605
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Add"
            Height          =   375
            Left            =   12480
            TabIndex        =   137
            Top             =   870
            Width           =   1215
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00404040&
            Caption         =   "Final Document Collect"
            ForeColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   3180
            TabIndex        =   136
            Top             =   330
            Width           =   2715
         End
         Begin MSFlexGridLib.MSFlexGrid MsDoc1 
            Height          =   2895
            Left            =   240
            TabIndex        =   140
            Top             =   1320
            Width           =   13755
            _ExtentX        =   24262
            _ExtentY        =   5106
            _Version        =   393216
            Cols            =   3
         End
         Begin VB.Label Label60 
            Caption         =   "Document Details"
            Height          =   375
            Left            =   240
            TabIndex        =   142
            Top             =   630
            Width           =   3315
         End
         Begin VB.Label Label59 
            Caption         =   "No Of Copies"
            Height          =   375
            Left            =   10830
            TabIndex        =   141
            Top             =   630
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Vehical Information"
         ForeColor       =   &H00FF0000&
         Height          =   3645
         Left            =   300
         TabIndex        =   121
         Top             =   1260
         Visible         =   0   'False
         Width           =   14295
         Begin VB.TextBox TxtPNo 
            Height          =   375
            Left            =   8310
            TabIndex        =   135
            Top             =   2655
            Width           =   2535
         End
         Begin VB.TextBox TxtEngNo 
            Height          =   375
            Left            =   8340
            TabIndex        =   134
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox TxtCNo 
            Height          =   375
            Left            =   8370
            TabIndex        =   133
            Top             =   945
            Width           =   2535
         End
         Begin VB.TextBox TxtRCBook 
            Height          =   375
            Left            =   1350
            TabIndex        =   132
            Top             =   2655
            Width           =   2535
         End
         Begin VB.ComboBox CboSupCode 
            Height          =   360
            Left            =   1380
            TabIndex        =   131
            Top             =   1807
            Width           =   2505
         End
         Begin MSComCtl2.DTPicker DTPVD 
            Height          =   345
            Left            =   1380
            TabIndex        =   130
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Format          =   22806529
            CurrentDate     =   38692
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00404040&
            Caption         =   "Vehical Delivered"
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   5490
            TabIndex        =   123
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            Caption         =   "Parsing No:"
            Height          =   285
            Left            =   7020
            TabIndex        =   129
            Top             =   2730
            Width           =   1245
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            Caption         =   "Engine No:"
            Height          =   285
            Left            =   7050
            TabIndex        =   128
            Top             =   1875
            Width           =   1245
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            Caption         =   "Chasis No:"
            Height          =   285
            Left            =   7050
            TabIndex        =   127
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Caption         =   "RC Book:"
            Height          =   285
            Left            =   300
            TabIndex        =   126
            Top             =   2730
            Width           =   975
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            Caption         =   "Sup Code:"
            Height          =   285
            Left            =   300
            TabIndex        =   125
            Top             =   1875
            Width           =   1005
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   285
            Left            =   300
            TabIndex        =   124
            Top             =   1020
            Width           =   1005
         End
      End
      Begin VB.ComboBox CboLoginStatus 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         ItemData        =   "FrmCustomerInfo.frx":0070
         Left            =   1920
         List            =   "FrmCustomerInfo.frx":007D
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   660
         Width           =   2445
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Document Collect"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -74760
         TabIndex        =   118
         Top             =   750
         Width           =   2085
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Height          =   375
         Left            =   -61680
         TabIndex        =   115
         Top             =   1410
         Width           =   1215
      End
      Begin VB.TextBox TxtCopy 
         Height          =   375
         Left            =   -63360
         TabIndex        =   114
         Top             =   1410
         Width           =   1605
      End
      Begin VB.TextBox TxtDoc 
         Height          =   375
         Left            =   -74790
         TabIndex        =   113
         Top             =   1410
         Width           =   11355
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Customer Respones"
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   -73350
         TabIndex        =   112
         Top             =   7980
         Width           =   2445
      End
      Begin MSFlexGridLib.MSFlexGrid MsDoc 
         Height          =   7125
         Left            =   -74760
         TabIndex        =   111
         Top             =   2040
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   12568
         _Version        =   393216
         Cols            =   3
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00808080&
         Caption         =   "C&lear"
         Height          =   315
         Left            =   -62250
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   9120
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00808080&
         Caption         =   "&Next >"
         Height          =   315
         Left            =   -61230
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   9120
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scheme"
         ForeColor       =   &H00C00000&
         Height          =   8685
         Left            =   -64920
         TabIndex        =   70
         Top             =   360
         Width           =   4635
         Begin VB.Label Label49 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   108
            Top             =   8280
            Width           =   2535
         End
         Begin VB.Label Label48 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   107
            Top             =   7800
            Width           =   2535
         End
         Begin VB.Label Label47 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   106
            Top             =   7350
            Width           =   2535
         End
         Begin VB.Label Label46 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   105
            Top             =   6885
            Width           =   2535
         End
         Begin VB.Label Label45 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   104
            Top             =   6420
            Width           =   2535
         End
         Begin VB.Label Label44 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   103
            Top             =   5955
            Width           =   2535
         End
         Begin VB.Label Label43 
            BackColor       =   &H00404040&
            Caption         =   "10.00 %"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   102
            Top             =   5490
            Width           =   2535
         End
         Begin VB.Label Label42 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   101
            Top             =   5025
            Width           =   2535
         End
         Begin VB.Label Label41 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   100
            Top             =   4560
            Width           =   2535
         End
         Begin VB.Label Label40 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   99
            Top             =   4095
            Width           =   2535
         End
         Begin VB.Label Label39 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   98
            Top             =   3645
            Width           =   2535
         End
         Begin VB.Label Label38 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   97
            Top             =   3180
            Width           =   2535
         End
         Begin VB.Label Label37 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   96
            Top             =   2715
            Width           =   2535
         End
         Begin VB.Label Label36 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   95
            Top             =   2250
            Width           =   2535
         End
         Begin VB.Label Label35 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   94
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label34 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   93
            Top             =   1785
            Width           =   2535
         End
         Begin VB.Label Label33 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   92
            Top             =   855
            Width           =   2535
         End
         Begin VB.Label Label32 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   91
            Top             =   390
            Width           =   2535
         End
         Begin VB.Label Label31 
            Caption         =   "TOTAL AMT:"
            Height          =   315
            Left            =   180
            TabIndex        =   90
            Top             =   8280
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "TOTAL EMI:"
            Height          =   315
            Left            =   180
            TabIndex        =   89
            Top             =   7800
            Width           =   1455
         End
         Begin VB.Label Label29 
            Caption         =   "CHEQUE AMT:"
            Height          =   315
            Left            =   180
            TabIndex        =   88
            Top             =   7350
            Width           =   1455
         End
         Begin VB.Label Label28 
            Caption         =   "NET PF %:"
            Height          =   315
            Left            =   180
            TabIndex        =   87
            Top             =   6886
            Width           =   1455
         End
         Begin VB.Label Label27 
            Caption         =   "DBD:"
            Height          =   315
            Left            =   180
            TabIndex        =   86
            Top             =   6422
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "CBD:"
            Height          =   315
            Left            =   180
            TabIndex        =   85
            Top             =   5958
            Width           =   1455
         End
         Begin VB.Label Label25 
            Caption         =   "ROI % :"
            Height          =   315
            Left            =   180
            TabIndex        =   84
            Top             =   5494
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "AD EMI AMT:"
            Height          =   315
            Left            =   180
            TabIndex        =   83
            Top             =   5030
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "AD EMI:"
            Height          =   315
            Left            =   180
            TabIndex        =   82
            Top             =   4566
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "EMI:"
            Height          =   315
            Left            =   180
            TabIndex        =   81
            Top             =   4102
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "TENOR:"
            Height          =   315
            Left            =   180
            TabIndex        =   80
            Top             =   3638
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "TOTAL DP:"
            Height          =   315
            Left            =   180
            TabIndex        =   79
            Top             =   3174
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "PF:"
            Height          =   315
            Left            =   180
            TabIndex        =   78
            Top             =   2710
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "D.P.:"
            Height          =   315
            Left            =   180
            TabIndex        =   77
            Top             =   2246
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "PRICE:"
            Height          =   315
            Left            =   180
            TabIndex        =   74
            Top             =   1318
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "LOAN"
            Height          =   315
            Left            =   180
            TabIndex        =   73
            Top             =   1782
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "COLOR:"
            Height          =   315
            Left            =   180
            TabIndex        =   72
            Top             =   854
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "MODEL:"
            Height          =   315
            Left            =   180
            TabIndex        =   71
            Top             =   390
            Width           =   1455
         End
      End
      Begin VB.ComboBox CboBankName 
         Height          =   360
         Left            =   -75030
         TabIndex        =   67
         Text            =   "Combo2"
         Top             =   5460
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox CboEmpName 
         Height          =   360
         Left            =   -75030
         TabIndex        =   66
         Text            =   "Combo1"
         Top             =   4830
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtNOtes 
         Height          =   360
         Left            =   -73380
         TabIndex        =   65
         Top             =   7380
         Width           =   6075
      End
      Begin VB.TextBox TxtEmail 
         Height          =   360
         Left            =   -73380
         TabIndex        =   64
         Top             =   4260
         Width           =   4215
      End
      Begin VB.TextBox TxtMobile 
         Height          =   360
         Left            =   -73380
         TabIndex        =   63
         Top             =   3630
         Width           =   2055
      End
      Begin VB.TextBox TxtPhone 
         Height          =   360
         Left            =   -73380
         TabIndex        =   62
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox TxtAddress 
         Height          =   360
         Left            =   -73380
         TabIndex        =   61
         Top             =   2370
         Width           =   6075
      End
      Begin VB.TextBox TxtName 
         Height          =   360
         Left            =   -73380
         TabIndex        =   60
         Top             =   1740
         Width           =   4215
      End
      Begin VB.TextBox TxtCode 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   -73380
         TabIndex        =   59
         Top             =   1110
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   345
         Left            =   -73380
         TabIndex        =   58
         Top             =   6135
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   38688
      End
      Begin VB.ComboBox CboBankCode 
         Height          =   360
         Left            =   -73380
         TabIndex        =   57
         Top             =   5520
         Width           =   2835
      End
      Begin VB.ComboBox CboEmpCode 
         Height          =   360
         Left            =   -73380
         TabIndex        =   56
         Top             =   4890
         Width           =   2835
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00808080&
         Caption         =   "&Print"
         Height          =   315
         Left            =   -63270
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   9150
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "C&lear"
         Height          =   315
         Left            =   -62250
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   9150
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "&Next >"
         Height          =   315
         Left            =   -61230
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   9150
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MsItem 
         Height          =   3735
         Left            =   -73290
         TabIndex        =   35
         Top             =   5670
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   6588
         _Version        =   393216
         ForeColor       =   16711935
         BackColorBkg    =   -2147483633
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.TextBox TxtCAmt 
         DataField       =   "PF"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -69390
         TabIndex        =   33
         Top             =   5160
         Width           =   1755
      End
      Begin VB.TextBox TxtDBD 
         DataField       =   "AdEmi"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   31
         Top             =   5175
         Width           =   1755
      End
      Begin VB.TextBox TxtCBD 
         DataField       =   "AdEmi"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   29
         Top             =   4710
         Width           =   1755
      End
      Begin VB.TextBox TxtTotalEmi 
         DataField       =   "PF"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -69390
         TabIndex        =   28
         Top             =   4230
         Width           =   1755
      End
      Begin VB.TextBox TxtDPay 
         DataField       =   "DPay"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   27
         Top             =   2880
         Width           =   1755
      End
      Begin VB.TextBox TxtTotalAmt 
         DataField       =   "PF"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -69390
         TabIndex        =   25
         Top             =   4695
         Width           =   1755
      End
      Begin VB.TextBox TxtNetPF 
         DataField       =   "PF"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -69390
         TabIndex        =   22
         Top             =   3315
         Width           =   1755
      End
      Begin VB.TextBox TXTAdEmi 
         DataField       =   "AdEmi"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   20
         Top             =   4260
         Width           =   1755
      End
      Begin VB.ComboBox CboType 
         Height          =   360
         ItemData        =   "FrmCustomerInfo.frx":009D
         Left            =   -73380
         List            =   "FrmCustomerInfo.frx":00AA
         TabIndex        =   19
         Top             =   6750
         Width           =   2835
      End
      Begin VB.ComboBox CboColor 
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   -73290
         TabIndex        =   17
         Top             =   1530
         Width           =   3615
      End
      Begin VB.ComboBox CboModel 
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   -73290
         TabIndex        =   16
         Top             =   1095
         Width           =   3615
      End
      Begin VB.TextBox TxtROI 
         DataField       =   "ROI"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -69390
         TabIndex        =   15
         Text            =   "10.00"
         Top             =   3780
         Width           =   1755
      End
      Begin VB.TextBox TxtEMI 
         DataField       =   "AdEmi"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   13
         Top             =   3810
         Width           =   1755
      End
      Begin VB.TextBox TxtTenor 
         DataField       =   "Tenor"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   11
         Top             =   3345
         Width           =   1755
      End
      Begin VB.TextBox TxtLAmt 
         DataField       =   "LAmt"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   9
         Top             =   2430
         Width           =   1755
      End
      Begin VB.TextBox TxtPF 
         DataField       =   "PF"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -69390
         TabIndex        =   6
         Top             =   2850
         Width           =   1755
      End
      Begin VB.TextBox TxtPrice 
         DataField       =   "Price"
         DataSource      =   "datPrimaryRS"
         Height          =   360
         Left            =   -73290
         TabIndex        =   4
         Top             =   1965
         Width           =   1755
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Caption         =   "Login Status:"
         Height          =   375
         Left            =   240
         TabIndex        =   119
         Top             =   690
         Width           =   1515
      End
      Begin VB.Label Label51 
         Caption         =   "No Of Copies"
         Height          =   375
         Left            =   -63330
         TabIndex        =   117
         Top             =   1170
         Width           =   1575
      End
      Begin VB.Label Label50 
         Caption         =   "Document Details"
         Height          =   375
         Left            =   -74790
         TabIndex        =   116
         Top             =   1170
         Width           =   3495
      End
      Begin VB.Label LblBankName 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -70530
         TabIndex        =   69
         Top             =   5520
         Width           =   4665
      End
      Begin VB.Label LblEmpName 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -70530
         TabIndex        =   68
         Top             =   4920
         Width           =   4665
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   55
         Top             =   7410
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   54
         Top             =   6780
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   53
         Top             =   6150
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Bank Code:"
         Height          =   315
         Left            =   -74640
         TabIndex        =   52
         Top             =   5520
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Emp Code:"
         Height          =   315
         Left            =   -74700
         TabIndex        =   51
         Top             =   4890
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   50
         Top             =   4260
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Mobile:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   49
         Top             =   3630
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   48
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   47
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   315
         Left            =   -74460
         TabIndex        =   46
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label LblTotalDP 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66270
         TabIndex        =   42
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Total DP:"
         Height          =   315
         Left            =   -67410
         TabIndex        =   41
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4125
         Left            =   -66480
         Top             =   4590
         Width           =   6105
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque Amount:"
         Height          =   255
         Index           =   16
         Left            =   -71190
         TabIndex        =   34
         Top             =   5220
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Dlr. BuyDown:"
         Height          =   255
         Index           =   15
         Left            =   -74910
         TabIndex        =   32
         Top             =   5220
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Co. BuyDown:"
         Height          =   255
         Index           =   14
         Left            =   -74910
         TabIndex        =   30
         Top             =   4770
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Amount:"
         Height          =   255
         Index           =   13
         Left            =   -70830
         TabIndex        =   26
         Top             =   4755
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Total EMI:"
         Height          =   255
         Index           =   12
         Left            =   -70500
         TabIndex        =   24
         Top             =   4290
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Net PF:"
         Height          =   255
         Index           =   11
         Left            =   -70500
         TabIndex        =   23
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "AdEmi:"
         Height          =   255
         Index           =   10
         Left            =   -74430
         TabIndex        =   21
         Top             =   4275
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Code:"
         Height          =   255
         Index           =   0
         Left            =   -74460
         TabIndex        =   18
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "ROI(%):"
         Height          =   255
         Index           =   9
         Left            =   -70530
         TabIndex        =   14
         Top             =   3825
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Emi:"
         Height          =   255
         Index           =   8
         Left            =   -74430
         TabIndex        =   12
         Top             =   3810
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tenor:"
         Height          =   255
         Index           =   7
         Left            =   -74400
         TabIndex        =   10
         Top             =   3375
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "LAmt:"
         Height          =   255
         Index           =   6
         Left            =   -74400
         TabIndex        =   8
         Top             =   2460
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DPay:"
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   7
         Top             =   2895
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "PF:"
         Height          =   255
         Index           =   4
         Left            =   -70500
         TabIndex        =   5
         Top             =   2910
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Price:"
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   3
         Top             =   2025
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Colour:"
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   2
         Top             =   1590
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Model:"
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   1
         Top             =   1155
         Width           =   1005
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Label12"
      Height          =   315
      Left            =   0
      TabIndex        =   76
      Top             =   330
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Label12"
      Height          =   315
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCustomerinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim JM As Integer

Private Sub CboBankCode_Click()
    If CboBankCode.Text <> "" Then
        CboBankName.ListIndex = CboBankCode.ListIndex
        LblBankName.Caption = CboBankName.Text
    End If
End Sub

Private Sub CboEmpCode_Click()
    If CboEmpCode.Text <> "" Then
        CboEmpName.ListIndex = CboEmpCode.ListIndex
        LblEmpName.Caption = CboEmpName.Text
    End If
End Sub

Private Sub CboModel_Click()
    If CboModel.Text <> "" Then
        CboColor.Clear
        RS.Open "SELECT COLOUR FROM COLOURMASTER WHERE  MODEL = '" & CboModel.Text & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        If RS.EOF = False And RS.BOF = False Then
            Do While Not RS.EOF
                CboColor.AddItem (RS(0))
                RS.MoveNext
            Loop
        End If
        RS.Close
        RS.Open "SELECT PRICE FROM VEHICALMASTER WHERE  MODEL = '" & CboModel.Text & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        If RS.EOF = False And RS.BOF = False Then
            TxtPrice.Text = RS(0)
        End If
        RS.Close
    End If
End Sub

Private Sub CmdAdd_Click()
    If Flg <> "add" Then
        Flg = "add"
        Call CtrlEnabled(Me)
        Call Enbl
        Call BlankText(Me)
        ItemName
        CboModel.SetFocus
    End If
End Sub

Private Sub CmdClose_Click()
    Unload Me
    Flg = ""
    Set RS = Nothing
    Set FrmCustomerinfo = Nothing
End Sub

Private Sub CmdSave_Click()
Dim Con1 As New ADODB.Connection
    If Flg = "add" Then
        Con1 = GETCON
        Con1.Open
        Con1.BeginTrans
        SSql = "INSERT INTO CUSTOMERINFO(CODE, IDATE, NAME, ADDRESS, PHONE, MOBILE, EMAIL, EMPCODE, BANKCODE, NOTES, TYPE, MODEL, COLOUR, PRICE, PF, DPAY, LAMT, TENOR, EMI, ADEMI, ROI, CDB,DBD, STATUS, DOCUMENT, USERID, SYSDATE) VALUES('" & TxtCode.Text & "','" & Format(DTP1.Value, "MM/dd/yyyy") & "','" & TxtName.Text & "','" & TxtAddress.Text & "','" & TxtPhone.Text & "','" & TxtMobile.Text & "','" & TxtEmail.Text & "', '" & CboEmpCode.Text & "','" & CboBankCode.Text & "','" & TxtNOtes.Text & "','" & CboType.Text & "','" & CboModel.Text & "', '" & CboColor.Text & "','" & TxtPrice.Text & "','" & TxtPF.Text & "','" & TxtDPay.Text & "','" & TxtLAmt.Text & "','" & TxtTenor.Text & "', '" & TxtEMI.Text & "','" & TXTAdEmi.Text & "','" & TxtROI.Text & "','" & TxtCBD.Text & "','" & TxtDBD.Text & "','" & Check1.Value & "','" & Check2.Value & "','" & UN & "', '" & Format(Date, "MM/dd/yyyy") & "')"
        Con1.Execute SSql
        Con1.CommitTrans
        Con1.Close
        If Check2.Value = 1 Then
            For I = 1 To MsDoc.Rows - 2
                Con1.Open
                Con1.BeginTrans
                SSql = "INSERT INTO DOCUMENT VALUES('" & TxtCode.Text & "','" & MsDoc.TextMatrix(I, 0) & "','" & MsDoc.TextMatrix(I, 1) & "','" & MsDoc.TextMatrix(I, 2) & "')"
                Con1.Execute SSql
                Con1.CommitTrans
                Con1.Close
            Next
        End If
        MsgBox "RECORD SAVE SUCESSFULLY", vbInformation, "FD"
        Flg = ""
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        CmdAdd.SetFocus
    End If
    If Flg = "edit" Then
        Con1 = GETCON
        Con1.Open
        Con1.BeginTrans
        SSql = "UPDATE CUSTOMERINFO SET IDATE='" & Format(DTP1.Value, "MM/dd/yyyy") & "', NAME='" & TxtName.Text & "', ADDRESS='" & TxtAddress.Text & "', PHONE='" & TxtPhone.Text & "', MOBILE='" & TxtMobile.Text & "', EMAIL='" & TxtEmail.Text & "', EMPCODE='" & CboEmpCode.Text & "', BANKCODE='" & CboBankCode.Text & "', NOTES='" & TxtNOtes.Text & "', TYPE='" & CboType.Text & "', MODEL='" & CboModel.Text & "', COLOUR='" & CboColor.Text & "', PRICE='" & TxtPrice.Text & "', PF='" & TxtPF.Text & "', DPAY='" & TxtDPay.Text & "', LAMT='" & TxtLAmt.Text & "', TENOR='" & TxtTenor.Text & "', EMI='" & TxtEMI.Text & "', ADEMI='" & TXTAdEmi.Text & "', ROI='" & TxtROI.Text & "', CDB='" & TxtCBD.Text & "',DBD='" & TxtDBD.Text & "', STATUS='" & Check1.Value & "', DOCUMENT='" & Check2.Value & "', LOGINSTATUS = '" & CboLoginStatus.Text & "', DELV = '" & Check3.Value & "',DELVDATE = '" & Format(DTPVD.Value, "MM/dd/yyyy") & "',SUPCODE = '" & CboSupCode.Text & "'," & _
               "RCBOOK = '" & TxtRCBook.Text & "',CNO = '" & TxtCNo.Text & "', ENGNO = '" & TxtEngNo.Text & "', PNO = '" & TxtPNo.Text & "', CAGR = '" & Check4.Value & "', FD = '" & Check5.Value & "', USERID='" & UN & "', SYSDATE='" & Format(Date, "MM/dd/yyyy") & "' WHERE CODE='" & TxtCode.Text & "'"
        Con1.Execute SSql
        Con1.CommitTrans
        Con1.Close
        If Check2.Value = 1 Then
            For I = 1 To MsDoc.Rows - 2
                SSql = "SELECT COUNT(*) FROM DOCUMENT WHERE CODE = '" & TxtCode.Text & "' AND INO = " & MsDoc.TextMatrix(I, 0) & ""
                RS.Open SSql, GETCON, adOpenForwardOnly, adLockReadOnly
                If RS.EOF = False And RS.BOF = False Then
                    Con1.Open
                    Con1.BeginTrans
                    SSql = "UPDATE DOCUMENT SET TYPE = '" & MsDoc.TextMatrix(I, 1) & "',COPY = '" & MsDoc.TextMatrix(I, 2) & "' WHERE CODE = '" & TxtCode.Text & "' AND INO = " & MsDoc.TextMatrix(I, 0) & ""
                    Con1.Execute SSql
                    Con1.CommitTrans
                    Con1.Close
                Else
                    Con1.Open
                    Con1.BeginTrans
                    SSql = "INSERT INTO DOCUMENT VALUES('" & TxtCode.Text & "','" & MsDoc.TextMatrix(I, 0) & "','" & MsDoc.TextMatrix(I, 1) & "','" & MsDoc.TextMatrix(I, 2) & "')"
                    Con1.Execute SSql
                    Con1.CommitTrans
                    Con1.Close
                End If
                RS.Close
            Next
        End If
        If Check5.Value = 1 Then
            For I = 1 To MsDoc1.Rows - 2
                RS.Open "SELECT COUNT(*) FROM FDOCUMENT WHERE CODE = '" & TxtCode.Text & "' AND INO = " & MsDoc1.TextMatrix(I, 0) & "", GETCON, adOpenForwardOnly, adLockReadOnly
                If RS.EOF = False And RS.BOF = False Then
                    Con1.Open
                    Con1.BeginTrans
                    SSql = "UPDATE FDOCUMENT SET TYPE = '" & MsDoc1.TextMatrix(I, 1) & "',COPY = '" & MsDoc1.TextMatrix(I, 2) & "' WHERE CODE = '" & TxtCode.Text & "' AND INO = " & MsDoc1.TextMatrix(I, 0) & ""
                    Con1.Execute SSql
                    Con1.CommitTrans
                    Con1.Close
                Else
                    Con1.Open
                    Con1.BeginTrans
                    SSql = "INSERT INTO FDOCUMENT VALUES('" & TxtCode.Text & "','" & MsDoc1.TextMatrix(I, 0) & "','" & MsDoc1.TextMatrix(I, 1) & "','" & MsDoc1.TextMatrix(I, 2) & "')"
                    Con1.Execute SSql
                    Con1.CommitTrans
                    Con1.Close
                End If
                RS.Close
            Next
        End If
        MsgBox "RECORD UPDATE SUCESSFULLY", vbInformation, "FD"
        Flg = ""
        Call CtrlDisabled(Me)
        Call BlankText(Me)
        ItemName
        CmdSave.Caption = "&Save"
        CmdAdd.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    Dim R As String
    STab.Tab = 1
    CboEmpCode.Clear
    CboEmpName.Clear
    RS.Open "SELECT EMPCODE, EMPNAME FROM EMPMASTER ORDER BY EMPCODE", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            CboEmpCode.AddItem (RS(0))
            CboEmpName.AddItem (RS(1))
            RS.MoveNext
        Loop
    End If
    RS.Close
    RS.Open "SELECT DCode, DNAME FROM SupDealMaster WHERE Type = 'SUPPLIER' ORDER BY DCODE", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            CboEmpCode.AddItem (RS(0))
            CboEmpName.AddItem (RS(1))
            RS.MoveNext
        Loop
    End If
    RS.Close
    CboBankCode.Clear
    CboBankName.Clear
    RS.Open "SELECT BCODE, BNAME FROM BANKMASTER ORDER BY BCODE", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            CboBankCode.AddItem (RS(0))
            CboBankName.AddItem (RS(1))
            RS.MoveNext
        Loop
    End If
    RS.Close
    Label32.Caption = " " & CboModel.Text
    Label33.Caption = " " & CboColor.Text
    Label34.Caption = " " & TxtPrice.Text
    Label35.Caption = " " & TxtLAmt.Text
    Label36.Caption = " " & TxtDPay.Text
    Label37.Caption = " " & TxtPF.Text
    Label38.Caption = " " & LblTotalDP.Caption
    Label39.Caption = " " & TxtTenor.Text
    Label40.Caption = " " & TxtEMI.Text
    Label41.Caption = " " & TXTAdEmi.Text
    Label42.Caption = " " & Val(TXTAdEmi.Text) * Val(TxtEMI.Text)
    Label43.Caption = " " & TxtROI.Text
    Label44.Caption = " " & TxtCBD.Text
    Label45.Caption = " " & TxtDBD.Text
    Label46.Caption = " " & TxtNetPF.Text
    Label47.Caption = " " & TxtCAmt.Text
    Label48.Caption = " " & TxtTotalEmi.Text
    Label49.Caption = " " & TxtTotalAmt.Text
End Sub

Private Sub Command2_Click()
    CboColor.Text = ""
    CboModel.Text = ""
    TxtPF.Text = ""
    TXTAdEmi.Text = ""
    TxtCAmt.Text = ""
    TxtCBD.Text = ""
    TxtDBD.Text = ""
    TxtDPay.Text = ""
    TxtEMI.Text = ""
    TxtLAmt.Text = ""
    TxtNetPF.Text = ""
    TxtPrice.Text = ""
    TxtROI.Text = "10.00"
    TxtTenor.Text = ""
    TxtTotalAmt.Text = ""
    TxtTotalEmi.Text = ""
    Image1.Picture = LoadPicture("")
    TxtLAmt.Text = ""
    TxtDPay.Text = ""
    CboModel.SetFocus
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim K, J, JJ As Integer
    Dim A As Double
    Dim Z As String
    
    Dim LName As String
        LName = "Cus_Scheme"
        Set xlapp = New Excel.Application
        xlapp.Workbooks.Open App.Path & "\format\Scheme_Format"
        Set wksSheet = xlapp.Worksheets("Scheme")
        wksSheet.Activate
        wksSheet.SaveAs App.Path & "\" & LName & ".xls"
        xlapp.ActiveWorkbook.Close savechanges:=False, FileName:=App.Path & "\format\Scheme_Format"
        Set wksSheet = Nothing
        xlapp.Quit
        Set xlapp = Nothing
        
        ''''''''''''''''''''''''''Insert Data Into Excel
        Set XLA = New Excel.Application
        XLA.Workbooks.Open App.Path & "\" & LName & ".xls"
        Set XLS = XLA.Worksheets("Scheme")
        XLS.Activate
        
        XLS.Cells(4, 2) = CboModel.Text
        XLS.Cells(5, 2) = CboColor.Text
        XLS.Cells(6, 2) = TxtPrice.Text
        XLS.Cells(7, 2) = TxtDPay.Text
        XLS.Cells(8, 2) = TxtLAmt.Text
        XLS.Cells(7, 5) = TxtPF.Text
        XLS.Cells(9, 2) = TxtROI.Text
        XLS.Cells(10, 2) = TxtTenor.Text
        XLS.Cells(11, 2) = TxtEMI.Text
        XLS.Cells(12, 2) = TXTAdEmi.Text
        XLS.Cells(12, 3) = Val(TxtEMI.Text) * Val(TXTAdEmi.Text)
        XLS.Cells(14, 2) = TxtCBD.Text
        XLS.Cells(15, 2) = TxtDBD.Text
        XLS.Cells(8, 5) = LblTotalDP.Caption
        XLS.Cells(10, 5) = TxtNetPF.Text
        XLS.Cells(12, 5) = TxtTotalAmt.Text
        XLS.Cells(11, 5) = TxtTotalEmi.Text
        XLS.Cells(13, 5) = TxtCAmt.Text
        
        XLA.ActiveWorkbook.Close savechanges:=True
        Set XLW = Nothing
        XLA.Quit
        Set XLA = Nothing
        MsgBox "Create file sucessfully", vbInformation, "FD"
End Sub

Private Sub Command4_Click()
    STab.Tab = 2
End Sub

Private Sub Command6_Click()
    If TxtDoc.Text <> "" And Check1.Value = 1 Then
        MsDoc.TextMatrix(JM, o) = JM
        MsDoc.TextMatrix(JM, 1) = TxtDoc.Text
        MsDoc.TextMatrix(JM, 2) = TxtCopy.Text
        MsDoc.Rows = MsDoc.Rows + 1
        JM = JM + 1
        TxtDoc.Text = ""
        TxtCopy.Text = ""
        TxtDoc.SetFocus
    End If
End Sub

Private Sub Form_Load()
'On Error Resume Next
    Me.Top = 200
    Me.Left = 250
    Call CtrlDisabled(Me)
    Call BlankText(Me)
    RS.Open "SELECT MODEL FROM VehicalMaster", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            CboModel.AddItem (RS(0))
            RS.MoveNext
        Loop
    End If
    RS.Close
    ItemName
    CmdSave.Caption = "&Save"
    Check1.Value = 1
    JM = 1
End Sub

Private Sub MsDoc_DblClick()
    Dim RA As String
    Dim E, JJ As Integer
    JJ = 1
    RA = MsgBox("You want remove this row", vbYesNo, "Vinesh")
    If RA = vbYes Then
        MsDoc.RemoveItem (MsDoc.Row)
        JM = JM - 1
        For E = 1 To MsDoc.Rows - 1
            MsDoc.TextMatrix(E, 0) = JJ
            JJ = JJ - 1
        Next
    End If
End Sub

Private Sub TXTAdEmi_Change()
    On Error Resume Next
    LblTotalDP.Caption = Round(Val(TxtDPay.Text) + Val(TxtPF.Text) + (Val(TXTAdEmi.Text) * Val(TxtEMI.Text)), 0)
    TxtNetPF.Text = Format(Round((Val(TxtPF.Text) * 100) / (Val(TxtPrice.Text) - Val(TxtDPay.Text) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text))), 2), "0.00")
    TxtCAmt.Text = Val(TxtPrice.Text) - (Val(TxtDPay.Text) + Val(TxtPF.Text)) - Val(TxtCBD.Text) - Val(TxtDBD.Text)
    TxtTotalEmi.Text = Round(Val(TxtTenor.Text) * Val(TxtEMI.Text), 0)
    TxtTotalAmt.Text = Val(LblTotalDP.Caption) + Val(TxtTotalEmi.Text) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text))
    ItemLoad
End Sub

Private Sub TxtCBD_Change()
    TxtCAmt.Text = Val(TxtPrice.Text) - (Val(TxtDPay.Text) + Val(TxtPF.Text)) - Val(TxtCBD.Text) - Val(TxtDBD.Text)
End Sub

Private Sub TxtDBD_Change()
    TxtCAmt.Text = Val(TxtPrice.Text) - (Val(TxtDPay.Text) + Val(TxtPF.Text)) - Val(TxtCBD.Text) - Val(TxtDBD.Text)
End Sub

Private Sub TxtDPay_Change()
On Error Resume Next
    TxtPF.Text = Round(Val(TxtLAmt.Text) * (2.25 / 100) + 220, 0)
    'LblTotalDP.Caption = Round(Val(TxtDPay.Text) + Val(TxtPF.Text) + (Val(TxtTenor.Text) * Val(TxtEMI.Text)), 0)
    TxtLAmt.Text = Round(Val(TxtPrice.Text) - Val(TxtDPay.Text), 0)
    TxtEMI.Text = Format(((Val(TxtLAmt.Text) * Val(TxtROI.Text) * Val(TxtTenor.Text) / 1200) + Val(TxtLAmt.Text)) / Val(TxtTenor.Text), "0.00")
    LblTotalDP.Caption = Round(Val(TxtDPay.Text) + Val(TxtPF.Text) + (Val(TXTAdEmi.Text) * Val(TxtEMI.Text)), 0)
    TxtNetPF.Text = Format(Round((Val(TxtPF.Text) * 100) / (Val(TxtPrice.Text) - Val(TxtDPay.Text) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text))), 2), "0.00")
    TxtCAmt.Text = Val(TxtPrice.Text) - (Val(TxtDPay.Text) + Val(TxtPF.Text)) - Val(TxtCBD.Text) - Val(TxtDBD.Text)
    TxtTotalEmi.Text = Round((Val(TxtTenor.Text) * Val(TxtEMI.Text)) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text)), 0)
    TxtTotalAmt.Text = Val(LblTotalDP.Caption) + Val(TxtTotalEmi.Text)
End Sub

Private Sub TxtLAmt_Change()
On Error Resume Next
    TxtPF.Text = Round(Val(TxtLAmt.Text) * (2.25 / 100) + 220, 0)
    TxtDPay.Text = Round((Val(TxtPrice.Text) - Val(TxtLAmt.Text)), 0)
    TxtEMI.Text = Format(((Val(TxtLAmt.Text) * Val(TxtROI.Text) * Val(TxtTenor.Text) / 1200) + Val(TxtLAmt.Text)) / Val(TxtTenor.Text), "0.00")
    LblTotalDP.Caption = Round(Val(TxtDPay.Text) + Val(TxtPF.Text) + (Val(TXTAdEmi.Text) * Val(TxtEMI.Text)), 0)
    TxtNetPF.Text = Format(Round((Val(TxtPF.Text) * 100) / (Val(TxtPrice.Text) - Val(TxtDPay.Text) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text))), 2), "0.00")
    TxtCAmt.Text = Val(TxtPrice.Text) - (Val(TxtDPay.Text) + Val(TxtPF.Text)) - Val(TxtCBD.Text) - Val(TxtDBD.Text)
    TxtTotalEmi.Text = Round((Val(TxtTenor.Text) * Val(TxtEMI.Text)) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text)), 0)
    TxtTotalAmt.Text = Val(LblTotalDP.Caption) + Val(TxtTotalEmi.Text)
End Sub

Private Sub TxtName_LostFocus()
    If TxtName.Text <> "" Then 'And Flg = "add"
        RS.Open "SELECT COUNT(CODE) FROM CUSTOMERINFO WHERE LEFT(CODE,3) = '" & Left(TxtName.Text, 3) & "'", GETCON, adOpenForwardOnly, adLockReadOnly
        If RS.EOF = True And RS.BOF = True Then
            TxtCode.Text = UCase(Left(TxtName.Text, 3)) & Format(1, "0000")
        Else
            TxtCode.Text = UCase(Left(TxtName.Text, 3)) & Format(RS(0) + 1, "0000")
        End If
        RS.Close
    End If
End Sub

Private Sub TxtTenor_Change()
On Error Resume Next
    TxtEMI.Text = Format(((Val(TxtLAmt.Text) * Val(TxtROI.Text) * Val(TxtTenor.Text) / 1200) + Val(TxtLAmt.Text)) / Val(TxtTenor.Text), "0.00")
    LblTotalDP.Caption = Round(Val(TxtDPay.Text) + Val(TxtPF.Text) + (Val(TXTAdEmi.Text) * Val(TxtEMI.Text)), 0)
    TxtNetPF.Text = Format(Round((Val(TxtPF.Text) * 100) / (Val(TxtPrice.Text) - Val(TxtDPay.Text) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text))), 2), "0.00")
    TxtCAmt.Text = Val(TxtPrice.Text) - (Val(TxtDPay.Text) + Val(TxtPF.Text)) - Val(TxtCBD.Text) - Val(TxtDBD.Text)
    TxtTotalEmi.Text = Round((Val(TxtTenor.Text) * Val(TxtEMI.Text)) - (Val(TxtEMI.Text) * Val(TXTAdEmi.Text)), 0)
    TxtTotalAmt.Text = Val(LblTotalDP.Caption) + Val(TxtTotalEmi.Text)
    ItemLoad
End Sub

Private Sub ItemLoad()
    Dim J As Integer
    Dim I As Integer
    MsItem.Rows = 2
    MsItem.TextMatrix(1, 0) = "0"
    MsItem.TextMatrix(1, 1) = Format(-(Val(TxtLAmt.Text) + (Val(TxtEMI.Text) * Val(TXTAdEmi.Text)) + Val(TxtCBD.Text) + Val(TxtDBD.Text)), "0.00")
    MsItem.Rows = MsItem.Rows + 1
    J = Val(TxtTenor.Text) - Val(TXTAdEmi.Text)
    For I = 1 To J
        MsItem.TextMatrix(I + 1, 0) = I
        MsItem.TextMatrix(I + 1, 1) = TxtEMI.Text
        MsItem.Rows = MsItem.Rows + 1
    Next
End Sub

Public Sub VIEWDATA(A As Variant)
On Error Resume Next
Dim RS1 As New ADODB.Recordset
    RS.Open "SELECT * FROM CUSTOMERINFO WHERE CODE = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False And RS.BOF = False Then
        TxtCode.Text = RS(0) & ""
        DTP1.Value = RS(1) & ""
        TxtName.Text = RS(2) & ""
        TxtAddress.Text = RS(3) & ""
        TxtPhone.Text = RS(4) & ""
        TxtMobile.Text = RS(5) & ""
        TxtEmail.Text = RS(6) & ""
        TxtNOtes.Text = RS(9) & ""
        CboType.Text = RS(10) & ""
        CboModel.Text = RS(11) & ""
        CboColor.Text = RS(12) & ""
        TxtPrice.Text = RS(13) & ""
        TxtPF.Text = RS(14) & ""
        TxtDPay.Text = RS(15) & ""
        TxtLAmt.Text = RS(16) & ""
        TxtTenor.Text = RS(17) & ""
        TxtEMI.Text = RS(18) & ""
        TXTAdEmi.Text = RS(19) & ""
        TxtROI.Text = RS(20) & ""
        TxtCBD.Text = RS(21) & ""
        TxtDBD.Text = RS(22) & ""
        Check1.Value = RS(23) & ""
        Check2.Value = RS(24) & ""
        CboLoginStatus.Text = RS(25) & ""
        Check3.Value = RS(26) & ""
        DTPVD.Value = RS(27) & ""
        CboSupCode.Text = RS(28) & ""
        TxtRCBook.Text = RS(29) & ""
        TxtCNo.Text = RS(30) & ""
        TxtEngNo.Text = RS(31) & ""
        TxtPNo.Text = RS(32) & ""
        Check4.Value = RS(33) & ""
        Check5.Value = RS(34) & ""
        If CboLoginStatus.Text = "Approval" Then
            Frame2.Visible = True
            Frame3.Visible = True
        Else
            Frame3.Visible = False
            Frame2.Visible = False
        End If
        I = 1
        MsDoc.Clear
        MsDoc1.Clear
        ItemName
        MsDoc.Rows = 2
        MsDoc1.Rows = 2
        If RS(24) = 1 Then
            RS1.Open "SELECT * FROM DOCUMENT WHERE CODE = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
            If RS1.EOF = False And RS1.BOF = False Then
                Do While Not RS1.EOF
                    MsDoc.TextMatrix(I, 0) = RS1(1) & ""
                    MsDoc.TextMatrix(I, 1) = RS1(2) & ""
                    MsDoc.TextMatrix(I, 2) = RS1(3) & ""
                    MsDoc.Rows = MsDoc.Rows + 1
                    I = I + 1
                    RS1.MoveNext
                Loop
            End If
            RS1.Close
        End If
        I = 1
        If RS(34) = 1 Then
            RS1.Open "SELECT * FROM FDOCUMENT WHERE CODE = '" & A & "'", GETCON, adOpenForwardOnly, adLockReadOnly
            If RS1.EOF = False And RS1.BOF = False Then
                Do While Not RS1.EOF
                    MsDoc1.TextMatrix(I, 0) = RS1(1) & ""
                    MsDoc1.TextMatrix(I, 1) = RS1(2) & ""
                    MsDoc1.TextMatrix(I, 2) = RS1(3) & ""
                    MsDoc1.Rows = MsDoc1.Rows + 1
                    I = I + 1
                    RS1.MoveNext
                Loop
            End If
            RS1.Close
        End If
        Command1_Click
        CboEmpCode.Text = RS(7) & ""
        CboBankCode.Text = RS(8) & ""
    End If
    RS.Close
    Call CtrlEnabled(Me)
    Call Enbl
    Flg = "edit"
    CmdSave.Caption = "&Edit"
    STab.Tab = 0
    CboModel.SetFocus
End Sub

Private Sub ItemName()
    MsItem.TextMatrix(0, 0) = "No"
    MsItem.ColAlignment(0) = 3
    MsItem.ColWidth(0) = 1000
    MsItem.TextMatrix(0, 1) = "Amount"
    MsItem.ColAlignment(1) = 3
    MsItem.ColWidth(1) = 1500
    
    MsDoc.TextMatrix(0, 0) = "Sr No"
    MsDoc.TextMatrix(0, 1) = "Document Details"
    MsDoc.TextMatrix(0, 2) = "No Of Copy"
    MsDoc.ColAlignment(0) = 3
    MsDoc.ColAlignment(2) = 3
    MsDoc.FixedAlignment(0) = 3
    MsDoc.FixedAlignment(1) = 3
    MsDoc.FixedAlignment(2) = 3
    MsDoc.ColWidth(1) = 8500
    MsDoc.ColWidth(2) = 1500
    
    MsDoc1.TextMatrix(0, 0) = "Sr No"
    MsDoc1.TextMatrix(0, 1) = "Final Document Details"
    MsDoc1.TextMatrix(0, 2) = "No Of Copy"
    MsDoc1.ColAlignment(0) = 3
    MsDoc1.ColAlignment(2) = 3
    MsDoc1.FixedAlignment(0) = 3
    MsDoc1.FixedAlignment(1) = 3
    MsDoc1.FixedAlignment(2) = 3
    MsDoc1.ColWidth(1) = 8500
    MsDoc1.ColWidth(2) = 1500
End Sub

Private Sub Enbl()
    TxtCode.Enabled = False
    TxtPrice.Enabled = False
    TxtEMI.Enabled = False
    TxtPF.Enabled = False
    TxtNetPF.Enabled = False
    TxtROI.Enabled = False
    TxtTotalEmi.Enabled = False
    TxtTotalAmt.Enabled = False
    TxtCAmt.Enabled = False
End Sub
