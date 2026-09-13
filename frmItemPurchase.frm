VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmItemPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15240
   Begin TabDlg.SSTab SSTab2 
      Height          =   1335
      Left            =   120
      TabIndex        =   72
      Top             =   1560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "By Issue Unit"
      TabPicture(0)   =   "frmItemPurchase.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblIssueUnit(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblIssueUnit(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtLastPPI"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtLastSPI"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "By Pack Units"
      TabPicture(1)   =   "frmItemPurchase.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label57"
      Tab(1).Control(1)=   "Label51"
      Tab(1).Control(2)=   "lblPackUnit(4)"
      Tab(1).Control(3)=   "lblPackUnit(5)"
      Tab(1).Control(4)=   "txtLastPPP"
      Tab(1).Control(5)=   "txtLastSPP"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtLastSPI 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   80
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtLastPPI 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   79
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtLastSPP 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73080
         TabIndex        =   76
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtLastPPP 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73080
         TabIndex        =   75
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblPackUnit 
         Caption         =   "Pack Unit"
         Height          =   375
         Index           =   5
         Left            =   -71520
         TabIndex        =   90
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblPackUnit 
         Caption         =   "Pack Unit"
         Height          =   375
         Index           =   4
         Left            =   -71520
         TabIndex        =   89
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblIssueUnit 
         Caption         =   "Issue Unit"
         Height          =   375
         Index           =   5
         Left            =   3480
         TabIndex        =   85
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblIssueUnit 
         Caption         =   "Issue Unit"
         Height          =   375
         Index           =   4
         Left            =   3480
         TabIndex        =   84
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Last Sale Price"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Last Purchase Price"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label51 
         Caption         =   "Last Sale Price"
         Height          =   255
         Left            =   -74880
         TabIndex        =   74
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label57 
         Caption         =   "Last Purchase Price"
         Height          =   255
         Left            =   -74880
         TabIndex        =   73
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox txtSaleValue 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   9024
      Width           =   3255
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   9480
      Width           =   3255
   End
   Begin VB.TextBox txtBillID 
      Height          =   375
      Left            =   1800
      TabIndex        =   67
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDeleteID 
      Height          =   375
      Left            =   960
      TabIndex        =   66
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPaymentComments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   47
      Top             =   8088
      Width           =   3255
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   12360
      TabIndex        =   54
      Top             =   8085
      Width           =   2535
   End
   Begin VB.TextBox txtInvoice 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1440
      TabIndex        =   38
      Top             =   8040
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1440
      TabIndex        =   40
      Top             =   8520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   131072003
      CurrentDate     =   39691
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   5040
      TabIndex        =   59
      Top             =   120
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "By Issue Units"
      TabPicture(0)   =   "frmItemPurchase.frx":0038
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtSPrice"
      Tab(0).Control(1)=   "txtPPrice"
      Tab(0).Control(2)=   "txtQty"
      Tab(0).Control(3)=   "txtFQty"
      Tab(0).Control(4)=   "lblIssueUnit(3)"
      Tab(0).Control(5)=   "lblIssueUnit(2)"
      Tab(0).Control(6)=   "lblIssueUnit(1)"
      Tab(0).Control(7)=   "lblIssueUnit(0)"
      Tab(0).Control(8)=   "Label56"
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(10)=   "Label32"
      Tab(0).Control(11)=   "Label54"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "By Pack Units"
      TabPicture(1)   =   "frmItemPurchase.frx":0054
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label52"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label53"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label58"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblPackUnit(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblPackUnit(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblPackUnit(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblPackUnit(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtFPQty"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtPQty"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtPPPrice"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtPSPrice"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.TextBox txtPSPrice 
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtSPrice 
         Height          =   375
         Left            =   -73200
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtPPPrice 
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtPQty 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFPQty 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtPPrice 
         Height          =   375
         Left            =   -73200
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtQty 
         Height          =   375
         Left            =   -73200
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFQty 
         Height          =   375
         Left            =   -73200
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblPackUnit 
         Caption         =   "Pack Unit"
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   88
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblPackUnit 
         Caption         =   "Pack Unit"
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   87
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblPackUnit 
         Caption         =   "Pack Unit"
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   86
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblIssueUnit 
         Caption         =   "Issue Unit"
         Height          =   375
         Index           =   3
         Left            =   -71400
         TabIndex        =   83
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblIssueUnit 
         Caption         =   "Issue Unit"
         Height          =   375
         Index           =   2
         Left            =   -71400
         TabIndex        =   82
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblIssueUnit 
         Caption         =   "Issue Unit"
         Height          =   375
         Index           =   1
         Left            =   -71400
         TabIndex        =   81
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblIssueUnit 
         Caption         =   "Issue Unit"
         Height          =   375
         Index           =   0
         Left            =   -71400
         TabIndex        =   71
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblPackUnit 
         Caption         =   "Pack Unit"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   70
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "&Sale Price"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label56 
         Caption         =   "&Sale Price"
         Height          =   375
         Left            =   -74760
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label58 
         Caption         =   "P&urchase Price"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label53 
         Caption         =   "&Quantity"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label52 
         Caption         =   "&Free Quantity"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "P&urchase Price"
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label32 
         Caption         =   "&Quantity"
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label54 
         Caption         =   "&Free Quantity"
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   10440
      TabIndex        =   55
      Top             =   9600
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtBatch 
      Height          =   375
      Left            =   12600
      TabIndex        =   23
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtPurchaseValue 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtEditID 
      Height          =   375
      Left            =   120
      TabIndex        =   58
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin btButtonEx.ButtonEx btnPurchase 
      Height          =   375
      Left            =   11880
      TabIndex        =   56
      Top             =   9600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Purchase"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4455
      Left            =   120
      TabIndex        =   33
      Top             =   3000
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7858
      _Version        =   393216
      WordWrap        =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   13560
      TabIndex        =   57
      Top             =   9600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCode 
      Height          =   360
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDOM 
      Height          =   375
      Left            =   12600
      TabIndex        =   25
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   130678787
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker dtpDOE 
      Height          =   375
      Left            =   12600
      TabIndex        =   27
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   130678787
      CurrentDate     =   39545
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   13800
      TabIndex        =   32
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   13800
      TabIndex        =   34
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbChecked 
      Height          =   360
      Left            =   1440
      TabIndex        =   42
      Top             =   9000
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbReceived 
      Height          =   360
      Left            =   1440
      TabIndex        =   44
      Top             =   9480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   360
      Left            =   1440
      TabIndex        =   36
      Top             =   7560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPayment 
      Height          =   360
      Left            =   6480
      TabIndex        =   60
      Top             =   7620
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtBatchID 
      Height          =   375
      Left            =   2640
      TabIndex        =   91
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   6480
      TabIndex        =   49
      Top             =   8556
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDepartment 
      Height          =   360
      Left            =   6840
      TabIndex        =   92
      Top             =   2595
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label15 
      Caption         =   "Department"
      Height          =   255
      Left            =   5040
      TabIndex        =   93
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Drawer"
      Height          =   255
      Left            =   4920
      TabIndex        =   48
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Sale Value"
      Height          =   375
      Left            =   10800
      TabIndex        =   30
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Total Sale Value"
      Height          =   255
      Left            =   10440
      TabIndex        =   69
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label lblTotalSaleValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   12360
      TabIndex        =   68
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Printer"
      Height          =   255
      Left            =   4920
      TabIndex        =   50
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "PPaper"
      Height          =   255
      Left            =   4920
      TabIndex        =   52
      Top             =   9480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Paymen&t Details"
      Height          =   255
      Left            =   4920
      TabIndex        =   46
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Purchase &Date"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "&Discount"
      Height          =   255
      Left            =   10440
      TabIndex        =   65
      Top             =   8085
      Width           =   1215
   End
   Begin VB.Label lblNetTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   12360
      TabIndex        =   62
      Top             =   8595
      Width           =   2535
   End
   Begin VB.Label lblGrossTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   12360
      TabIndex        =   61
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Label Label23 
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   4920
      TabIndex        =   45
      Top             =   7620
      Width           =   1935
   End
   Begin VB.Label Label22 
      Caption         =   "Invoice &No."
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "C&hecked by"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "&Received by"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Label Label24 
      Caption         =   "&Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   7620
      Width           =   1455
   End
   Begin VB.Label Label38 
      Caption         =   "&Batch"
      Height          =   375
      Left            =   10800
      TabIndex        =   22
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label37 
      Caption         =   "Date of &Manufacture"
      Height          =   375
      Left            =   10800
      TabIndex        =   24
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "Date of &Expiary"
      Height          =   375
      Left            =   10800
      TabIndex        =   26
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label35 
      Caption         =   "C&ode"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "&Item"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label30 
      Caption         =   "Purchase Value"
      Height          =   375
      Left            =   10800
      TabIndex        =   28
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "&Catogery"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Net Total"
      Height          =   255
      Left            =   10440
      TabIndex        =   64
      Top             =   8595
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   10440
      TabIndex        =   63
      Top             =   7560
      Width           =   1215
   End
End
Attribute VB_Name = "frmItemPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NumForms As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    
    Dim FSys As New Scripting.FileSystemObject
    
    Private CsetPrinter As New cSetDfltPrinter
    
    Dim temSql As String
    Dim i As Integer
    
    Dim DisplayChange As Boolean
    
    Dim MyBatch As ItemBatch

    Dim FormPurchaseBill As New clsPurchaseBill
    Dim FormPurchaseItem As New clsItem
    
    Dim FocusControl As Control

Private Sub FillCombos()
    Dim Item As New clsFillCombos
    Item.FillBoolCombo cmbItem, "Item", "Item", "IsItemName", True
    
    Dim ICode As New clsFillCombos
    ICode.FillSpecificField cmbCode, "Item", "ItemCode", True
    
    Dim ICat As New clsFillCombos
    ICat.FillBoolCombo cmbCategory, "Category", "Category", "IsItemCategory", True
    
    Dim Sup As New clsFillCombos
    Sup.FillBoolCombo cmbSupplier, "Institution", "Institution", "IsSupplier", True
    
    Dim Pay As New clsFillCombos
    Pay.FillBoolCombo cmbPayment, "PaymentMethod", "PaymentMethod", "ToPay", True
    
    Dim CheckedUser As New clsFillCombos
    CheckedUser.FillAnyCombo cmbChecked, "Staff", True
    
    Dim ReceivedUser As New clsFillCombos
    ReceivedUser.FillAnyCombo cmbReceived, "Staff", True
        
    Dim Drawer As New clsFillCombos
    Drawer.FillAnyCombo cmbDrawer, "Drawer", True


    Dim Dept As New clsFillCombos
    Dept.FillAnyCombo cmbDepartment, "Department", True

End Sub

Private Sub btnAdd_Click()
    Dim temBool As Boolean
    If IsNumeric(txtBillID.Text) = False Then
        txtBillID.Text = FormPurchaseBill.NewPurchaseBillID
    End If
    If Trim(txtBatch.Text) = Empty Then
        txtBatch.Text = "Batch No. Not Known"
    End If
    If IsNumeric(cmbDepartment.BoundText) = False Then
        MsgBox "Please select a Department"
        cmbDepartment.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Please select an Item"
        cmbItem.SetFocus
        Exit Sub
    Else
        If Val(cmbItem.BoundText) <> FormPurchaseItem.ID Then FormPurchaseItem.ID = Val(cmbItem.BoundText)
    End If
    If Val(txtQty.Text) = 0 Then
        MsgBox "Please select a quentity"
        If SSTab1.Tab = 0 Then
            txtQty.SetFocus
        Else
            txtPQty.SetFocus
        End If
        Exit Sub
    End If
    If Val(txtPurchaseValue.Text) = 0 Then
        MsgBox "Please select a purchase price"
        If SSTab1.Tab = 0 Then
            txtPPrice.SetFocus
        Else
            txtPPPrice.SetFocus
        End If
        Exit Sub
    End If
    Dim rsTemBatch As New ADODB.Recordset
    
    Dim rsTemPurchaseItem As New ADODB.Recordset
    With rsTemPurchaseItem
        If .State = 1 Then .Close
        If Val(txtEditID.Text) <> 0 Then
            temSql = "Select * from tblPurchaseItem where PurchaseItemID = " & Val(txtEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If !ItemID <> FormPurchaseItem.ID Then
                temBool = ChangeStocks(!ItemID, ItemBatchID(!ItemID, UCase(txtBatch.Text), dtpDOE.Value, dtpDOM.Value), Val(!PurchaseQuentity) + Val(!PurchaseFreeQuentity), !DepartmentID, False, True, False)
                temBool = ChangeStocks(FormPurchaseItem.ID, ItemBatchID(FormPurchaseItem.ID, UCase(txtBatch.Text), dtpDOE.Value, dtpDOM.Value), Val(txtQty.Text) + Val(txtFQty.Text), Val(cmbDepartment.BoundText), True, False, False)
            Else
                If !ItemBatchID = ItemBatchID(FormPurchaseItem.ID, UCase(txtBatch.Text), dtpDOE.Value, dtpDOM.Value) Then
                    temBool = ChangeStocks(FormPurchaseItem.ID, !ItemBatchID, !PurchaseQuentity + !PurchaseFreeQuentity, !DepartmentID, False, True, False)
                    temBool = ChangeStocks(FormPurchaseItem.ID, !ItemBatchID, Val(txtQty.Text) + Val(txtFQty.Text), cmbDepartment.BoundText, True, False, False)
                Else
                    temBool = ChangeStocks(FormPurchaseItem.ID, !ItemBatchID, !PurchaseQuentity + !PurchaseFreeQuentity, !DepartmentID, False, True, False)
                    temBool = ChangeStocks(FormPurchaseItem.ID, ItemBatchID(FormPurchaseItem.ID, UCase(txtBatch.Text), dtpDOE.Value, dtpDOM.Value), Val(txtQty.Text) + Val(txtFQty.Text), cmbDepartment.BoundText, True, False, False)
                End If
            End If
        Else
            temSql = "Select * from tblPurchaseItem"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PurchaseBillID = Val(txtBillID.Text)
            temBool = ChangeStocks(FormPurchaseItem.ID, ItemBatchID(FormPurchaseItem.ID, UCase(txtBatch.Text), dtpDOE.Value, dtpDOM.Value), Val(txtQty.Text) + Val(txtFQty.Text), cmbDepartment.BoundText, True, False, False)
        End If
        !ItemBatchID = ItemBatchID(FormPurchaseItem.ID, UCase(txtBatch.Text), dtpDOE.Value, dtpDOM.Value)
        !ItemID = FormPurchaseItem.ID
        !DepartmentID = Val(cmbDepartment.BoundText)
        !PurchaseRate = Val(txtPPrice.Text)
        !PurchaseQuentity = Val(txtQty.Text)
        !PurchaseFreeQuentity = Val(txtFQty.Text)
        !PurchaseValue = Val(txtPurchaseValue.Text)
        !SaleRate = Val(txtSPrice.Text)
        !SaleValue = Val(txtSaleValue.Text)
        !LastPurchaseRate = Val(txtLastPPI.Text)
        !LastSaleRate = Val(txtLastSPI.Text)
        .Update
    End With
    
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSql = "Select * from tblItem where ItemID = " & FormPurchaseItem.ID
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !SalePrice = Val(txtSPrice.Text)
            !PurchasePrice = Val(txtPPrice.Text)
            !LastSalePrice = Val(txtLastSPI.Text)
            !LastPurchasePrice = Val(txtLastPPI.Text)
            .Update
        End If
        .Close
    End With
    
    Call ClearAddValues
    Call FormatGrid
    Call FillGrid
    FocusControl.SetFocus
End Sub

Private Sub ClearAddValues()
    txtBatch.Text = Empty
    txtBatchID.Text = Empty
    txtDeleteID.Text = Empty
    txtEditID.Text = Empty
    txtSPrice.Text = Empty
    
    txtQty.Text = Empty
    txtFPQty.Text = Empty
    txtFQty.Text = Empty
    txtLastPPI.Text = Empty
    txtLastPPP.Text = Empty
    txtLastSPI.Text = Empty
    txtLastSPP.Text = Empty
    txtPPPrice.Text = Empty
    txtPPrice.Text = Empty
    txtPQty.Text = Empty
    txtPSPrice.Text = Empty
    txtSaleValue.Text = Empty
    
    For i = 0 To lblIssueUnit.UBound
        lblIssueUnit.Item(i).Caption = Empty
    Next i
    For i = 0 To lblPackUnit.UBound
        lblPackUnit.Item(i).Caption = Empty
    Next i
    

    cmbCode.Text = Empty
    cmbCategory.Text = Empty
    cmbItem.Text = Empty
    
End Sub

Private Sub ClearBillValues()
    lblGrossTotal.Caption = "0.00"
    lblNetTotal.Caption = "0.00"
    lblTotalSaleValue.Caption = "0.00"
    txtDiscount.Text = "0.00"
    txtBillID.Text = Empty
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim temBool As Boolean
    If Val(txtDeleteID.Text) = 0 Then
        MsgBox "noting to Delete"
        Exit Sub
    End If
    Dim rsItem As New ADODB.Recordset
    With rsItem
        If .State = 1 Then .Close
        temSql = "Select * from tblPurchaseItem where PurchaseItemID = " & Val(txtDeleteID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            temBool = ChangeStocks(!ItemID, !ItemBatchID, !PurchaseQuentity + !PurchaseFreeQuentity, !DepartmentID, False, True, False)
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    Call ClearAddValues
    FocusControl.SetFocus
End Sub

Private Sub btnPurchase_Click()
    txtDiscount_Change
    If Val(txtBillID.Text) = 0 Then
        MsgBox "Noting to purchase"
        Exit Sub
    End If
    If GridItem.Rows < 2 Then
        MsgBox "Noting to purchase"
        Exit Sub
    End If
    If Val(cmbSupplier.BoundText) = 0 Then
        MsgBox "Please select a supplier"
        cmbSupplier.SetFocus
        Exit Sub
    End If
    If Val(cmbPayment.BoundText) = 0 Then
        MsgBox "Please select a payment method"
        cmbPayment.SetFocus
        Exit Sub
    End If
    
    Dim rsExpence As New ADODB.Recordset
    With rsExpence
        If .State = 1 Then .Close
        temSql = "Select * from tblExpence"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DepartmentID = UserDepartmentID
        
        !ExpenceMethodID = 1
        !ToInstitutionID = Val(cmbSupplier.BoundText)
        !BillID = Val(txtBillID.Text)
        !PaymentMethodID = Val(cmbPayment.BoundText)
        !PaymentComments = txtPaymentComments.Text
        !Value = lblGrossTotal.Caption
        !AddedUserID = UserID
        !AddedDateTime = Now
        !AddedDate = Date
        !AddedTime = Time
        !ExpenceDate = Date
        !ExpenceTime = Time
        !ExpenceDateTime = Now
        !ExpenceUserID = UserID
        .Update
        .Close
    End With
    
    Dim rsPurchaseBill As New ADODB.Recordset
    With rsPurchaseBill
        If .State = 1 Then .Close
        temSql = "Select * from tblPurchaseBill where PurchaseBillID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Completed = True
            !CompletedUserID = UserID
            !CompletedDate = Date
            !CompletedTime = Time
            !CompletedDateTime = Now
            !PaymentMethodID = Val(cmbPayment.BoundText)
            !PaymentComments = txtPaymentComments.Text
            !SupplierID = Val(cmbSupplier.BoundText)
            !InvoiceDate = dtpDate.Value
            !InvoiceNumber = txtInvoice.Text
            !Total = Val(lblGrossTotal.Caption)
            !NetTotal = Val(lblNetTotal.Caption)
            !Discount = Val(txtDiscount.Text)
            !SaleValue = Val(lblTotalSaleValue.Caption)
            !SupplierID = Val(cmbSupplier.BoundText)
            !DrawerID = Val(cmbDrawer.BoundText)
            !CheckedUserID = Val(cmbChecked.BoundText)
            !ReceivedUserID = Val(cmbReceived.BoundText)
            .Update
        Else
            .Close
            MsgBox "Not Saved"
            Exit Sub
        End If
    End With
    
    If Val(cmbPayment.BoundText) = 2 Then
        Dim rsSup As New ADODB.Recordset
        With rsSup
            If .State = 1 Then .Close
            temSql = "Select * from tblInstitution where InstitutionID = " & Val(cmbSupplier.BoundText)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Credit = !Credit + Val(lblNetTotal.Caption)
                .Update
            End If
            .Close
        End With
    End If
    
    ChangeDrawerBalances Val(cmbDrawer.BoundText), Val(cmbPayment.BoundText), Val(lblNetTotal.Caption), False, True, False
    
    Call ClearBillValues
    Call ClearAddValues
    Call FormatGrid
    FocusControl.SetFocus
End Sub

Private Sub cmbCategory_Change()
    Dim TemFillCombo As New clsFillCombos
    Dim TemFIllCOmboCode As New clsFillCombos
    If IsNumeric(cmbCategory.BoundText) = True Then
        TemFillCombo.FillLongBoolCombo cmbItem, "Item", "Item", "ItemCategoryID", Val(cmbCategory.BoundText), "IsItemName", True
        TemFIllCOmboCode.FillLongBoolCombo cmbCode, "Item", "ItemCode", "ItemCategoryID", Val(cmbCategory.BoundText), "IsItemName", True
    Else
        TemFillCombo.FillBoolCombo cmbItem, "Item", "Item", "IsItemName", True
        TemFIllCOmboCode.FillBoolCombo cmbCode, "Item", "ItemCode", "IsItemName", True
    End If
End Sub

Private Sub cmbCategory_Click(Area As Integer)
    Set FocusControl = cmbCategory
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If FocusControl.Name = cmbItem.Name Then
            cmbItem.SetFocus
        Else
            cmbCode.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbCode_Change()
    If IsNumeric(cmbCode.BoundText) = True Then cmbItem.BoundText = cmbCode.BoundText
End Sub

Private Sub cmbCode_Click(Area As Integer)
    Set FocusControl = cmbCode
End Sub

Private Sub cmbCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If SSTab1.Tab = 0 Then
            txtQty.SetFocus
        Else
            txtPQty.SetFocus
        End If
        On Error Resume Next: SendKeys "{home}+{end}"
    ElseIf KeyCode = vbKeyEscape Then
        cmbCode.Text = Empty
    End If
End Sub

Private Sub cmbItem_Change()
    If IsNumeric(cmbItem.BoundText) = False Then Exit Sub
    cmbCode.BoundText = cmbItem.BoundText
    FormPurchaseItem.ID = Val(cmbItem.BoundText)
    Call WriteLabels
'    DisplayChange = True
    txtLastPPI.Text = Format(LastPurchasePrice(FormPurchaseItem.ID), "0.00")
    txtLastSPI.Text = Format(LastSalePrice(FormPurchaseItem.ID), "0.00")
'    DisplayChange = False
End Sub

Private Sub NewSalePrice()
    SaveCommonSettings Me
    
    
    If SSTab1.Tab = 0 Then
        txtSPrice.Text = Format(Val(txtPPrice.Text) * (FormPurchaseItem.SalesMargin + 100) / 100, "0.00")
    Else
        txtPSPrice.Text = Format(Val(txtPPPrice.Text) * (FormPurchaseItem.SalesMargin + 100) / 100, "0.00")
    End If
End Sub


Private Sub ChangeIssueToPack()
    Dim TemIssueUnitsForAPack As Double
    
    If FormPurchaseItem.IssueUnitsPerPack <> 0 Then
        TemIssueUnitsForAPack = Abs(FormPurchaseItem.IssueUnitsPerPack)
    Else
        TemIssueUnitsForAPack = 1
    End If

    
    If DisplayChange = True Then
        txtLastPPP.Text = Val(txtLastPPI.Text) * TemIssueUnitsForAPack
        txtLastSPP.Text = Val(txtLastSPI.Text) * TemIssueUnitsForAPack
        txtPSPrice.Text = Val(txtSPrice.Text) * TemIssueUnitsForAPack
        txtPPPrice.Text = Val(txtPPrice.Text) * TemIssueUnitsForAPack
        
        txtPQty.Text = Val(txtQty.Text) / TemIssueUnitsForAPack
        txtFPQty.Text = Val(txtFQty.Text) / TemIssueUnitsForAPack
    Else
        If SSTab1.Tab = 0 Then
            txtLastPPP.Text = Val(txtLastPPI.Text) * TemIssueUnitsForAPack
            txtLastSPP.Text = Val(txtLastSPI.Text) * TemIssueUnitsForAPack
            txtPSPrice.Text = Val(txtSPrice.Text) * TemIssueUnitsForAPack
            txtPPPrice.Text = Val(txtPPrice.Text) * TemIssueUnitsForAPack
            
            txtPQty.Text = Val(txtQty.Text) / TemIssueUnitsForAPack
            txtFPQty.Text = Val(txtFQty.Text) / TemIssueUnitsForAPack
            
        Else
            txtLastPPI.Text = Val(txtLastPPP.Text) / TemIssueUnitsForAPack
            txtLastSPI.Text = Val(txtLastSPP.Text) / TemIssueUnitsForAPack
            txtPPrice.Text = Val(txtPPPrice.Text) / TemIssueUnitsForAPack
            txtSPrice.Text = Val(txtPSPrice.Text) / TemIssueUnitsForAPack
            
            txtFQty.Text = Val(txtFPQty.Text) * TemIssueUnitsForAPack
            txtQty.Text = Val(txtPQty.Text) * TemIssueUnitsForAPack
        End If
    End If
End Sub

Private Sub cmbItem_Click(Area As Integer)
    cmbItem_Change
    Set FocusControl = cmbItem
End Sub

Private Sub SetColours()
    Me.ForeColor = DefaultColourScheme.LabelForeColour
    Me.BackColor = DefaultColourScheme.LabelBackColour
    On Error Resume Next
    Dim MyControl As Control
    For Each MyControl In Controls
        If InStr(UCase(MyControl.Name), "BTN") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
        ElseIf InStr(UCase(MyControl.Name), "LST") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXTID") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "CMB") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ComboForeColour
            MyControl.BackColor = DefaultColourScheme.ComboBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXT") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        ElseIf InStr(UCase(MyControl.Name), "DTP") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        Else
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
            MyControl.BackStyle = 0
        End If
    Next
End Sub

Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If FocusControl.Name = cmbCode.Name Then
            cmbCode.SetFocus
        Else
        If SSTab1.Tab = 0 Then
            txtQty.SetFocus
        Else
            txtPQty.SetFocus
        End If
        On Error Resume Next: SendKeys "{home}+{end}"
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbItem.Text = Empty
    End If
End Sub


Private Sub dtpDOM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDOE.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtpDOM.Value = Date
    End If
End Sub

Private Sub dtpDOE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
    ElseIf KeyCode = vbKeyEscape Then
        dtpDOE.Value = Date
    End If
End Sub

Private Sub Form_Load()
    
    
    Call SetColours
    Call FormatGrid
    txtBillID.Text = FormPurchaseBill.NewPurchaseBillID
    Call FillCombos
    Call FillPrinters
    Call GetSettings
    txtBillID.Text = FormPurchaseBill.NewPurchaseBillID
    Call FillGrid
    
End Sub

Private Sub WriteLabels()
    For i = 0 To lblIssueUnit.UBound
        lblIssueUnit.Item(i).Caption = FormPurchaseItem.IssueUnit
    Next i
    For i = 0 To lblPackUnit.UBound
        lblPackUnit.Item(i).Caption = FormPurchaseItem.PackUnit
    Next i
End Sub

Private Sub GridItem_Click()
    With GridItem
        txtDeleteID.Text = Val(.TextMatrix(.row, 0))
    End With
End Sub

Private Sub GridItem_DblClick()
    DisplayChange = True
    With GridItem
        txtEditID.Text = Val(.TextMatrix(.row, 0))
    End With
    Dim rsItemEdit As New ADODB.Recordset
    With rsItemEdit
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.ItemCategoryID, tblItemBatch.ItemBatch, tblItemBatch.DOE, tblItemBatch.DOM, tblPurchaseItem.* FROM (tblPurchaseItem LEFT JOIN tblItem ON tblPurchaseItem.ItemID = tblItem.ItemID) LEFT JOIN tblItemBatch ON tblPurchaseItem.ItemBatchID = tblItemBatch.ItemBatchID WHERE (((tblPurchaseItem.PurchaseItemID)=" & Val(txtEditID.Text) & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbDepartment.BoundText = !DepartmentID
            cmbCategory.BoundText = !ItemCategoryID
            cmbItem.BoundText = !ItemID
            cmbCode.BoundText = !ItemID
            txtLastPPI.Text = Format(!LastPurchaseRate, "0.00")
            txtLastSPI.Text = Format(!LastSaleRate, "0.00")
            txtQty.Text = !PurchaseQuentity
            txtFQty.Text = !PurchaseFreeQuentity
            txtPPrice.Text = Format(!PurchaseRate, "0.00")
            txtSPrice.Text = Format(!SaleRate, "0.00")
            txtPurchaseValue.Text = Format(!PurchaseValue, "0.00")
            txtBatch.Text = !ItemBatch
            dtpDOE.Value = !DOE
            dtpDOM.Value = !DOE
            txtBatchID.Text = !ItemBatchID
        End If
        .Close
    End With
    DisplayChange = False
End Sub

Private Sub GetSettings()
    Dim temCtrlName As String
    Dim MyControl As Control
    
    dtpDOM.Value = Date
    dtpDOE.Value = Date
    
    cmbChecked.BoundText = UserID
    cmbReceived.BoundText = UserID
    
    GetCommonSettings Me
    
    
    temCtrlName = GetSetting(App.EXEName, Me.Name, "FocusControl", cmbCategory.Name)
    
    For Each MyControl In Controls
        If MyControl.Name = temCtrlName Then
            Set FocusControl = MyControl
        End If
    Next
    
    FocusControl.TabIndex = 1
    
    dtpDate.Value = Date
    
    cmbChecked.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbChecked.Name, UserID))
    cmbReceived.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbReceived.Name, UserID))
    cmbSupplier.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbSupplier.Name, 1))
    cmbPayment.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbPayment.Name, 1))
    chkPrint.Value = Val(GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1))
    cmbDrawer.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbDrawer.Name, 1))
    cmbDepartment.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbDepartment.Name, UserDepartmentID))
    
    SSTab1.Tab = Val(GetSetting(App.EXEName, Me.Name, SSTab1.Name, 0))
    SSTab2.Tab = Val(GetSetting(App.EXEName, Me.Name, SSTab2.Name, 0))
    
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPrinter_Click
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbReceived.Name, cmbReceived.BoundText
    SaveSetting App.EXEName, Me.Name, cmbChecked.Name, cmbChecked.BoundText
    SaveSetting App.EXEName, Me.Name, cmbSupplier.Name, cmbSupplier.BoundText
    SaveSetting App.EXEName, Me.Name, cmbPayment.Name, cmbPayment.BoundText
    SaveSetting App.EXEName, Me.Name, cmbDrawer.Name, cmbDrawer.BoundText
    SaveSetting App.EXEName, Me.Name, cmbDepartment.Name, cmbDepartment.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
    SaveSetting App.EXEName, Me.Name, "FocusControl", FocusControl.Name
    SaveSetting App.EXEName, Me.Name, SSTab1.Name, SSTab1.Tab
    SaveSetting App.EXEName, Me.Name, SSTab2.Name, SSTab2.Tab
End Sub

Private Sub cmbPrinter_Change()
    cmbPrinter_Click
End Sub

Private Sub cmbPrinter_Click()
    'cmbPaper.Clear
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = PrescreptionPaperHeight
            .cy = PrescreptionPaperWidth
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
    
End Sub

Private Sub FormatGrid()
    With GridItem
        .Cols = 11
        .Rows = 2
        .RowHeight(0) = .RowHeight(1) * 2
        .Rows = 1
        .row = 0
        


        .col = 0
        For i = 0 To .Cols - 1
            .col = i
            .CellAlignment = 4
            Select Case i
                Case 1:     .Text = "No"
                            .ColWidth(i) = 400
                Case 2:     .Text = "Item"
                            .ColWidth(i) = 3600
                Case 3:     .Text = "Purchase Rate"
                            .ColWidth(i) = 1000
                Case 4:     .Text = "Purchase Qty"
                            .ColWidth(i) = 1000
                Case 5:     .Text = "Free Qty"
                            .ColWidth(i) = 1000
                Case 6:     .Text = "Purchase Value"
                            .ColWidth(i) = 1000
                Case 7:     .Text = "Batch"
                            .ColWidth(i) = 1400
                Case 8:     .Text = "Expiary"
                            .ColWidth(i) = 1600
                Case 9:     .Text = "Slaes Rate"
                            .ColWidth(i) = 1000
                Case 10:     .Text = "Sale Value"
                            .ColWidth(i) = 1000
                Case Else:  .ColWidth(i) = 0
            End Select
        Next i
    
    End With
    '   0   PurchaseItemID
    '   1   No
    '   2   Item
    '   3   Purchase Rate
    '   4   Purchase Qty
    '   5   Free Qty
    '   6   Purchase Value
    '   7   Batch
    '   8   Expiary
    '   9   Sale Rate
    '   10  Sale Value

End Sub

Private Sub FillGrid()
    Dim GrossTotal As Double
    Dim SaleValueTotal As Double
    Dim rsPItem As New ADODB.Recordset
    With rsPItem
        temSql = "SELECT tblItemBatch.ItemBatch, tblItemBatch.DOE, tblPurchaseItem.* FROM tblPurchaseItem LEFT JOIN tblItemBatch ON tblPurchaseItem.ItemBatchID = tblItemBatch.ItemBatchID WHERE tblPurchaseItem.PurchaseBillID = " & Val(txtBillID.Text) & " AND tblPurchaseItem.Deleted = false"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                FormPurchaseItem.ID = !ItemID
                GridItem.Rows = GridItem.Rows + 1
                GridItem.row = GridItem.Rows - 1
                GridItem.col = 0
                GridItem.Text = !PurchaseItemID
                GridItem.col = 1
                GridItem.Text = GridItem.row
                GridItem.col = 2
                If Trim(FormPurchaseItem.Code) <> "" Then
                    GridItem.Text = FormPurchaseItem.Item & " (" & FormPurchaseItem.Code & ")"
                Else
                    GridItem.Text = FormPurchaseItem.Item
                End If
                GridItem.col = 6
                GrossTotal = GrossTotal + !PurchaseValue
                SaleValueTotal = SaleValueTotal + !SaleValue
                GridItem.Text = Format(!PurchaseValue, "0.00")
                GridItem.col = 7
                GridItem.Text = !ItemBatch
                GridItem.col = 8
                GridItem.Text = Format(!DOE, "dd MMM yyyy")
                If SSTab1.Tab = 0 Then
                    GridItem.col = 3
                    GridItem.Text = Format(!PurchaseRate, "0.00")
                    GridItem.col = 4
                    GridItem.Text = !PurchaseQuentity
                    GridItem.col = 5
                    GridItem.Text = !PurchaseFreeQuentity
                    GridItem.col = 9
                    GridItem.Text = Format(!SaleRate, "0.00")
                    GridItem.col = 10
                    GridItem.Text = Format(!SaleValue, "0.00")
                Else
                    GridItem.col = 3
                    GridItem.Text = Format(!PurchaseRate * FormPurchaseItem.IssueUnitsPerPack, "0.00")
                    GridItem.col = 4
                    GridItem.Text = Format(!PurchaseQuentity / FormPurchaseItem.IssueUnitsPerPack, "0.00")
                    GridItem.col = 5
                    GridItem.Text = Format(!PurchaseFreeQuentity / FormPurchaseItem.IssueUnitsPerPack, "0.00")
                    GridItem.col = 9
                    GridItem.Text = Format(!SaleRate * FormPurchaseItem.IssueUnitsPerPack, "0.00")
                    GridItem.col = 10
                    GridItem.Text = Format(!SaleValue, "0.00")
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
    lblGrossTotal.Caption = Format(GrossTotal, "0.00")
    lblTotalSaleValue.Caption = Format(SaleValueTotal, "0.00")
End Sub

Private Sub CalculateTotals()
    lblNetTotal.Caption = Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)
End Sub

Private Sub CalculateItemValues()
    txtPurchaseValue.Text = Format((Val(txtQty.Text)) * Val(txtPPrice.Text), "0.00")
    txtSaleValue.Text = Format((Val(txtQty.Text) + Val(txtFQty.Text)) * Val(txtSPrice.Text), "0.00")
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTab2.Tab = SSTab1.Tab
    Call FormatGrid
    Call FillGrid
    Call CalculateTotals
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
    SSTab1.Tab = SSTab2.Tab
End Sub

Private Sub txtBatch_Change()
    MyBatch = FindBatch(Val(cmbItem.BoundText), txtBatch.Text)
    If MyBatch.Exists = True Then
        dtpDOM.Value = MyBatch.DOM
        dtpDOE.Value = MyBatch.DOE
    End If
End Sub

Private Sub txtBatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If MyBatch.Exists = True Then
            btnAdd_Click
        Else
            dtpDOM.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtBatch.Text = Empty
    End If
End Sub

Private Sub txtDiscount_Change()
    lblNetTotal.Caption = Format(Val(lblGrossTotal.Caption) - Val(txtDiscount.Text), "0.00")
End Sub

Private Sub txtFPQty_Change()
    Call ChangeIssueToPack
End Sub

Private Sub txtFPQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPPPrice.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtFPQty.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtFPQty.Text = Val(txtFPQty.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtFPQty.Text) > 0 Then
            txtFPQty.Text = Val(txtFPQty.Text) - 1
        End If
    End If
End Sub

Private Sub txtFQty_Change()
    Call ChangeIssueToPack
    Call CalculateItemValues
End Sub

Private Sub txtFQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPPrice.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtFQty.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtFQty.Text = Val(txtFQty.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtFQty.Text) > 0 Then
            txtFQty.Text = Val(txtFQty.Text) - 1
        End If
    End If
End Sub

Private Sub txtLastPPI_Change()
    Call ChangeIssueToPack
    Call CalculateItemValues
End Sub

Private Sub txtLastPPP_Change()
    Call ChangeIssueToPack
End Sub

Private Sub txtLastSPI_Change()
    Call ChangeIssueToPack
    Call CalculateItemValues
End Sub

Private Sub txtLastSPP_Change()
    Call ChangeIssueToPack
End Sub

Private Sub txtPPPrice_Change()
    Call ChangeIssueToPack
    Call NewSalePrice
End Sub

Private Sub txtPPPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPSPrice.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtPSPrice.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtPSPrice.Text = Val(txtPSPrice.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtPSPrice.Text) > 0 Then
            txtPSPrice.Text = Val(txtPSPrice.Text) - 1
        End If
    End If

End Sub

Private Sub txtPPrice_Change()
    Call ChangeIssueToPack
    Call CalculateItemValues
    Call NewSalePrice
End Sub

Private Sub txtPPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtSPrice.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtPPrice.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtPPrice.Text = Val(txtPPrice.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtPPrice.Text) > 0 Then
            txtPPrice.Text = Val(txtPPrice.Text) - 1
        End If
    End If
End Sub

Private Sub txtPQty_Change()
    Call ChangeIssueToPack
End Sub

Private Sub txtPQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFPQty.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtPQty.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtPQty.Text = Val(txtPQty.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtPQty.Text) > 0 Then
            txtPQty.Text = Val(txtPQty.Text) - 1
        End If
    End If
End Sub

Private Sub txtPSPrice_Change()
    Call ChangeIssueToPack
End Sub

Private Sub txtPSPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtBatch.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtPSPrice.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtPSPrice.Text = Val(txtPSPrice.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtPSPrice.Text) > 0 Then
            txtPSPrice.Text = Val(txtPSPrice.Text) - 1
        End If
    End If

End Sub

Private Sub txtQty_Change()
    Call ChangeIssueToPack
    Call CalculateItemValues
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFQty.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtQty.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtQty.Text = Val(txtQty.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtQty.Text) > 0 Then
            txtQty.Text = Val(txtQty.Text) - 1
        End If
    End If
End Sub

Private Sub txtSPrice_Change()
    Call ChangeIssueToPack
    Call CalculateItemValues
End Sub

Private Sub txtSPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtBatch.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtSPrice.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtSPrice.Text = Val(txtSPrice.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtSPrice.Text) > 0 Then
            txtSPrice.Text = Val(txtSPrice.Text) - 1
        End If
    End If
End Sub
