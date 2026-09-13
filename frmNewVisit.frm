VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNewVisit 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Visit - "
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13410
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13410
   Begin VB.CheckBox chkDelete 
      Height          =   255
      Left            =   1440
      TabIndex        =   103
      Top             =   8880
      Visible         =   0   'False
      Width           =   255
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   120
      TabIndex        =   102
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   255
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   3000
      TabIndex        =   95
      Top             =   8760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin EmrForGps.UCPopup ucpPrescreption 
      Left            =   2640
      Top             =   8760
      _extentx        =   318
      _extenty        =   370
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   12120
      TabIndex        =   86
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkClose 
      Caption         =   "Close after Printing"
      Height          =   360
      Left            =   10800
      TabIndex        =   54
      Top             =   8160
      Width           =   2415
   End
   Begin VB.Frame framePP 
      Caption         =   "Frame7"
      Height          =   495
      Left            =   0
      TabIndex        =   55
      Top             =   8160
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton optPrescreption 
         Caption         =   "Option1"
         Height          =   240
         Left            =   360
         TabIndex        =   57
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optPharmacy 
         Caption         =   "Option1"
         Height          =   240
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0080FF80&
      Caption         =   "Visit I&nvestigations"
      Height          =   1815
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   4935
      Begin MSDataListLib.DataCombo cmbIx 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridIx 
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1931
         _Version        =   393216
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnIxAdd 
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   65280
         Caption         =   "&Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnIxDelete 
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   255
         Caption         =   "&Delete"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtIxDeleteID 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1560
         TabIndex        =   73
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtIxEditID 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1080
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin btButtonEx.ButtonEx btnIxResults 
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "&Results"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtIxID 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2040
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Patient Details And Visits"
      Height          =   3015
      Left            =   4920
      TabIndex        =   60
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtHistory 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   74
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtPatientDetails 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   63
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtVisit 
         Height          =   2655
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   240
         Width           =   2895
      End
      Begin MSDataListLib.DataList lstVisit 
         Height          =   825
         Left            =   120
         TabIndex        =   61
         Top             =   2040
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1455
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin VB.TextBox txtVisitTime 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   68
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtVisitDate 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         TabIndex        =   69
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   3840
         TabIndex        =   71
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   255
         Left            =   5160
         TabIndex        =   70
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      Caption         =   "Dia&gnosis"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtDiagnosis 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   3735
      End
      Begin MSDataListLib.DataCombo cmbDiagnosis 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridDiagnosis 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1931
         _Version        =   393216
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnDxAdd 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   65280
         BorderColor     =   16384
         Caption         =   "&Add"
         ForeColor       =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnDxDelete 
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   255
         Caption         =   "&Delete"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtDxEditID 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1080
         TabIndex        =   64
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDxDeleteID 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1560
         TabIndex        =   65
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin btButtonEx.ButtonEx btnDxDetails 
         Height          =   375
         Left            =   3960
         TabIndex        =   101
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   255
         Caption         =   "&Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Medicines"
      Height          =   4455
      Left            =   0
      TabIndex        =   12
      Top             =   4200
      Width           =   13335
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   5
         TabHeight       =   520
         BackColor       =   16761087
         TabCaption(0)   =   "Prescreption"
         TabPicture(0)   =   "frmNewVisit.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "gridPrescreption"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Items"
         TabPicture(1)   =   "frmNewVisit.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(1)=   "Label5"
         Tab(1).Control(2)=   "cmbPIssueUnit"
         Tab(1).Control(3)=   "cmbPItem"
         Tab(1).Control(4)=   "gridPharmacy"
         Tab(1).Control(5)=   "txtPIssueQty"
         Tab(1).ControlCount=   6
         Begin VB.TextBox txtPIssueQty 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   -67680
            TabIndex        =   96
            Top             =   600
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid gridPrescreption 
            Height          =   2895
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   5106
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid gridPharmacy 
            Height          =   2175
            Left            =   -74880
            TabIndex        =   32
            Top             =   1080
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   3836
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo cmbPItem 
            Height          =   360
            Left            =   -74880
            TabIndex        =   97
            Top             =   600
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbPIssueUnit 
            Height          =   360
            Left            =   -66840
            TabIndex        =   98
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "&Issue"
            Height          =   255
            Left            =   -67680
            TabIndex        =   100
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "&Item"
            Height          =   255
            Left            =   -74880
            TabIndex        =   99
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtProcedureFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtMedicinesFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtConsultationFee 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   80
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox txtTotalCharge 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtScaleHeight 
         Height          =   360
         Left            =   8520
         TabIndex        =   78
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox txtScaleWidth 
         Height          =   360
         Left            =   7920
         TabIndex        =   77
         Top             =   3600
         Width           =   1575
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   3480
         Width           =   3255
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox txtIssueQty 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   11280
         TabIndex        =   26
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtDuration 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   10680
         TabIndex        =   23
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtDose 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6600
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox chkIndoor 
         BackColor       =   &H00FFC0FF&
         Caption         =   "&Indoor Issue"
         Height          =   375
         Left            =   10320
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbItem 
         Height          =   360
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDoseUnit 
         Height          =   360
         Left            =   7320
         TabIndex        =   19
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbFrequencyUnit 
         Height          =   360
         Left            =   8880
         TabIndex        =   21
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDurationUnit 
         Height          =   360
         Left            =   11280
         TabIndex        =   24
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbIssueUnit 
         Height          =   360
         Left            =   11880
         TabIndex        =   27
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCategory 
         Height          =   360
         Left            =   4920
         TabIndex        =   16
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx btnRxAdd 
         Height          =   375
         Left            =   10320
         TabIndex        =   28
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   65280
         Caption         =   "&Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnRxDelete 
         Height          =   375
         Left            =   10320
         TabIndex        =   29
         Top             =   1920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   255
         Caption         =   "&Delete"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnRefill 
         Height          =   375
         Left            =   11880
         TabIndex        =   33
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "Refi&ll"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnVisitRefill 
         Height          =   375
         Left            =   11880
         TabIndex        =   34
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "&Visit Refill"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtAddID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtDeletedID 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         TabIndex        =   67
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fee"
         Height          =   255
         Left            =   10920
         TabIndex        =   83
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Category"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDuration 
         BackStyle       =   0  'Transparent
         Caption         =   "&Duration"
         Height          =   255
         Left            =   10680
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFrequency 
         BackStyle       =   0  'Transparent
         Caption         =   "&Frequency"
         Height          =   255
         Left            =   8880
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDose 
         BackStyle       =   0  'Transparent
         Caption         =   "&Dose"
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Item"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H008080FF&
      Caption         =   "This &visit details"
      Height          =   1215
      Left            =   4920
      TabIndex        =   53
      Top             =   3000
      Width           =   8415
      Begin VB.TextBox txtBmi 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6600
         TabIndex        =   51
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtHt 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6600
         TabIndex        =   48
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtPr 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4440
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtComments 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtWeight 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   6600
         TabIndex        =   45
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtSBP 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4440
         TabIndex        =   40
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtDBP 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4920
         TabIndex        =   41
         Top             =   720
         Width           =   495
      End
      Begin btButtonEx.ButtonEx btnAddBP 
         Height          =   255
         Left            =   5520
         TabIndex        =   43
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         Appearance      =   3
         BackColor       =   8421631
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   104
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "kg/m"
         Height          =   255
         Left            =   7680
         TabIndex        =   52
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "BMI"
         Height          =   255
         Left            =   6240
         TabIndex        =   50
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         Height          =   255
         Left            =   7680
         TabIndex        =   49
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ht."
         Height          =   255
         Left            =   6240
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "/min"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "P&R"
         Height          =   255
         Left            =   4080
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Wt."
         Height          =   255
         Left            =   6240
         TabIndex        =   44
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         Height          =   255
         Left            =   7680
         TabIndex        =   46
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "&BP"
         Height          =   255
         Left            =   4080
         TabIndex        =   39
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "mmHg"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   3015
      Left            =   11280
      TabIndex        =   56
      Top             =   0
      Width           =   2055
      Begin btButtonEx.ButtonEx btnPatient 
         Height          =   300
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "Pa&tient"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnPastHistory 
         Height          =   300
         Left            =   120
         TabIndex        =   87
         Top             =   525
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "&History"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnClinical 
         Height          =   300
         Left            =   120
         TabIndex        =   88
         Top             =   810
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "&Clinical"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnIx 
         Height          =   300
         Left            =   120
         TabIndex        =   89
         Top             =   1095
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "In&vestigations"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnImages 
         Height          =   300
         Left            =   120
         TabIndex        =   90
         Top             =   1380
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "I&mages"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnReferrals 
         Height          =   300
         Left            =   120
         TabIndex        =   91
         Top             =   1665
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "&Referrals"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnCertificates 
         Height          =   300
         Left            =   120
         TabIndex        =   92
         Top             =   1950
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "Certi&ficates"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnProcedures 
         Height          =   300
         Left            =   120
         TabIndex        =   93
         Top             =   2235
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "Pr&ocedures"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx btnCharges 
         Height          =   300
         Left            =   120
         TabIndex        =   94
         Top             =   2520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Appearance      =   3
         BackColor       =   65535
         Caption         =   "C&harges"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   10800
      TabIndex        =   85
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   255
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmNewVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
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
    Dim rsIxList As New ADODB.Recordset
    Dim i As Integer
    Dim rsPastVisits As New ADODB.Recordset

    Dim AlternateItem() As String
    Dim AlternateItemID() As Long
    Dim AlternateItemCount As Long
    Dim ItemRow As Long
    Dim TopRow As Long
    Dim ItemCatID() As Long
    Dim Dose() As Double
    Dim DoseUnitID() As Long
    Dim FrequencyID() As Long
    Dim NewDose() As String
    Dim NewDoseCount As Long

    Dim BasicItem As New clsItem

    Public FormPatient As New clsPatient
    Public FormVisit As New clsVisit

    Dim PrescreptionIndoorForeColour As Long
    Dim PrescreptionOutDoorForeColour As Long

Public Sub SaveVisitDetails()
    
    If chkDelete.Value = 1 Then Exit Sub
    
    Dim rsVisit As New ADODB.Recordset
    With rsVisit
        If .State = 1 Then .Close
        temSql = "Select * from tblVisit where VisitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !SBP = Val(txtSBP.Text)
            !DBP = Val(txtDBP.Text)
            !VisitWeight = Val(txtWeight.Text)
            !pr = Val(txtPr.Text)
            !height = Val(txtHt.Text)
            !bmi = Val(txtBmi.Text)
            !Completed = True
            !CompletedTime = Now
            !CompletedUserID = UserID
            !Comments = txtComments.Text
            !TotalFee = Val(txtTotalCharge.Text)
            !ConsultationFee = Val(txtConsultationFee.Text)
            !ProcedureFee = Val(txtProcedureFee.Text)
            !MedicineFee = Val(txtMedicinesFee.Text)
            If AutoIssueIndoor = True Then
                !IndoorIssued = True
                !IndoorIssuedTime = Now
                !IndoorIssuedUserID = UserID
            End If
            !HasIndoorDrugs = HasIndoorDrugs(FormVisit.VisitID)
            !HasOutdoorDrugs = HasOutdoorDrugs(FormVisit.VisitID)
            .Update
        End If
        .Close
        temSql = "Select * from tblQueue where QueueID = " & FormVisit.QueueID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If AutoIssueIndoor = True Then
                !IndoorIssued = True
                !IndoorIssuedTime = Now
                !IndoorIssuedUserID = UserID
            End If
            !weight = Val(txtWeight.Text)
            !HasIndoorDrugs = HasIndoorDrugs(FormVisit.VisitID)
            !HasOutdoorDrugs = HasOutdoorDrugs(FormVisit.VisitID)
            .Update
        End If
        .Close
    End With
    FormVisit.Refresh
End Sub


Private Sub btnAddBP_Click()
    If Trim(txtComments.Text) = Empty Then
        txtComments.Text = "BP : " & txtSBP.Text & "/" & txtDBP.Text & " mmHg"
    Else
        txtComments.Text = txtComments.Text & vbNewLine & "BP : " & txtSBP.Text & "/" & txtDBP.Text & " mmHg"
    End If
    If IsNumeric(txtWeight.Text) = True Then
        txtComments.Text = txtComments.Text & vbNewLine & "Weight : " & txtWeight.Text & " kg"
    End If
End Sub

Private Sub btnCertificates_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmCertificates
        frmCertificates.Show
        frmCertificates.ZOrder 0
    Else
        Dim MyCet As New frmCertificates
        MyCet.Show
        MyCet.ZOrder 0
    End If
End Sub

Private Sub btnCharges_Click()
    Call SaveVisitDetails
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmCharges
        frmCharges.Show
        frmCharges.ZOrder 0
    Else
        Dim MyChg As New frmCharges
        MyChg.Show
        MyChg.ZOrder 0
    End If
End Sub

Private Sub btnClinical_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmVisitClinicalFinding
        frmVisitClinicalFinding.Show
        frmVisitClinicalFinding.ZOrder 0
    Else
        Dim MyCF As New frmVisitClinicalFinding
        MyCF.Show
        MyCF.ZOrder 0
    End If
End Sub

Private Sub btnClose_Click()
    Call SaveVisitDetails
    Unload Me
End Sub

Private Sub btnDxDetails_Click()
    Dim temDxID As Long
    temDxID = Val(gridDiagnosis.TextMatrix(gridDiagnosis.row, 3))
    If temDxID <> 0 Then
        frmDiseaseInformation.Show
        frmDiseaseInformation.ZOrder 0
        frmDiseaseInformation.cmbDiagnosis.BoundText = temDxID
    End If
End Sub

Private Sub btnImages_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmVisitImages
        frmVisitImages.Show
        frmVisitImages.ZOrder 0
    Else
        Dim MyImg As New frmVisitImages
        MyImg.Show
        MyImg.ZOrder 0
    End If
End Sub

Private Sub btnIx_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmPatientIx
        frmPatientIx.Show
        frmPatientIx.ZOrder 0
    Else
        Dim MyIx As New frmPatientIx
        frmPatientIx.Show
        frmPatientIx.ZOrder 0
    End If
End Sub

Private Sub btnIxResults_Click()
    If IsNumeric(txtIxDeleteID.Text) = False Then
        MsgBox "Please select an investigation"
        gridIx.SetFocus
        Exit Sub
    End If
    Dim temGridRow As Long
    Dim temPatientIXID As Long
    Dim rsTemPatientIx As New ADODB.Recordset
    With rsTemPatientIx
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientIx where PatientIxID = " & Val(txtIxDeleteID.Text) & " AND IxID = " & Val(txtIxID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            temPatientIXID = !PatientIxID
        Else
            .AddNew
            !PatientID = FormPatient.ID
            !VisitID = FormVisit.VisitID
            !IxID = Val(txtIxID.Text)
            !DoctorID = UserID
            !PatientIxID = Val(txtIxDeleteID.Text)
            .Update
            temPatientIXID = !PatientIxID
        End If
        .Close
    End With
    Unload frmIxReport
    frmIxReport.myPATIENTiXid = temPatientIXID
    frmIxReport.InvestigationDetails
    frmIxReport.Show
    frmIxReport.ZOrder 0
End Sub

Private Sub btnPastHistory_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmPastHistory
        frmPastHistory.Show
        frmPastHistory.ZOrder 0
    Else
        Dim MyPastHx As New frmPastHistory
        MyPastHx.Show
        MyPastHx.ZOrder 0
    End If
End Sub

Private Sub btnPatient_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    Unload frmPatientDetails
    frmPatientDetails.Show
    frmPatientDetails.ZOrder 0
    frmPatientDetails.cmbPatient.BoundText = FormVisit.PatientID
End Sub

Private Sub btnPrint_Click()
    Call SaveVisitDetails
    Call PrintVisit
    If chkClose.Value = 1 Then Unload Me
End Sub

Private Sub PrintVisit()
    If printPOS = False Then
        Call PrintVisitNormal
    Else
        Call PrintVisitPOS
    End If
    
End Sub

Private Sub PrintVisitPOS()
    
    On Error GoTo eh
    
    Dim rstemReportItems As New ADODB.Recordset
    Dim temText As String
    Dim MyPrinter As Printer
    
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    
    Dim i As Integer
    Dim tab1 As Integer
    Dim tab2 As Integer
    Dim Tab3 As Integer
    Dim Tab4 As Integer
    Dim Tab5 As Integer
    Dim tab6 As Integer
    Dim tab7 As Integer
    Dim tab8 As Integer
    Dim tab9 As Integer
    
    Printer.Font = "Tw Cen MT Condensed Extra Bold"
    Printer.FontSize = 14
    Printer.ForeColor = vbRed
    Printer.Print InstitutionName

    Printer.FontName = "Tw Cen MT Condensed"

'    Printer.FontSize = 11
'    Printer.Print HospitalDescreption
    
    Printer.ForeColor = vbBlack
    Printer.FontSize = 12
    Printer.Print InstitutionAddress
    Printer.Print InstitutionTelephone
    
    Dim displacement As Integer
    displacement = -7
    tab1 = 8 + displacement
    tab2 = 13 + displacement
    Tab3 = 52 + displacement
    Tab4 = 46 + displacement
    Tab5 = 46 + displacement
    tab6 = 7 + displacement
    tab7 = 9 + displacement
    tab8 = 6 + displacement
    tab9 = 43 + displacement

    Printer.FontSize = 11
    Printer.Print
    
    Printer.Print Tab(tab1); "Patient Name : " & FormPatient.NameWithTitle
    Printer.Print Tab(tab1); "Patient Age  : " & FormPatient.AgeInWords
    Printer.Print Tab(tab1); "Visit Date   : " & Format(FormVisit.VisitDate, LongDateFormat)
    Printer.Print Tab(tab1); "Visit Time   : " & Format(FormVisit.VisitTime, "hh:mm AMPM")
    Printer.Print
    Printer.Print Tab(tab1); "Consultation, Medicines and Other Fees"
    Printer.Print Tab(tab1); Format(FormVisit.TotalFee, "0.00")
    Printer.Print
    Printer.EndDoc
    Exit Sub
    
eh:

    MsgBox "Printer Error."
    Exit Sub

End Sub


Private Sub PrintVisitNormal()
    
    On Error GoTo eh
    
    Dim rstemReportItems As New ADODB.Recordset
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long
    Dim temText As String
    Dim MyPrinter As Printer
    
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    If SelectForm(PrescreptionPaperName, Me.hwnd) <> 1 Then
        MsgBox "Printer Error"
        Exit Sub
    End If
    
    
    Printer.Print
    
'    Printer.ScaleWidth = Val(txtScaleWidth.Text) * 1440
'    Printer.ScaleHeight = Val(txtScaleHeight.Text) * 1440

    
    Printer.FillStyle = vbFSTransparent
    
    With rstemReportItems
        If .State = 1 Then .Close
        temSql = "Select * from tblPrintItem  where Deleted = false AND PrintItemCategoryID = 1"
         
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            TemX1 = Printer.ScaleWidth * !X1
            TemX2 = Printer.ScaleWidth * !X2
            TemY1 = Printer.ScaleHeight * !Y1
            TemY2 = Printer.ScaleHeight * !Y2
            If IsNull(!CircleRadius) = False Then
                temRadius = Printer.ScaleWidth * !CircleRadius
            End If
            If !IsLine = True Then
                Printer.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour
            ElseIf !IsRectangle = True Then
                Printer.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour, B
            ElseIf !IsCircle = True Then
                Printer.Circle (TemX1, TemY1), temRadius, !ForeColour
            ElseIf !IsLabel = True Then
                temText = Format(!ItemText, "")
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            ElseIf !IsText = True Then
                Select Case !ItemText
                    Case "Patient Name": temText = FormPatient.NameWithTitle
                    Case "Patient ID": temText = FormPatient.ID
                    Case "Patient Age in Words": temText = FormPatient.AgeInWords
                    Case "Patient Data of Birth": temText = FormPatient.DateOfBirth
                    Case "Patient Sex": temText = FormPatient.Sex
                    Case "Patient Civil Status": temText = ""
                    Case "Patient Address": temText = FormPatient.Address
                    Case "Patient Telephone": temText = FormPatient.Telephone
                    Case "List of Indoor Medicines": temText = IndoorText
                    Case "List of Outdoor Medicines": temText = OutDoorText
                    Case "List of All Medicines": temText = OutDoorText & vbNewLine & IndoorText
                    Case "List of Investigations": temText = InvestidationText
                    Case "Visit Comments": temText = txtComments.Text
                    Case "Visit Comments, Investigations & Outdoor drugs": temText = txtComments.Text & vbNewLine & vbNewLine & InvestidationText & vbNewLine & vbNewLine & OutDoorText
                    Case "Visit Comments, Outdoor drugs & Investigations": temText = txtComments.Text & vbNewLine & vbNewLine & OutDoorText & vbNewLine & vbNewLine & InvestidationText
                    Case "Total Fee": temText = Format(FormVisit.TotalFee, "0.00")
                    Case "Date (dd mm yy)": temText = Format(FormVisit.VisitDate, "dd mm yy")
                    Case "Date (dd mmm yy)": temText = Format(FormVisit.VisitDate, "dd mmm yy")
                    Case "Date (dd mmm yyyy)": temText = Format(FormVisit.VisitDate, "dd mmm yyyy")
                    Case "Date (dd mmmm yyyy)": temText = Format(FormVisit.VisitDate, "dd mmmm yyyy")
                    Case "Date (dd mm yyyy)": temText = Format(FormVisit.VisitDate, "dd mm yyy")
                    Case "Date (yy mm dd)": temText = Format(FormVisit.VisitDate, "yy mm dd")
                    Case "Date (yy mmm dd)": temText = Format(FormVisit.VisitDate, "yy mmm dd")
                    Case "Date (yyyy mmm dd)": temText = Format(FormVisit.VisitDate, "yyyy mmm dd")
                    Case "Date (yyyy mmmm dd)": temText = Format(FormVisit.VisitDate, "yyyy mmmm dd")
                    Case "Date (yyyy mm dd)": temText = Format(FormVisit.VisitDate, "yyyy mm dd")
                    Case "Short Time": temText = Format(FormVisit.VisitTime, "h:m AMPM")
                    Case "Long Time": temText = Format(FormVisit.VisitTime, "HH:MM AMPM")
                    Case "Queue No.": temText = FormVisit.QueueNo
                    Case "Secession": temText = FormVisit.Secession
                    Case "Visit BP": temText = FormVisit.VisitBP
                    Case "Visit ID": temText = FormVisit.VisitID
                    Case Else
                        temText = ""
                End Select
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            End If
            .MoveNext
        Wend
        .Close
    End With

'    cmbText.AddItem "List of Investigations"
'    cmbText.AddItem "List of Indoor Medicines"
'    cmbText.AddItem "List of Outdoor Medicines"
'    cmbText.AddItem "List of All Medicines"
'    cmbText.AddItem "List of Procedurs"
'    cmbText.AddItem "List of Procedurs"
'    cmbText.AddItem "List of All Medicines"
'    cmbText.AddItem "List of Investigations"
'    cmbText.AddItem "Visit Comments"
'    cmbText.AddItem "Visit Comments, Investigations & Outdoor drugs"
'    cmbText.AddItem "Doctors Name"
'    cmbText.AddItem "Doctors Reg No"
'    cmbText.AddItem "Doctors Designation Line 1"
'    cmbText.AddItem "Doctors Designation Line 2"
'    cmbText.AddItem "Doctors Designation Line 3"
'    cmbText.AddItem "Doctors Designation Line 4"
'    cmbText.AddItem "Institution Name"
'    cmbText.AddItem "Institution Address Line 1"
'    cmbText.AddItem "Institution Address Line 2"
'    cmbText.AddItem "Institution Address Line 3"
'    cmbText.AddItem "Institution Address Line 4"
'    cmbText.AddItem "Date (dd mm yy)"
'    cmbText.AddItem "Date (dd mmm yy)"
'    cmbText.AddItem "Date (dd mmm yyyy)"
'    cmbText.AddItem "Date (dd mmmm yyyy)"
'    cmbText.AddItem "Date (dd mm yyy)"
'    cmbText.AddItem "Date (dd mm yy)"
'    cmbText.AddItem "Date (yy mm dd)"
'    cmbText.AddItem "Date (yy mmm dd)"
'    cmbText.AddItem "Date (yyyy mmm dd)"
'    cmbText.AddItem "Date (yyyy mmmm dd)"
'    cmbText.AddItem "Date (yyy mm dd)"
'    cmbText.AddItem "Date (yy mm dd)"
'    cmbText.AddItem "Short Time"
'    cmbText.AddItem "Long Time"
    '  date
'  outdoor


    Printer.EndDoc
    Exit Sub
    
eh:

    MsgBox "Printer Error."
    Exit Sub

End Sub

Private Function IndoorText() As String
    IndoorText = Empty
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT [tblItem.Item] & ' ' & [tblCategory.category] & ' ' & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' (' & tblVisitItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit & ')'  AS Prescreption " & _
                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
                    "WHERE (((tblVisitItem.VisitID)=" & FormVisit.VisitID & ") AND ((tblVisitItem.Deleted)=False) AND ((tblVisitItem.Indoor)=True) )"
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
'            IndoorText = IndoorText & !Prescreption & vbNewLine
            IndoorText = IndoorText & !Prescreption & Chr(13) & Chr(10)
            .MoveNext
        Wend
        .Close
    End With

End Function


Private Function OutDoorText() As String
    OutDoorText = Empty
    Dim HasOutDoorRx As Boolean
    HasOutDoorRx = False
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT [tblItem.Item] & ' ' & [tblCategory.category] & ',  ' & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblVisitItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] AS Prescreption " & _
                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
                    "WHERE (((tblVisitItem.VisitID)=" & FormVisit.VisitID & ") AND ((tblVisitItem.Deleted)=False) AND ((tblVisitItem.Indoor)=False) )"
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            HasOutDoorRx = True
            OutDoorText = OutDoorText & !Prescreption & vbNewLine
            .MoveNext
        Wend
        .Close
    End With
    If HasOutDoorRx = True Then
        OutDoorText = "Rx" & vbNewLine & OutDoorText
    Else
        OutDoorText = Empty
    End If
End Function


Private Function InvestidationText() As String
    InvestidationText = Empty
    Dim HasIx As Boolean
    HasIx = False
    With rsTem
        temSql = "SELECT tblIx.Ix FROM tblPatientIX LEFT JOIN tblIx ON tblPatientIX.IxID = tblIx.IxID WHERE (((tblPatientIX.VisitID)=" & FormVisit.VisitID & ") AND ((tblPatientIX.Deleted)=False))"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            HasIx = True
            InvestidationText = InvestidationText & vbNewLine & !Ix
            .MoveNext
        Wend
        .Close
    End With
    If HasIx = True Then
        InvestidationText = "Ix" & InvestidationText
    Else
        InvestidationText = Empty
    End If
End Function

Private Sub btnProcedures_Click()
    Call SaveVisitDetails
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmVisitProcedures
        frmVisitProcedures.Show
        frmVisitProcedures.ZOrder 0
    Else
        Dim MyProc As New frmVisitProcedures
        MyProc.Show
        MyProc.ZOrder 0
    End If
End Sub

Private Sub btnReferrals_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    If SingleVisit = True Then
        Unload frmReferrals
        frmReferrals.Show
        frmReferrals.ZOrder 0
    Else
        Dim MyRef As New frmReferrals
        MyRef.Show
        MyRef.ZOrder 0
    End If
End Sub

Private Sub btnRefill_Click()
    Dim rsTem As New ADODB.Recordset
    Dim rsVisitItem As New ADODB.Recordset
    Dim ThisItem As New clsItem
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientItem.* " & _
                    "FROM tblPatientItem " & _
                    "WHERE (((tblPatientItem.PatientID)=" & FormPatient.ID & ") AND ((tblPatientItem.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            ThisItem.ID = !ItemID
            If IsAllergic(FormPatient.ID, ThisItem.ID) = True Then
                i = MsgBox(ThisItem.Item & " may be allergic to this patient. It will NOT be added to the list")
            Else
                If rsVisitItem.State = 1 Then rsVisitItem.Close
                temSql = "Select * from tblVisitItem where VisitItemID = 0"
                rsVisitItem.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                rsVisitItem.AddNew
                rsVisitItem!PatientID = FormPatient.ID
                rsVisitItem!VisitID = FormVisit.VisitID
                rsVisitItem!UserID = UserID
                rsVisitItem!ItemID = !ItemID
                rsVisitItem!ItemCategoryID = !ItemCategoryID
                rsVisitItem!Dose = !Dose
                rsVisitItem!DoseUnitID = !DoseUnitID
                rsVisitItem!FrequencyUnitID = !FrequencyUnitID
                rsVisitItem!Duration = !Duration
                rsVisitItem!DurationUnitID = !DurationUnitID
                rsVisitItem!IssueQuentity = !IssueQuentity
                rsVisitItem!IssueUnitID = !IssueUnitID
                rsVisitItem!Indoor = !Indoor
                rsVisitItem!IndoorOutdoor = !IndoorOutdoor
                
                
                Dim PItem As New clsItem
                Dim temDose As Double
                Dim temDoseUnitID As Long
                
                
                ThisItem.ID = !ItemID
                If ThisItem.IsItem = True Then
                    PItem.ID = ThisItem.ID
                ElseIf ThisItem.IsTradeName = True Then
                    If chkIndoor.Value = 1 Then
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, IndoorDepartmentID, !DoseUnitID, !Dose)
                    Else
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, OutdoorDepartmentID, !DoseUnitID, !Dose)
                    End If
                ElseIf ThisItem.IsGenericName = True Then
                    If chkIndoor.Value = 1 Then
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, IndoorDepartmentID, !DoseUnitID, !Dose)
                    Else
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, OutdoorDepartmentID, !DoseUnitID, !Dose)
                    End If
                Else
                    PItem.ID = 0
                End If
                If PItem.ID <> 0 Then
                    If !IssueUnitID <> 0 Then
                        temDoseUnitID = !IssueUnitID
                    ElseIf !DoseUnitID <> 0 Then
                        temDoseUnitID = !DoseUnitID
                    Else
                        temDoseUnitID = 0
                    End If
                    If PItem.IssueUnitID = !DoseUnitID Then
                        temDose = Val(!Dose)
                    ElseIf PItem.StrengthUnitID = !DoseUnitID Then
                        temDose = Val(!Dose) / PItem.StrengthUnitsPerIssueUnit
                    Else
                        temDose = 1
                    End If
                    If !IssueQuentity <> 0 Then
                        If PItem.IssueUnitID = !IssueUnitID Then
                            rsVisitItem!PIssueQuentity = !IssueQuentity
                        ElseIf PItem.StrengthUnitID = !IssueUnitID Then
                            rsVisitItem!PIssueQuentity = !IssueQuentity / PItem.StrengthUnitsPerIssueUnit
                        Else
                            rsVisitItem!PIssueQuentity = 1
                        End If
                    Else
                        rsVisitItem!PIssueQuentity = CalIssueQty(PItem.ID, temDose, !FrequencyUnitID, !Duration, !DurationUnitID, PItem.MinIQty)
                    End If
                    rsVisitItem!PItemID = PItem.ID
                End If
                    
                
                
                
                
                
                rsVisitItem.Update
            End If
            .MoveNext
        Wend
    End With
    Call FillPrescreption
    Call FillPharmacy
    cmbItem.SetFocus
    On Error Resume Next: SendKeys "{Escape}"
End Sub

Private Sub btnRxAdd_Click()
    Dim ThisItem As New clsItem
    Dim PItem As New clsItem
    Dim temDose As Double
    Dim temDoseUnitID As Long
    
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Please select an Item"
        cmbItem.SetFocus
        On Error Resume Next: SendKeys "{Escape}"
        Exit Sub
    End If
    If IsNumeric(cmbFrequencyUnit.BoundText) = False Then
        MsgBox "Please select a frequency"
        cmbFrequencyUnit.SetFocus
        Exit Sub
    End If
    
    
   
    If IsAllergic(FormPatient.ID, cmbItem.BoundText) = True Then
        i = MsgBox(cmbItem.Text & " may be allergic to this patient. You can not Prescribe this medicine?")
        i = vbNo
        Exit Sub
'        If i <> vbYes Then
'            Unload frmPastHistory
'            frmPastHistory.Show
'            frmPastHistory.ZOrder 0
'            frmPastHistory.SSTab1.Tab = 2
'            Exit Sub
'        End If
    End If
    
    ThisItem.ID = Val(cmbItem.BoundText)
    With rsTem
        If .State = 1 Then .Close
        If Val(txtAddID.Text) <> 0 Then
            temSql = "Select * from tblVisitItem where VisitItemID = " & Val(txtAddID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        Else
            temSql = "Select * from tblVisitItem where VisitItemID = 0 "
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = FormPatient.ID
            !VisitID = FormVisit.VisitID
        End If
        !UserID = UserID
        !ItemID = Val(cmbItem.BoundText)
        !ItemCategoryID = Val(cmbCategory.BoundText)
        
        If IsNumeric(txtDose.Text) = True Then
            !Dose = txtDose.Text
        Else
            !Dose = Null
        End If
        
        !DoseUnitID = Val(cmbDoseUnit.BoundText)
        !FrequencyUnitID = Val(cmbFrequencyUnit.BoundText)
        !Duration = Val(txtDuration.Text)
        !DurationUnitID = Val(cmbDurationUnit.BoundText)
        !IssueQuentity = Val(txtIssueQty.Text)
        !IssueUnitID = Val(cmbIssueUnit.BoundText)
        If chkIndoor.Value = 1 Then
            !Indoor = True
            !IndoorOutdoor = "Indoor"
        Else
            !Indoor = False
            !IndoorOutdoor = "Outdoor"
        End If
        
        
        If ThisItem.IsItem = True Then
            PItem.ID = ThisItem.ID
        ElseIf ThisItem.IsTradeName = True Then
            If chkIndoor.Value = 1 Then
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, Val(cmbCategory.BoundText), IndoorDepartmentID, Val(Format(!DoseUnitID, "0.00")), Val(Format(!Dose, "0.00")))
            Else
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, Val(cmbCategory.BoundText), OutdoorDepartmentID, Val(Format(!DoseUnitID, "0.00")), Val(Format(!Dose, "0.00")))
            End If
        ElseIf ThisItem.IsGenericName = True Then
            If chkIndoor.Value = 1 Then
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, Val(cmbCategory.BoundText), IndoorDepartmentID, !DoseUnitID, Val(Format(!Dose, "0.00")))
            Else
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, Val(cmbCategory.BoundText), OutdoorDepartmentID, !DoseUnitID, Val(Format(!Dose, "0")))
            End If
        Else
            PItem.ID = 0
        End If
        
        If PItem.ID <> 0 Then
            If IsNumeric(cmbIssueUnit.BoundText) = True Then
                temDoseUnitID = cmbIssueUnit.BoundText
            ElseIf IsNumeric(cmbDoseUnit.BoundText) = True Then
                temDoseUnitID = cmbDoseUnit.BoundText
            Else
                temDoseUnitID = 0
            End If
            If PItem.IssueUnitID = temDoseUnitID Then
                temDose = Val(txtDose.Text)
            ElseIf PItem.StrengthUnitID = temDoseUnitID Then
                temDose = Val(txtDose.Text) / PItem.StrengthUnitsPerIssueUnit
            Else
                temDose = 1
            End If
            If Val(txtIssueQty.Text) <> 0 Then
                If PItem.IssueUnitID = Val(cmbIssueUnit.BoundText) Then
                    !PIssueQuentity = Val(txtIssueQty.Text)
                ElseIf PItem.StrengthUnitID = Val(cmbIssueUnit.BoundText) Then
                    !PIssueQuentity = Val(txtIssueQty.Text) / PItem.StrengthUnitsPerIssueUnit
                Else
                    !PIssueQuentity = 1
                End If
            Else
                !PIssueQuentity = CalIssueQty(PItem.ID, temDose, Val(cmbFrequencyUnit.BoundText), Val(txtDuration.Text), Val(cmbDurationUnit.BoundText), PItem.MinIQty)
            End If
            !PItemID = PItem.ID
        End If
        If SSTab1.Tab = 1 Then
            If IsNumeric(txtPIssueQty.Text) = True Then
                !PIssueQuentity = Val(txtPIssueQty.Text)
            End If
            If IsNumeric(cmbPItem.BoundText) = True Then
                !PItemID = Val(cmbPItem.BoundText)
            End If
        End If
        .Update
    End With
    
    Call FillPrescreption
    Call FillPharmacy
    Call ClearItemValues
    
    cmbItem.SetFocus
    On Error Resume Next
    On Error Resume Next: SendKeys "{Escape}"
End Sub



Private Function PharmacyItem(GenericNameID As Long, TradeNameID As Long, ICategoryID As Long, PharmacyDepartmentID As Long, ItemStrengthUnitID As Long, ItemStrength As Double) As Long
    Dim arrPItemID() As Long
    Dim arrPSItemID() As Long
    
    Dim PItemCount As Long
    Dim PSItemCount As Long
    Dim TemSItem As New clsItem
    Dim SameStrengthItemCount As Long
    
    Dim strPItem As String
    Dim i As Integer
    
    Dim rsTemItem As New ADODB.Recordset
    If ICategoryID = 0 Then
        With rsTemItem
           temSql = "SELECT tblVisitItem.ItemCategoryID " & _
                        "FROM (tblVisitItem LEFT JOIN tblPatient ON tblVisitItem.PatientID = tblPatient.PatientID) RIGHT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID " & _
                        "WHERE (((tblItem.GenericNameID)=" & GenericNameID & ") AND ((tblPatient.DateOfBirth) Between #" & Format(DateSerial(Year(FormPatient.DateOfBirth) - (FormPatient.AgeInYears \ 2), 1, 1), "dd MMMM yyyy") & "# And #" & Format(DateSerial(Year(FormPatient.DateOfBirth) + (FormPatient.AgeInYears \ 2), 12, 31), "dd MMMM yyyy") & "#)) " & _
                        "GROUP BY tblVisitItem.ItemCategoryID " & _
                        "Having (((tblVisitItem.ItemCategoryID) <> 0)) " & _
                        "ORDER BY Count(tblVisitItem.ItemCategoryID) "
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If IsNull(!ItemCategoryID) = False Then
                    ICategoryID = !ItemCategoryID
                End If
            End If
        End With
    End If
    If ICategoryID = 0 Then
        With rsTemItem
            temSql = "SELECT First(tblCategory.CategoryID) AS FirstOfCategoryID FROM tblCategory WHERE (((tblCategory.IsItemCategory)=True) AND ((tblCategory.Deleted)=False))"
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                ICategoryID = !FirstOfCategoryID
            End If
        End With
    End If

    If TradeNameID <> 0 Then
        With rsTemItem
            If .State = 1 Then .Close
            temSql = "Select * from tblItem where TradeNameID = " & TradeNameID & " And ItemCategoryID = " & CStr(ICategoryID) & " And IsItemName = true And Deleted = False order by StrengthUnitsPerIssueUnit Desc"
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                ReDim arrPItemID(.RecordCount)
                .MoveFirst
                PItemCount = 0
                While .EOF = False
                    'PharmacyItem = !ItemID
                    arrPItemID(PItemCount) = !ItemID
                    PItemCount = PItemCount + 1
                    .MoveNext
                Wend
            Else
                If .State = 1 Then .Close
                temSql = "Select * from tblItem where TradeNameID = " & TradeNameID & " And IsItemName = true And Deleted = False order by StrengthUnitsPerIssueUnit Desc"
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    .MoveLast
                    ReDim arrPItemID(.RecordCount)
                    .MoveFirst
                    PItemCount = 0
                    While .EOF = False
                        'PharmacyItem = !ItemID
                        arrPItemID(PItemCount) = !ItemID
                        PItemCount = PItemCount + 1
                        .MoveNext
                    Wend
                Else
                    If .State = 1 Then .Close
                    temSql = "Select * from tblItem where GenericNameID = " & GenericNameID & " And ItemCategoryID = " & CStr(ICategoryID) & " And IsItemName = true And Deleted = False order by StrengthUnitsPerIssueUnit Desc"
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    If .RecordCount > 0 Then
                        .MoveLast
                        ReDim arrPItemID(.RecordCount)
                        .MoveFirst
                        PItemCount = 0
                        While .EOF = False
                            'PharmacyItem = !ItemID
                            arrPItemID(PItemCount) = !ItemID
                            PItemCount = PItemCount + 1
                            .MoveNext
                        Wend
                    Else
                        PharmacyItem = 0
                    End If
                    End If
            End If
        End With
    Else
        If GenericNameID <> 0 Then
            With rsTemItem
                If .State = 1 Then .Close
                temSql = "Select * from tblItem where GenericNameID = " & GenericNameID & " And ItemCategoryID = " & CStr(ICategoryID) & " And IsItemName = true And Deleted = False order by StrengthUnitsPerIssueUnit Desc"
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    .MoveLast
                    ReDim arrPItemID(.RecordCount)
                    .MoveFirst
                    PItemCount = 0
                    While .EOF = False
                        'PharmacyItem = !ItemID
                        arrPItemID(PItemCount) = !ItemID
                        PItemCount = PItemCount + 1
                        .MoveNext
                    Wend
                Else
                    PharmacyItem = 0
                End If
            End With
        End If
    End If

    If PItemCount > 0 Then
        For i = 0 To UBound(arrPItemID) - 1
            TemSItem.ID = arrPItemID(i)
            If TemSItem.StrengthUnitID = ItemStrengthUnitID And TemSItem.StrengthUnitsPerIssueUnit = ItemStrength Then
                SameStrengthItemCount = SameStrengthItemCount + 1
            End If
        Next
    End If

    If SameStrengthItemCount >= 1 Then
        ReDim arrPSItemID(SameStrengthItemCount)
        For i = 0 To UBound(arrPItemID) - 1
            TemSItem.ID = arrPItemID(i)
            If TemSItem.StrengthUnitID = ItemStrengthUnitID And TemSItem.StrengthUnitsPerIssueUnit = ItemStrength Then
                arrPSItemID(PSItemCount) = arrPItemID(i)
                PSItemCount = PSItemCount + 1
            End If
        Next
        For i = 0 To UBound(arrPSItemID) - 1
            strPItem = strPItem & " tblItem.ItemID = " & arrPSItemID(i) & " "
            If i <> UBound(arrPSItemID) - 1 Then
                strPItem = strPItem & " OR "
            End If
        Next
        With rsTemItem
            If .State = 1 Then .Close
            temSql = "SELECT tblItem.ItemID " & _
                        "FROM (tblItem LEFT JOIN tblItemBatchStock ON tblItem.ItemID = tblItemBatchStock.ItemID) LEFT JOIN tblItemBatch ON tblItemBatchStock.ItemBatchID = tblItemBatch.ItemBatchID " & _
                        "Where tblItemBatchStock.DepartmentID = " & PharmacyDepartmentID & " And tblItemBatchStock.Stock > 0 And ( " & strPItem & "  ) " & _
                        "ORDER BY tblItemBatch.DOE, tblItemBatchStock.Stock DESC"
        
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                    PharmacyItem = !ItemID
            Else
                PharmacyItem = arrPSItemID(0)
            End If
        End With
        
    Else
        If PItemCount > 0 Then
            For i = 0 To UBound(arrPItemID) - 1
                strPItem = strPItem & " tblItem.ItemID = " & arrPItemID(i) & " "
                If i <> UBound(arrPItemID) - 1 Then
                    strPItem = strPItem & " OR "
                End If
            Next
            With rsTemItem
                If .State = 1 Then .Close
                temSql = "SELECT tblItem.ItemID " & _
                            "FROM (tblItem LEFT JOIN tblItemBatchStock ON tblItem.ItemID = tblItemBatchStock.ItemID) LEFT JOIN tblItemBatch ON tblItemBatchStock.ItemBatchID = tblItemBatch.ItemBatchID " & _
                            "Where tblItemBatchStock.DepartmentID = " & PharmacyDepartmentID & " And tblItemBatchStock.Stock > 0 And ( " & strPItem & "  ) " & _
                            "ORDER BY tblItemBatch.DOE, tblItemBatchStock.Stock DESC"
            
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                        PharmacyItem = !ItemID
                Else
                    PharmacyItem = arrPItemID(0)
                End If
            End With
        End If
    End If
    
End Function



Private Sub ClearItemValues()
    cmbItem.Text = Empty
    cmbCategory.Text = Empty
    txtDose.Text = Empty
    cmbDoseUnit.Text = Empty
    cmbFrequencyUnit.Text = Empty
    txtDuration.Text = Empty
    cmbDurationUnit.Text = Empty
    txtIssueQty.Text = Empty
    cmbIssueUnit.BoundText = Empty
    txtAddID.Text = Empty
    txtDeletedID.Text = Empty
    optPrescreption.Value = True
    cmbPIssueUnit.Text = Empty
    cmbPItem.Text = Empty
    txtPIssueQty.Text = Empty
End Sub

Private Sub btnRxDelete_Click()
    Dim rsTem As New ADODB.Recordset
    If IsNumeric(txtDeletedID.Text) = False Then
        MsgBox "Please select an Item to delete"
        Exit Sub
    End If
    With rsTem
        If Val(txtDeletedID.Text) <> 0 Then
            temSql = "Select * from tblVisitItem where VisitItemID = " & Val(txtDeletedID.Text)
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedTime = Now
                !DeletedUserID = UserID
                .Update
            End If
        End If
        .Close
    End With
    Call ClearItemValues
    Call FillPrescreption
    Call FillPharmacy
    cmbItem.SetFocus
    On Error Resume Next: SendKeys "{Escape}"
End Sub

Private Sub btnVisitRefill_Click()
    Dim rsTem As New ADODB.Recordset
    Dim rsVisitItem As New ADODB.Recordset
    Dim ThisItem As New clsItem
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblVisitItem WHERE (((tblVisitItem.VisitID)=" & Val(lstVisit.BoundText) & ") AND ((tblVisitItem.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            ThisItem.ID = !ItemID
            If IsAllergic(FormPatient.ID, ThisItem.ID) = True Then
                i = MsgBox(ThisItem.Item & " may be allergic to this patient. It will NOT be added to the list")
            Else
                If rsVisitItem.State = 1 Then rsVisitItem.Close
                temSql = "Select * from tblVisitItem where VisitItemID = 0 "
                rsVisitItem.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                rsVisitItem.AddNew
                rsVisitItem!PatientID = FormPatient.ID
                rsVisitItem!VisitID = FormVisit.VisitID
                rsVisitItem!UserID = UserID
                rsVisitItem!ItemID = !ItemID
                rsVisitItem!ItemCategoryID = !ItemCategoryID
                rsVisitItem!Dose = !Dose
                rsVisitItem!DoseUnitID = !DoseUnitID
                rsVisitItem!FrequencyUnitID = !FrequencyUnitID
                rsVisitItem!Duration = !Duration
                rsVisitItem!DurationUnitID = !DurationUnitID
                rsVisitItem!IssueQuentity = !IssueQuentity
                rsVisitItem!IssueUnitID = !IssueUnitID
                rsVisitItem!Indoor = !Indoor
                rsVisitItem!IndoorOutdoor = !IndoorOutdoor

                rsVisitItem!PItemID = !PItemID
                rsVisitItem!PIssueQuentity = !PIssueQuentity
                rsVisitItem!PIssueUnitID = !PIssueUnitID

                Dim PItem As New clsItem
                Dim temDose As Double
                Dim temDoseUnitID As Long
                
                
                ThisItem.ID = !ItemID
                If ThisItem.IsItem = True Then
                    PItem.ID = ThisItem.ID
                ElseIf ThisItem.IsTradeName = True Then
                    If chkIndoor.Value = 1 Then
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, IndoorDepartmentID, !DoseUnitID, !Dose)
                    Else
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, OutdoorDepartmentID, !DoseUnitID, !Dose)
                    End If
                ElseIf ThisItem.IsGenericName = True Then
                    If chkIndoor.Value = 1 Then
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, IndoorDepartmentID, !DoseUnitID, !Dose)
                    Else
                        PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, OutdoorDepartmentID, !DoseUnitID, !Dose)
                    End If
                Else
                    PItem.ID = 0
                End If
                
                
                
                If PItem.ID <> 0 Then
                    If !IssueUnitID <> 0 Then
                        temDoseUnitID = !IssueUnitID
                    ElseIf !DoseUnitID <> 0 Then
                        temDoseUnitID = !DoseUnitID
                    Else
                        temDoseUnitID = 0
                    End If
                    
                    If PItem.IssueUnitID = !DoseUnitID Then
                        temDose = Val(!Dose)
                    ElseIf PItem.StrengthUnitID = !DoseUnitID Then
                        temDose = Val(!Dose) / PItem.StrengthUnitsPerIssueUnit
                    Else
                        temDose = 1
                    End If
                    
                    If !IssueQuentity <> 0 Then
                        If PItem.IssueUnitID = !IssueUnitID Then
                            rsVisitItem!PIssueQuentity = !IssueQuentity
                        ElseIf PItem.StrengthUnitID = !IssueUnitID Then
                            rsVisitItem!PIssueQuentity = !IssueQuentity / PItem.StrengthUnitsPerIssueUnit
                        Else
                            rsVisitItem!PIssueQuentity = 1
                        End If
                    Else
                        rsVisitItem!PIssueQuentity = CalIssueQty(PItem.ID, temDose, !FrequencyUnitID, !Duration, !DurationUnitID, PItem.MinIQty)
                    End If
                    rsVisitItem!PItemID = PItem.ID
                End If
                
                
                
                
                
                rsVisitItem.Update
            End If
            .MoveNext
        Wend
    End With
    Call FillPrescreption
    Call FillPharmacy
    cmbItem.SetFocus
    On Error Resume Next: SendKeys "{Escape}"
End Sub


Private Sub chkIndoor_Click()
'    If chkIndoor.Value = 1 Then
'        cmbIssueUnit.Visible = True
'        txtIssueQty.Visible = True
'    Else
'        cmbIssueUnit.Visible = False
'        txtIssueQty.Visible = False
'    End If
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDose.Enabled = True Then
            KeyCode = Empty
            txtDose.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbDiagnosis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If AddFavouriteDiagnosis(FormPatient.AgeInMonths, Val(cmbDiagnosis.BoundText)) = True Then
            btnDxAdd_Click
        Else
            btnDxAdd_Click
        End If
    End If
    
End Sub

Private Function AddFavouriteDiagnosis(Age As Long, DiagnosisID As Long) As Boolean
    Dim rsTem1 As New ADODB.Recordset
    AddFavouriteDiagnosis = False
    With rsTem1
        If .State = 1 Then .Close
        temSql = "Select * from tblFavouriteDiagnosis where DiagnosisID = " & DiagnosisID & " AND Deleted = False AND ((FromAge < " & Age & " AND ToAge>" & Age & ") or (ToAge =  " & Age & " ) or (FromAge = " & Age & ") )"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If AlreadyItemAddToVisit(!ItemID, FormVisit.VisitID) = False Then
                If Val(txtWeight.Text) <> 0 Then
                    If AddWeightFavourite(!ItemID, txtWeight.Text) = True Then
                        AddFavouriteDiagnosis = True
                    Else
                        AddFavouriteDiagnosis = AddAgeFavourite(!ItemID, FormPatient.AgeInMonths)
                    End If
                Else
                    AddFavouriteDiagnosis = AddAgeFavourite(!ItemID, FormPatient.AgeInMonths)
                End If
            End If
            .MoveNext
        Wend
    End With
    Call ClearItemValues
    Call FillPrescreption
    Call FillPharmacy
End Function

Private Function AlreadyItemAddToVisit(ItemID As Long, VisitID As Long) As Boolean
    Dim rsTem1 As New ADODB.Recordset
    With rsTem1
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitItem where Deleted = False AND VisitID = " & VisitID & " AND ItemID = " & ItemID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            AlreadyItemAddToVisit = True
        Else
            AlreadyItemAddToVisit = False
        End If
        .Close
    End With
End Function

Private Sub cmbDoseUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbFrequencyUnit.Enabled = True Then
            KeyCode = Empty
            cmbFrequencyUnit.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbDoseUnit.Text = Empty
    End If
End Sub

Private Sub cmbDurationUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If chkIndoor.Enabled = True And chkIndoor.Visible = True Then
            chkIndoor.SetFocus
        ElseIf txtIssueQty.Enabled = True Then
            txtIssueQty.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbDurationUnit.Text = Empty
    End If
End Sub

Private Sub cmbFrequencyUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDuration.Enabled = True Then
            KeyCode = Empty
            txtDuration.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbFrequencyUnit.Text = Empty
    End If
End Sub

Private Sub cmbIssueUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnRxAdd_Click
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbIssueUnit.Text = Empty
    End If
End Sub

Private Sub cmbItem_Change()
    Dim TemItem As New clsItem
    TemItem.ID = Val(cmbItem.BoundText)
    'If optPrescreption.Value = True Then
'    If SSTab1.Tab = 0 Then
'        If temItem.IsItem = True Then
'            cmbCategory.Enabled = False
'            cmbIssueUnit.Enabled = False
'            cmbDoseUnit.Enabled = False
'            cmbCategory.BoundText = temItem.CategoryID
'            cmbIssueUnit.BoundText = temItem.IssueUnitID
'            cmbDoseUnit.BoundText = temItem.IssueUnitID
'        Else
'            cmbCategory.Enabled = True
'            cmbIssueUnit.Enabled = True
'            cmbDoseUnit.Enabled = True
'        End If
'    Else
'        cmbCategory.Enabled = False
'        cmbIssueUnit.Enabled = False
'        cmbDoseUnit.Enabled = False
'        cmbCategory.BoundText = temItem.CategoryID
'        cmbIssueUnit.BoundText = temItem.IssueUnitID
'        cmbDoseUnit.BoundText = temItem.IssueUnitID
'    End If
End Sub

Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If IsNumeric(cmbItem.BoundText) = False Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Val(txtWeight.Text) <> 0 Then
            If AddWeightFavourite(Val(cmbItem.BoundText), txtWeight.Text) = True Then
                Call FillPrescreption
                Call FillPharmacy
                Call ClearItemValues
                cmbItem.SetFocus
                cmbItem.SetFocus
                On Error Resume Next: SendKeys "{Escape}"
            Else
                If AddAgeFavourite(Val(cmbItem.BoundText), FormPatient.AgeInMonths) = True Then
                    Call FillPrescreption
                    Call FillPharmacy
                    Call ClearItemValues
                    cmbItem.SetFocus
                    On Error Resume Next: SendKeys "{Escape}"
                Else
                    If cmbCategory.Enabled = True Then
                        cmbCategory.SetFocus
                    End If
                End If
            End If
        Else
            If AddAgeFavourite(Val(cmbItem.BoundText), FormPatient.AgeInMonths) = True Then
                Call FillPrescreption
                Call FillPharmacy
                Call ClearItemValues
                cmbItem.SetFocus
                On Error Resume Next: SendKeys "{Escape}"
            Else
                cmbCategory.SetFocus
            End If
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbItem.Text = Empty
    ElseIf KeyCode = vbKeyF1 Then
        frmMedicineDetails.Show
        frmMedicineDetails.ZOrder 0
    End If
End Sub

Private Function AddWeightFavourite(ItemID As Long, weight As Double) As Boolean
    AddWeightFavourite = False
    Dim temDose As Double
    Dim temDoseUnitID As Long
    Dim temDuration As Double
    Dim temDurationUnitID As Long
    Dim temFrequencyUnitID As Long
    Dim temIndoor As Boolean
    Dim TemIssueQuentity As Double
    Dim TemIssueUnitID As Long
    Dim TemItemCategoryID As Long
    

    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Min(tblFavouriteItem.Dose) AS MinOfDose, Last(tblFavouriteItem.DoseUnitID) AS LastOfDoseUnitID  , Last(tblFavouriteItem.ItemCategoryID) AS LastOfItemCategoryID , Last(tblFavouriteItem.FrequencyUnitID) AS LastOfFrequencyUnitID, Min(tblFavouriteItem.Duration) AS MinOfDuration, Last(tblFavouriteItem.DurationUnitID) AS LastOfDurationUnitID, Last(tblFavouriteItem.Indoor) AS LastOfIndoor, Min(tblFavouriteItem.IssueQuentity) AS MinOfIssueQuentity, Last(tblFavouriteItem.IssueUnitID) AS LastOfIssueUnitID " & _
                    "FROM tblFavouriteItem " & _
                    "WHERE (((tblFavouriteItem.FromWeight)<" & weight & ") AND ((tblFavouriteItem.ToWeight)>" & weight & ") AND ((tblFavouriteItem.ByWeight)=True) AND ((tblFavouriteItem.Deleted)=False) AND ((tblFavouriteItem.ItemID)=" & ItemID & ")) OR (((tblFavouriteItem.FromWeight)=" & weight & ") AND ((tblFavouriteItem.ByWeight)=True) AND ((tblFavouriteItem.Deleted)=False) AND ((tblFavouriteItem.ItemID)=" & ItemID & ")) OR (((tblFavouriteItem.ToWeight)=" & weight & ") AND ((tblFavouriteItem.ByWeight)=True) AND ((tblFavouriteItem.Deleted)=False) AND ((tblFavouriteItem.ItemID)=" & ItemID & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MinOfDose) = False Then
                temDose = !MinOfDose
            Else
                Exit Function
            End If
            If IsNull(!LastOfDoseUnitID) = False Then
                temDoseUnitID = !LastOfDoseUnitID
            Else
                Exit Function
            End If
            If IsNull(!MinOfDuration) = False Then
                temDuration = !MinOfDuration
            Else
                Exit Function
            End If
            If IsNull(!LastOfDurationUnitID) = False Then
                temDurationUnitID = !LastOfDurationUnitID
            Else
                Exit Function
            End If
            If IsNull(!LastOfFrequencyUnitID) = False Then
                temFrequencyUnitID = !LastOfFrequencyUnitID
            Else
                Exit Function
            End If
            If IsNull(!LastOfIndoor) = False Then
                temIndoor = !LastOfIndoor
            Else
                Exit Function
            End If
            If IsNull(!MinOfIssueQuentity) = False Then
                TemIssueQuentity = !MinOfIssueQuentity
            Else
                Exit Function
            End If
            If IsNull(!LastOfIssueUnitID) = False Then
                TemIssueUnitID = !LastOfIssueUnitID
            Else
                Exit Function
            End If
            If IsNull(!LastOfItemCategoryID) = False Then
                TemItemCategoryID = !LastOfItemCategoryID
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
        .Close
    End With
    
    Dim TemItem As New clsItem
    TemItem.ID = ItemID
    
    If IsAllergic(FormPatient.ID, ItemID) = True Then
        i = MsgBox(TemItem.Item & " may be allergic to this patient will not be added to the list")
        Exit Function
'        If i <> vbYes Then
'            Unload frmPastHistory
'            frmPastHistory.Show
'            frmPastHistory.ZOrder 0
'            frmPastHistory.SSTab1.Tab = 2
'            Exit Function
'        End If
    End If
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitItem where VisitItemID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PatientID = FormPatient.ID
        !VisitID = FormVisit.VisitID
        !UserID = UserID
        !ItemID = ItemID
        !Dose = temDose
        !DoseUnitID = temDoseUnitID
        !FrequencyUnitID = temFrequencyUnitID
        !Duration = temDuration
        !DurationUnitID = temDurationUnitID
        !IssueQuentity = TemIssueQuentity
        !IssueUnitID = TemIssueUnitID
        !ItemCategoryID = TemItemCategoryID
        If temIndoor = True Then
            !Indoor = True
            !IndoorOutdoor = "Indoor"
        Else
            !Indoor = False
            !IndoorOutdoor = "Outdoor"
        End If
        
        Dim ThisItem As New clsItem
        Dim PItem As New clsItem
        Dim temPDose As Double
        Dim temPDoseUnitID As Long
        
        
        ThisItem.ID = ItemID
        If ThisItem.IsItem = True Then
            PItem.ID = ThisItem.ID
        ElseIf ThisItem.IsTradeName = True Then
            If chkIndoor.Value = 1 Then
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, Val(TemItemCategoryID), IndoorDepartmentID, temDoseUnitID, temDose)
            Else
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, Val(TemItemCategoryID), OutdoorDepartmentID, temDoseUnitID, temDose)
            End If
        ElseIf ThisItem.IsGenericName = True Then
            If chkIndoor.Value = 1 Then
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, Val(TemItemCategoryID), IndoorDepartmentID, temDoseUnitID, temDose)
            Else
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, Val(TemItemCategoryID), OutdoorDepartmentID, temDoseUnitID, temDose)
            End If
        Else
            PItem.ID = 0
        End If
        If PItem.ID <> 0 Then
            If TemIssueUnitID <> 0 Then
                temPDoseUnitID = TemIssueUnitID
            ElseIf temDoseUnitID <> 0 Then
                temPDoseUnitID = temDoseUnitID
            Else
                temPDoseUnitID = 0
            End If
            
            If PItem.IssueUnitID = temDoseUnitID Then
                temPDose = Val(temDose)
            ElseIf PItem.StrengthUnitID = temDoseUnitID Then
                temPDose = Val(temDose) / PItem.StrengthUnitsPerIssueUnit
            Else
                temPDose = 0
            End If
            
            If TemIssueQuentity <> 0 Then
                If PItem.IssueUnitID = TemIssueUnitID Then
                    !PIssueQuentity = TemIssueQuentity
                ElseIf PItem.StrengthUnitID = TemIssueUnitID Then
                    !PIssueQuentity = TemIssueQuentity / PItem.StrengthUnitsPerIssueUnit
                Else
                    !PIssueQuentity = 1
                End If
            Else
                !PIssueQuentity = CalIssueQty(PItem.ID, temPDose, temFrequencyUnitID, temDuration, temDurationUnitID, PItem.MinIQty)
            End If
            !PItemID = PItem.ID
        End If
        
        .Update
    End With
    AddWeightFavourite = True
End Function

Private Function AddAgeFavourite(ItemID As Long, Age As Double) As Boolean
    AddAgeFavourite = False
    Dim temDose As Double
    Dim temDoseUnitID As Long
    Dim temDuration As Double
    Dim temDurationUnitID As Long
    Dim temFrequencyUnitID As Long
    Dim temIndoor As Boolean
    Dim TemIssueQuentity As Double
    Dim TemIssueUnitID As Long
    Dim TemItemCategoryID As Long
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Min(tblFavouriteItem.Dose) AS MinOfDose, Last(tblFavouriteItem.DoseUnitID) AS LastOfDoseUnitID  , Last(tblFavouriteItem.ItemCategoryID) AS LastOfItemCategoryID , Last(tblFavouriteItem.FrequencyUnitID) AS LastOfFrequencyUnitID, Min(tblFavouriteItem.Duration) AS MinOfDuration, Last(tblFavouriteItem.DurationUnitID) AS LastOfDurationUnitID, Last(tblFavouriteItem.Indoor) AS LastOfIndoor, Min(tblFavouriteItem.IssueQuentity) AS MinOfIssueQuentity, Last(tblFavouriteItem.IssueUnitID) AS LastOfIssueUnitID " & _
                    "FROM tblFavouriteItem " & _
                    "WHERE (((tblFavouriteItem.FromAge)<" & Age & ") AND ((tblFavouriteItem.ToAge)>" & Age & ") AND ((tblFavouriteItem.ByAge)=True) AND ((tblFavouriteItem.Deleted)=False) AND ((tblFavouriteItem.ItemID)=" & ItemID & ")) OR (((tblFavouriteItem.FromAge)=" & Age & ") AND ((tblFavouriteItem.ByAge)=True) AND ((tblFavouriteItem.Deleted)=False) AND ((tblFavouriteItem.ItemID)=" & ItemID & ")) OR (((tblFavouriteItem.ToAge)=" & Age & ") AND ((tblFavouriteItem.ByAge)=True) AND ((tblFavouriteItem.Deleted)=False) AND ((tblFavouriteItem.ItemID)=" & ItemID & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MinOfDose) = False Then
                temDose = !MinOfDose
            Else
                Exit Function
            End If
            If IsNull(!LastOfDoseUnitID) = False Then
                temDoseUnitID = !LastOfDoseUnitID
            Else
                Exit Function
            End If
            If IsNull(!MinOfDuration) = False Then
                temDuration = !MinOfDuration
            Else
                Exit Function
            End If
            If IsNull(!LastOfDurationUnitID) = False Then
                temDurationUnitID = !LastOfDurationUnitID
            Else
                Exit Function
            End If
            If IsNull(!LastOfFrequencyUnitID) = False Then
                temFrequencyUnitID = !LastOfFrequencyUnitID
            Else
                Exit Function
            End If
            If IsNull(!LastOfIndoor) = False Then
                temIndoor = !LastOfIndoor
            Else
                Exit Function
            End If
            If IsNull(!MinOfIssueQuentity) = False Then
                TemIssueQuentity = !MinOfIssueQuentity
            Else
                Exit Function
            End If
            If IsNull(!LastOfIssueUnitID) = False Then
                TemIssueUnitID = !LastOfIssueUnitID
            Else
                Exit Function
            End If
            If IsNull(!LastOfItemCategoryID) = False Then
                TemItemCategoryID = !LastOfItemCategoryID
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
        .Close
    End With
    If IsAllergic(FormPatient.ID, ItemID) = True Then
    
        
    
        i = MsgBox("The medicine you are going to prescribe May be allergic to this patient. Do you still want to Prescribe this medicine?", vbYesNo)
        If i <> vbYes Then
            Unload frmPastHistory
            frmPastHistory.Show
            frmPastHistory.ZOrder 0
            frmPastHistory.SSTab1.Tab = 2
            Exit Function
        End If
    End If
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitItem where VisitItemID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PatientID = FormPatient.ID
        !VisitID = FormVisit.VisitID
        !UserID = UserID
        !ItemID = ItemID
        !Dose = temDose
        !DoseUnitID = temDoseUnitID
        !FrequencyUnitID = temFrequencyUnitID
        !Duration = temDuration
        !DurationUnitID = temDurationUnitID
        !IssueQuentity = TemIssueQuentity
        !IssueUnitID = TemIssueUnitID
        !ItemCategoryID = TemItemCategoryID
        If temIndoor = True Then
            !Indoor = True
            !IndoorOutdoor = "Indoor"
        Else
            !Indoor = False
            !IndoorOutdoor = "Outdoor"
        End If
        
        Dim ThisItem As New clsItem
        Dim PItem As New clsItem
        Dim temPDose As Double
        Dim temPDoseUnitID As Long
        
        
        ThisItem.ID = ItemID
        If ThisItem.IsItem = True Then
            PItem.ID = ThisItem.ID
        ElseIf ThisItem.IsTradeName = True Then
            If chkIndoor.Value = 1 Then
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, Val(TemItemCategoryID), IndoorDepartmentID, temDoseUnitID, temDose)
            Else
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, Val(TemItemCategoryID), OutdoorDepartmentID, temDoseUnitID, temDose)
            End If
        ElseIf ThisItem.IsGenericName = True Then
            If chkIndoor.Value = 1 Then
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, Val(TemItemCategoryID), IndoorDepartmentID, temDoseUnitID, temDose)
            Else
                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, Val(TemItemCategoryID), OutdoorDepartmentID, temDoseUnitID, temDose)
            End If
        Else
            PItem.ID = 0
        End If
        If PItem.ID <> 0 Then
            If TemIssueUnitID <> 0 Then
                temPDoseUnitID = TemIssueUnitID
            ElseIf temDoseUnitID <> 0 Then
                temPDoseUnitID = temDoseUnitID
            Else
                temPDoseUnitID = 0
            End If
            If PItem.IssueUnitID = temDoseUnitID Then
                temPDose = Val(temDose)
            ElseIf PItem.StrengthUnitID = temDoseUnitID Then
                temPDose = Val(temDose) / PItem.StrengthUnitsPerIssueUnit
            Else
                'temPDose = temdose * doseratio(
            End If
            If TemIssueQuentity <> 0 Then
                If PItem.IssueUnitID = TemIssueUnitID Then
                    !PIssueQuentity = TemIssueQuentity
                ElseIf PItem.StrengthUnitID = TemIssueUnitID Then
                    !PIssueQuentity = TemIssueQuentity / PItem.StrengthUnitsPerIssueUnit
                Else
                    !PIssueQuentity = 1
                End If
            Else
                !PIssueQuentity = CalIssueQty(PItem.ID, temPDose, temFrequencyUnitID, temDuration, temDurationUnitID, PItem.MinIQty)
            End If
            !PItemID = PItem.ID
        End If
        
        
        .Update
    End With
    AddAgeFavourite = True
End Function



Private Sub Form_Activate()
    FormPatient.ID = FormPatient.ID
    Call DisplayPatientDetails
    Call DisplayPastHistory
End Sub

Private Sub Form_Load()
    FormPatient.ID = CurrentPatient.ID
    FormVisit.VisitID = CurrentVisit.VisitID
    
    If FormVisit.VisitDate <> Date Then
        btnDelete.Enabled = False
        btnCharges.Enabled = False
    End If
    
    Call SetColours
    Call FillCombos
    Call DisplayPatientDetails
    Call DisplayPastHistory
    Call ListVisits
    Call FillPrescreption
    Call FillPharmacy
    Call FillDiagnosis
    Call FillVisitDetails
    Call FillIx
    Call FillPrinters
    Call DisplayVisitCharges
    Call GetSettings

    If Date <> TodayStartedDate Then End

End Sub

Private Sub GetSettings()
    chkClose.Value = GetSetting(App.EXEName, Me.Name, chkClose.Name, 1)
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, App.EXEName, "Printer", "")
    cmbPrinter_Click
    cmbPaper.Text = PrescreptionPaperName
    txtScaleWidth.Text = Val(GetSetting(App.EXEName, App.EXEName, "ScaleWidth", 4.1))
    txtScaleHeight.Text = Val(GetSetting(App.EXEName, App.EXEName, "ScaleHeight", 6.3))
    GetCommonSettings Me
    
    cmbDrawer.BoundText = GetSetting(App.EXEName, Me.Name, cmbDrawer.Name, 1)
End Sub

Private Sub DisplayPatientDetails()
    txtPatientDetails.Text = FormPatient.NameWithTitle
    txtPatientDetails.Text = txtPatientDetails.Text & vbNewLine & FormPatient.AgeInWords
    txtPatientDetails.Text = txtPatientDetails.Text & vbNewLine & FormPatient.Address
End Sub

Private Sub DisplayPastHistory()
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientDiagnosis.PatientDiagnosisID, tblPatientDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblPatientDiagnosis ON tblDiagnosis.DiagnosisID = tblPatientDiagnosis.DiagnosisID where tblPatientDiagnosis.Deleted = False AND tblPatientDiagnosis.PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtHistory.Text = "Past Hx: "
            While .EOF = False
                txtHistory.Text = txtHistory.Text & !Diagnosis
                .MoveNext
                If .EOF = False Then txtHistory.Text = txtHistory.Text & ", "
            Wend
        End If
        .Close
    End With
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientItem.PatientItemID, [tblItem.Item] & ' ' & [tblPatientItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] AS Prescreption, 'Issue ' & tblPatientItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM (((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblPatientItem ON tblIssueUnit.ItemUnitID = tblPatientItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblPatientItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblPatientItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblPatientItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblPatientItem.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblPatientItem.PatientID)=" & FormPatient.ID & ") AND ((tblPatientItem.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If txtHistory.Text <> "" Then
                txtHistory.Text = txtHistory.Text & vbNewLine & "Drug Hx: "
            Else
                txtHistory.Text = "Drug Hx: "
            End If
            While .EOF = False
                txtHistory.Text = txtHistory.Text & !Prescreption
                .MoveNext
                If .EOF = False Then txtHistory.Text = txtHistory.Text & ", "
            Wend
        End If
        .Close
    End With
    Set rsTem = Nothing
    
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientAllergy.PatientAllergyID, tblPatientAllergy.Comments, tblItem.Item, tblItem.ItemID FROM tblItem RIGHT JOIN tblPatientAllergy ON tblItem.ItemID = tblPatientAllergy.ItemID where tblPatientAllergy.Deleted = False AND tblPatientAllergy.PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If txtHistory.Text <> "" Then
                txtHistory.Text = txtHistory.Text & vbNewLine & "Allergies: "
            Else
                txtHistory.Text = "Allergies: "
            End If
            While .EOF = False
                If IsNull(!Item) = False Then
                    txtHistory.Text = txtHistory.Text & !Item
                End If
                .MoveNext
                If .EOF = False Then txtHistory.Text = txtHistory.Text & ", "
            Wend
        End If
    End With
    Set rsDx = Nothing
    
    
    
End Sub

Public Sub DisplayVisitCharges()
    FormVisit.Refresh
    txtProcedureFee.Text = Format(FormVisit.ProcedureFee, "0.00")
    txtMedicinesFee.Text = Format(FormVisit.MedicinesFee, "0.00")
    txtConsultationFee.Text = Format(FormVisit.ConsultationFee, "0.00")
    txtTotalCharge.Text = Format(FormVisit.TotalFee, "0.00")
End Sub

Private Sub ListVisits()
    Dim TemVisitID As Long
    With rsPastVisits
        If .State = 1 Then .Close
        temSql = "Select tblVisit.VisitID, format(tblVisit.VisitDate, 'dd MMM yyyy') & ' - ' & SBP & '/' & DBP  & ' - ' & VisitWeight & ' kg' as VisitDisplay from tblVisit where PatientID = " & FormPatient.ID & " AND VisitID <> " & FormVisit.VisitID & " And Deleted = FALSE Order by VisitID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
    End With
    With lstVisit
        Set .RowSource = rsPastVisits
        .ListField = "VisitDisplay"
        .BoundColumn = "VisitID"
        .BoundText = FormVisit.VisitID
        lstVisit_Click
    End With
End Sub

Private Sub gridPharmacy_Click()
    Dim temRow As Integer
    With gridPharmacy
        temRow = .row
        txtDeletedID.Text = .TextMatrix(temRow, 0)
    End With
End Sub

Private Sub PharmacyPrescreptionShift(MakeVisible As Boolean)
    lblDose.Visible = MakeVisible
    lblDuration.Visible = MakeVisible
    lblFrequency.Visible = MakeVisible
    cmbDoseUnit.Visible = MakeVisible
    cmbDurationUnit.Visible = MakeVisible
    cmbFrequencyUnit.Visible = MakeVisible
    txtDose.Visible = MakeVisible
    txtDuration.Visible = MakeVisible
End Sub


Private Sub gridPharmacy_DblClick()
'
'    Call PharmacyPrescreptionShift(True)
    
    Dim temRow As Integer
    Dim EditID As Long
    
'    optPrescreption.Value = True
    
    With gridPharmacy
        temRow = .row
        EditID = Val(.TextMatrix(temRow, 0))
    End With
    Call ClearItemValues
    txtAddID.Text = EditID
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitItem where VIsitItemID = " & EditID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbItem.BoundText = !ItemID
            cmbCategory.BoundText = !ItemCategoryID
            cmbDoseUnit.BoundText = !DoseUnitID
            cmbIssueUnit.BoundText = !IssueUnitID
            cmbFrequencyUnit.BoundText = !FrequencyUnitID
            cmbDurationUnit.BoundText = !DurationUnitID
            txtDose.Text = Val(Format(!Dose, "0.00"))
            txtDuration.Text = Val(Format(!Duration, "0.00"))
            txtIssueQty.Text = !IssueQuentity
            If !Indoor = True Then
                chkIndoor.Value = 1
            Else
                chkIndoor.Value = 0
            End If
            Dim MyPItem As New clsItem
            MyPItem.ID = !PItemID
            cmbPItem.BoundText = !PItemID
            cmbPIssueUnit.BoundText = MyPItem.IssueUnitID
            txtPIssueQty.Text = !PIssueQuentity
        End If
        .Close
    End With
End Sub

Private Sub gridPharmacy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim rsTemAlterItem As New ADODB.Recordset
    Dim i As Integer
    Dim temSelect As String
    Dim mySql As String

    If Shift = 0 Then
        temSelect = "Select top 5"
    Else
        temSelect = "Select "
    End If

        ItemRow = Y \ gridPharmacy.RowHeight(0)
        If ItemRow < 1 Then Exit Sub
        ItemRow = ItemRow + gridPharmacy.TopRow - 1
        TopRow = gridPharmacy.TopRow
        If ItemRow > gridPharmacy.Rows - 1 Then Exit Sub
        BasicItem.ID = Val(gridPharmacy.TextMatrix(ItemRow, 4))
        
        If gridPharmacy.TextMatrix(ItemRow, 1) = "" Then Exit Sub
        
    If Button = vbRightButton Then
        With rsTem
            temSql = temSelect & "   tblItem.ItemID, tblItem.Item   FROM tblItem LEFT JOIN tblVisitItem ON tblItem.ItemID = tblVisitItem.PItemID where tblItem.Deleted = False AND tblItem.IsItemName = True AND tblItem.GenericNameID = " & BasicItem.GenericNameID & "  GROUP BY tblItem.ItemID, tblItem.Item ORDER BY Count(tblVisitItem.PItemID) DESC"
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                ReDim AlternateItem(.RecordCount - 1 + 3 + 3) As String
                ReDim AlternateItemID(.RecordCount - 1 + 3 + 3) As Long
                AlternateItemCount = .RecordCount
                .MoveFirst
                For i = 0 To .RecordCount - 1
                    AlternateItem(i) = !Item
                    AlternateItemID(i) = !ItemID
                    .MoveNext
                Next
            End If
            AlternateItem(.RecordCount) = "-"
            AlternateItemID(.RecordCount) = 0
            AlternateItem(.RecordCount + 1) = "Delete"
            AlternateItemID(.RecordCount + 1) = -1
            AlternateItem(.RecordCount + 2) = "Edit"
            AlternateItemID(.RecordCount + 2) = -2
            AlternateItem(.RecordCount + 3) = "-"
            AlternateItemID(.RecordCount + 3) = 0
            AlternateItem(.RecordCount + 4) = "Indoor"
            AlternateItemID(.RecordCount + 4) = -3
            AlternateItem(.RecordCount + 5) = "Outdoor"
            AlternateItemID(.RecordCount + 5) = -4
            
            .Close
            
        End With
        
        ucpPrescreption.Popup AlternateItem()
        
        AlternatePharmacyItem (ucpPrescreption.ReturnIndex)

    End If


End Sub

Private Sub gridPrescreption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim rsTemAlterItem As New ADODB.Recordset
    Dim i As Integer
    Dim temSelect As String

    If Shift = 0 Then
        temSelect = "Select top 5"
    Else
        temSelect = "Select "
    End If

    If Button = vbRightButton Then
        ItemRow = Y \ gridPrescreption.RowHeight(0)
        If ItemRow < 1 Then Exit Sub
        ItemRow = ItemRow + gridPrescreption.TopRow - 1
        TopRow = gridPrescreption.TopRow
        If ItemRow > gridPrescreption.Rows - 1 Then Exit Sub
        BasicItem.ID = Val(gridPrescreption.TextMatrix(ItemRow, 3))
        With rsTemAlterItem
            If .State = 1 Then .Close
            
            temSql = temSelect & " Count(Prescreption) as CountOfPrescreption, Prescreption, ItemCategoryID, Dose, DoseUnitID, FrequencyUnitID " & _
                        "FROM ( SELECT [tblItem.Item] & ' ' & tblCategory.Category & ' ' & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] AS Prescreption, tblVisitItem.ItemCategoryID, tblVisitItem.Dose, tblVisitItem.DoseUnitID, tblVisitItem.FrequencyUnitID " & _
                        "FROM (((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID) LEFT JOIN tblPatient ON tblVisitItem.PatientID = tblPatient.PatientID " & _
                        "WHERE (((tblVisitItem.ItemID)=" & BasicItem.ID & ") AND ((tblVisitItem.Deleted)=False) AND ((tblPatient.DateOfBirth) Between #" & Format(DateSerial(Year(FormPatient.DateOfBirth) - (FormPatient.AgeInYears \ 2), 1, 1), "dd MMMM yyyy") & "# And #" & Format(DateSerial(Year(FormPatient.DateOfBirth) + (FormPatient.AgeInYears \ 2), 12, 31), "dd MMMM yyyy") & "#))  ) " & _
                        "GROUP BY Prescreption,  ItemCategoryID, Dose, DoseUnitID, FrequencyUnitID " & _
                        "ORDER BY Count(Prescreption) DESC"
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                NewDoseCount = .RecordCount
                ReDim ItemCatID(NewDoseCount)
                ReDim Dose(NewDoseCount)
                ReDim DoseUnitID(NewDoseCount)
                ReDim FrequencyID(NewDoseCount)
                ReDim NewDose(NewDoseCount)
                .MoveFirst
                For i = 0 To NewDoseCount - 1
                    If IsNull(!Prescreption) = False Then NewDose(i) = !Prescreption
                    If IsNull(!Dose) = False Then Dose(i) = !Dose
                    If IsNull(!DoseUnitID) = False Then DoseUnitID(i) = !DoseUnitID
                    If IsNull(!ItemCategoryID) = False Then ItemCatID(i) = !ItemCategoryID
                    If IsNull(!FrequencyUnitID) = False Then FrequencyID(i) = !FrequencyUnitID
                    .MoveNext
                Next i
            End If
            
            .Close
            
            If BasicItem.IsGenericName = True Or BasicItem.IsTradeName = True Or BasicItem.IsItem = True Then
                temSql = temSelect & "   tblItem.ItemID, tblItem.Item, Sum(tblItemBatchStock.Stock) AS SumOfStock  " & _
                    "FROM tblItem RIGHT JOIN tblItemBatchStock ON tblItem.ItemID = tblItemBatchStock.ItemID " & _
                    "WHERE tblItem.Deleted = False AND tblItemBatchStock.DepartmentID = " & OutdoorDepartmentID & " AND GenericNameID = " & BasicItem.GenericNameID & _
                    " GROUP BY tblItem.ItemID, tblItem.Item " & _
                    "ORDER BY Sum(tblItemBatchStock.Stock) DESC"
            ElseIf BasicItem.IsGroup = True Then
                temSql = temSelect & " tblItem.ItemID, tblItem.Item, Sum(tblItemBatchStock.Stock) AS SumOfStock   " & _
                    "FROM (tblItem LEFT JOIN tblVisitItem ON tblItem.ItemID = tblVisitItem.ItemID) LEFT JOIN tblItemBatchStock ON tblItem.ItemID = tblItemBatchStock.ItemID " & _
                    "where tblItem.Deleted = False AND tblItemBatchStock.DepartmentID = " & OutdoorDepartmentID & " AND GroupID = " & BasicItem.ID & _
                    " GROUP BY tblItem.ItemID, tblItem.Item " & _
                    "ORDER BY Sum(tblItemBatchStock.Stock) DESC"
            ElseIf BasicItem.IsSubGroup = True Then
                temSql = temSelect & " tblItem.ItemID, tblItem.Item, Sum(tblItemBatchStock.Stock) AS SumOfStock   " & _
                    "FROM (tblItem LEFT JOIN tblVisitItem ON tblItem.ItemID = tblVisitItem.ItemID) LEFT JOIN tblItemBatchStock ON tblItem.ItemID = tblItemBatchStock.ItemID " & _
                    "where tblItem.Deleted = False AND tblItemBatchStock.DepartmentID = " & OutdoorDepartmentID & " AND SubGroupID = " & BasicItem.ID & _
                    " GROUP BY tblItem.ItemID, tblItem.Item " & _
                    "ORDER BY Sum(tblItemBatchStock.Stock) DESC"
            End If
            
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                ReDim AlternateItem(.RecordCount - 1 + 3 + 3 + NewDoseCount + 1) As String
                ReDim AlternateItemID(.RecordCount - 1 + 3 + 3 + NewDoseCount + 1) As Long
                AlternateItemCount = .RecordCount
                .MoveFirst
                For i = 0 To .RecordCount - 1
                    AlternateItem(i) = !Item & " (" & Format(!SumOfStock, "0") & ")"
                    AlternateItemID(i) = !ItemID
                    .MoveNext
                Next
            End If
            AlternateItem(.RecordCount) = "-"
            AlternateItemID(.RecordCount) = 0
            AlternateItem(.RecordCount + 1) = "Delete"
            AlternateItemID(.RecordCount + 1) = -1
            AlternateItem(.RecordCount + 2) = "Edit"
            AlternateItemID(.RecordCount + 2) = -2
            AlternateItem(.RecordCount + 3) = "-"
            AlternateItemID(.RecordCount + 3) = 0
            AlternateItem(.RecordCount + 4) = "Indoor"
            AlternateItemID(.RecordCount + 4) = -3
            AlternateItem(.RecordCount + 5) = "Outdoor"
            AlternateItemID(.RecordCount + 5) = -4
            .Close
            If NewDoseCount > 0 Then
                AlternateItem(AlternateItemCount + 6) = "-"
                AlternateItemID(AlternateItemCount + 6) = -5
                For i = 0 To NewDoseCount - 1
                    AlternateItem(AlternateItemCount + 7 + i) = NewDose(i)
                    AlternateItemID(AlternateItemCount + 7 + i) = 0 - 6 - i
                Next
            Else
            End If
        End With
        ucpPrescreption.Popup AlternateItem()
        RxPopAction (ucpPrescreption.ReturnIndex)
    End If
End Sub

Private Sub AlternatePharmacyItem(RxPopUpReturnIndex As Long)
    Dim rsTemAlterItem As New ADODB.Recordset
    If RxPopUpReturnIndex < 0 Then Exit Sub
    If AlternateItemID(RxPopUpReturnIndex) > 0 Then
        With rsTemAlterItem
            temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !PItemID = AlternateItemID(RxPopUpReturnIndex)
                .Update
            End If
            .Close
        End With
        Call ClearItemValues
        Call FillPrescreption
        Call FillPharmacy
        gridPrescreption.TopRow = TopRow
    ElseIf AlternateItemID(RxPopUpReturnIndex) = -1 Then
        txtDeletedID.Text = gridPrescreption.TextMatrix(ItemRow, 0)
        btnRxDelete_Click
    ElseIf AlternateItemID(RxPopUpReturnIndex) = -2 Then
        gridPrescreption.row = ItemRow
        gridPrescreption_DblClick
    ElseIf AlternateItemID(RxPopUpReturnIndex) = -3 Then
        With rsTemAlterItem
            temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Indoor = True
                !IndoorOutdoor = "Indoor"
                .Update
            End If
            .Close
            Call ClearItemValues
            Call FillPrescreption
            Call FillPharmacy
            gridPrescreption.TopRow = TopRow
        End With
    ElseIf AlternateItemID(RxPopUpReturnIndex) = -4 Then
        With rsTemAlterItem
            temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Indoor = False
                !IndoorOutdoor = "Outdoor"
                .Update
            End If
            .Close
            Call ClearItemValues
            Call FillPrescreption
            Call FillPharmacy
            gridPrescreption.TopRow = TopRow
        End With
    End If
End Sub

Private Sub RxPopAction(RxPopUpReturnIndex As Long)
        Dim rsTemAlterItem As New ADODB.Recordset
        
        Dim ThisItem As New clsItem
        Dim PItem As New clsItem
        Dim temDose As Double
        Dim temDoseUnitID As Long
        
        If SSTab1.Tab = 1 Then
            Call AlternatePharmacyItem(RxPopUpReturnIndex)
            Exit Sub
        End If
        
        
        If RxPopUpReturnIndex >= 0 Then
            
            If AlternateItemID(RxPopUpReturnIndex) > 0 Then
                With rsTemAlterItem
                    temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !ItemID = AlternateItemID(RxPopUpReturnIndex)
                        ThisItem.ID = AlternateItemID(RxPopUpReturnIndex)
                        
                        If ThisItem.IsItem = True Then
                            PItem.ID = ThisItem.ID
                        ElseIf ThisItem.IsTradeName = True Then
                            If !Indoor = True Then
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, IndoorDepartmentID, ThisItem.StrengthUnitID, ThisItem.StrengthUnitsPerIssueUnit)
                            Else
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, OutdoorDepartmentID, ThisItem.StrengthUnitID, ThisItem.StrengthUnitsPerIssueUnit)
                            End If
                        ElseIf ThisItem.IsGenericName = True Then
                            If !Indoor = True = 1 Then
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, !ItemCategoryID, IndoorDepartmentID, ThisItem.StrengthUnitID, ThisItem.StrengthUnitsPerIssueUnit)
                            Else
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, !ItemCategoryID, OutdoorDepartmentID, ThisItem.StrengthUnitID, ThisItem.StrengthUnitsPerIssueUnit)
                            End If
                        Else
                            PItem.ID = 0
                        End If
                        
                        If PItem.ID <> 0 Then
                            If !IssueUnitID <> 0 Then
                                temDoseUnitID = !IssueUnitID
                            ElseIf !DoseUnitID <> 0 Then
                                temDoseUnitID = !DoseUnitID
                            Else
                                temDoseUnitID = 0
                            End If
                            If PItem.IssueUnitID = temDoseUnitID Then
                                temDose = Val(Format(!Dose, "0.00"))
                            ElseIf PItem.StrengthUnitID = temDoseUnitID Then
                                temDose = !Dose / PItem.StrengthUnitsPerIssueUnit
                            Else
                                temDose = 1
                            End If
                            If Val(!IssueUnitID) <> 0 And Val(!IssueQuentity) <> 0 Then
                                If PItem.IssueUnitID = Val(!IssueUnitID) Then
                                    !PIssueQuentity = Val(!IssueQuentity)
                                ElseIf PItem.StrengthUnitID = !IssueUnitID Then
                                    !PIssueQuentity = !IssueQuentity / PItem.StrengthUnitsPerIssueUnit
                                Else
                                    !PIssueQuentity = 1
                                End If
                            Else
                                !PIssueQuentity = CalIssueQty(PItem.ID, temDose, !FrequencyUnitID, CDbl(!Duration), Val(!DurationUnitID), PItem.MinIQty)
                            End If
                            !PItemID = PItem.ID
                        End If
                        .Update
                        Call ClearItemValues
                        Call FillPrescreption
                        Call FillPharmacy
                        gridPrescreption.TopRow = TopRow
                    End If
                    .Close
                End With
            ElseIf AlternateItemID(RxPopUpReturnIndex) = -1 Then
                txtDeletedID.Text = gridPrescreption.TextMatrix(ItemRow, 0)
                btnRxDelete_Click
            ElseIf AlternateItemID(RxPopUpReturnIndex) = -2 Then
                gridPrescreption.row = ItemRow
                gridPrescreption_DblClick

            ElseIf AlternateItemID(RxPopUpReturnIndex) = -3 Then
                With rsTemAlterItem
                    temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !Indoor = True
                        !IndoorOutdoor = "Indoor"
                        .Update
                        Call ClearItemValues
                        Call FillPrescreption
                        Call FillPharmacy
                        gridPrescreption.TopRow = TopRow
                    End If
                    .Close
                End With
            ElseIf AlternateItemID(RxPopUpReturnIndex) = -4 Then
                With rsTemAlterItem
                    temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !Indoor = False
                        !IndoorOutdoor = "Outdoor"
                        .Update
                        Call ClearItemValues
                        Call FillPrescreption
                        Call FillPharmacy
                        gridPrescreption.TopRow = TopRow
                    End If
                    .Close
                End With
            ElseIf AlternateItemID(RxPopUpReturnIndex) < -5 Then
                
                With rsTemAlterItem
                    temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !ItemCategoryID = ItemCatID(RxPopUpReturnIndex - AlternateItemCount - 7)
                        !Dose = Dose(RxPopUpReturnIndex - AlternateItemCount - 7)
                        !DoseUnitID = DoseUnitID(RxPopUpReturnIndex - AlternateItemCount - 7)
                        !FrequencyUnitID = FrequencyID(RxPopUpReturnIndex - AlternateItemCount - 7)
                        
                        ThisItem.ID = BasicItem.ID
                        
                        If ThisItem.IsItem = True Then
                            PItem.ID = ThisItem.ID
                        ElseIf ThisItem.IsTradeName = True Then
                            If !Indoor = True Then
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, IndoorDepartmentID, !DoseUnitID, !Dose)
                            Else
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, ThisItem.ID, !ItemCategoryID, OutdoorDepartmentID, !DoseUnitID, !Dose)
                            End If
                        ElseIf ThisItem.IsGenericName = True Then
                            If !Indoor = True = 1 Then
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, !ItemCategoryID, IndoorDepartmentID, !DoseUnitID, !Dose)
                            Else
                                PItem.ID = PharmacyItem(ThisItem.GenericNameID, 0, !ItemCategoryID, OutdoorDepartmentID, !DoseUnitID, !Dose)
                            End If
                        Else
                            PItem.ID = 0
                        End If
                        
                        If PItem.ID <> 0 Then
                            If !IssueUnitID <> 0 Then
                                temDoseUnitID = !IssueUnitID
                            ElseIf !DoseUnitID <> 0 Then
                                temDoseUnitID = !DoseUnitID
                            Else
                                temDoseUnitID = 0
                            End If
                            If PItem.IssueUnitID = temDoseUnitID Then
                                temDose = !Dose
                            ElseIf PItem.StrengthUnitID = temDoseUnitID Then
                                temDose = !Dose / PItem.StrengthUnitsPerIssueUnit
                            Else
                                temDose = 1
                            End If
                            If Val(!IssueUnitID) <> 0 Then
                                If PItem.IssueUnitID = Val(!IssueUnitID) Then
                                    !PIssueQuentity = Val(!IssueQuentity)
                                ElseIf PItem.StrengthUnitID = !IssueUnitID Then
                                    !PIssueQuentity = !IssueQuentity / PItem.StrengthUnitsPerIssueUnit
                                Else
                                    !PIssueQuentity = 1
                                End If
                            Else
                                !PIssueQuentity = CalIssueQty(PItem.ID, temDose, !FrequencyUnitID, CDbl(!Duration), Val(!DurationUnitID), PItem.MinIQty)
                            End If
                            !PItemID = PItem.ID
                        End If
                        
                        .Update
                        
                        Call ClearItemValues
                        Call FillPrescreption
                        Call FillPharmacy
                        
                        gridPrescreption.TopRow = TopRow
                    
                    End If
                    
                    .Close
                
                End With
            End If
        End If
End Sub

'Private Sub gridPrescreption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim ItemRow As Long
'    Dim TopRow As Long
'    Dim AlternateItem() As String
'    Dim AlternateItemID() As Long
'    Dim rsTemAlterItem As New ADODB.Recordset
'    Dim BasicItem As New clsItem
'    Dim i As Integer
'
'    If Button = vbRightButton Then
'
'        ItemRow = Y \ gridPrescreption.RowHeight(0)
'
'        If ItemRow < 1 Then Exit Sub
'
'        ItemRow = ItemRow + gridPrescreption.TopRow - 1
'        TopRow = gridPrescreption.TopRow
'
'        If ItemRow > gridPrescreption.Rows - 1 Then Exit Sub
'
'        BasicItem.ID = Val(gridPrescreption.TextMatrix(ItemRow, 3))
'
'        If BasicItem.IsGenericName = True Or BasicItem.IsTradeName = True Or BasicItem.IsItem = True Then
'            temSql = "Select * from tblItem where Deleted = False AND GenericNameID = " & BasicItem.GenericNameID & "  Order BY Item"
'        ElseIf BasicItem.IsGroup = True Then
'            temSql = "Select * from tblItem where Deleted = False AND GroupID = " & BasicItem.ID & "  Order BY Item"
'        ElseIf BasicItem.IsSubGroup = True Then
'            temSql = "Select * from tblItem where Deleted = False AND SubGroupID = " & BasicItem.ID & "  Order BY Item"
'        End If
'
'        With rsTemAlterItem
'            If .State = 1 Then .Close
'            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'            If .RecordCount > 0 Then
'                .MoveLast
'                ReDim AlternateItem(.RecordCount - 1 + 3 + 3) As String
'                ReDim AlternateItemID(.RecordCount - 1 + 3 + 3) As Long
'                .MoveFirst
'                For i = 0 To .RecordCount - 1
'                    AlternateItem(i) = !Item
'                    AlternateItemID(i) = !ItemID
'                    .MoveNext
'                Next
'            End If
'            AlternateItem(.RecordCount) = "-"
'            AlternateItemID(.RecordCount) = 0
'            AlternateItem(.RecordCount + 1) = "Delete"
'            AlternateItemID(.RecordCount + 1) = -1
'            AlternateItem(.RecordCount + 2) = "Edit"
'            AlternateItemID(.RecordCount + 2) = -2
'            AlternateItem(.RecordCount + 3) = "-"
'            AlternateItemID(.RecordCount + 3) = 0
'            AlternateItem(.RecordCount + 4) = "Indoor"
'            AlternateItemID(.RecordCount + 4) = -3
'            AlternateItem(.RecordCount + 5) = "Outdoor"
'            AlternateItemID(.RecordCount + 5) = -4
'
'            .Close
'        End With
'        ucPopup.Popup AlternateItem
'        If ucPopup.ReturnIndex >= 0 Then
'            If AlternateItemID(ucPopup.ReturnIndex) > 0 Then
'                With rsTemAlterItem
'                    temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
'                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        !ItemID = AlternateItemID(ucPopup.ReturnIndex)
'                        .Update
'                        Call ClearItemValues
'                        Call FillPrescreption
'                        gridPrescreption.TopRow = TopRow
'                    End If
'                    .Close
'                End With
'            ElseIf AlternateItemID(ucPopup.ReturnIndex) = -1 Then
'                txtDeletedID.Text = gridPrescreption.TextMatrix(ItemRow, 0)
'                btnRxDelete_Click
'            ElseIf AlternateItemID(ucPopup.ReturnIndex) = -2 Then
'                gridPrescreption.Row = ItemRow
'                gridPrescreption_DblClick
'
'            ElseIf AlternateItemID(ucPopup.ReturnIndex) = -3 Then
'                With rsTemAlterItem
'                    temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
'                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        !Indoor = True
'                        .Update
'                        Call ClearItemValues
'                        Call FillPrescreption
'                        gridPrescreption.TopRow = TopRow
'                    End If
'                    .Close
'                End With
'            ElseIf AlternateItemID(ucPopup.ReturnIndex) = -4 Then
'                With rsTemAlterItem
'                    temSql = "Select * from tblVisitItem where VisitItemID = " & Val(gridPrescreption.TextMatrix(ItemRow, 0))
'                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        !Indoor = False
'                        .Update
'                        Call ClearItemValues
'                        Call FillPrescreption
'                        gridPrescreption.TopRow = TopRow
'                    End If
'                    .Close
'                End With
'
'            End If
'        End If
'    End If
'
'End Sub

Private Sub lstVisit_Click()
    On Error Resume Next
    Dim temText As String
    Dim rsVisit As New ADODB.Recordset
    If IsNumeric(lstVisit.BoundText) = False Then Exit Sub
    With rsVisit
        If .State = 1 Then .Close
        temSql = "SELECT tblVisitDiagnosis.VisitDiagnosisID, tblVisitDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblVisitDiagnosis ON tblDiagnosis.DiagnosisID = tblVisitDiagnosis.DiagnosisID where tblVisitDiagnosis.Deleted = False AND tblVisitDiagnosis.VisitID = " & Val(lstVisit.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temText = temText & !Diagnosis & ", "
                .MoveNext
            Wend
            temText = temText & vbNewLine
        End If
    End With
    With rsVisit
        If .State = 1 Then .Close
        temSql = "SELECT [tblItem.Item] AS Prescreption " & _
                    "FROM (((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblVisitItem.VisitID)=" & Val(lstVisit.BoundText) & ") AND ((tblVisitItem.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            temText = temText & !Prescreption & ", "
            .MoveNext
        Wend
    End With
    With rsVisit
        If .State = 1 Then .Close
        temSql = "SELECT tblVisitProcedure.VisitProcedureID, tblProcedure.Procedure, tblProcedure.ProcedureID FROM tblProcedure RIGHT JOIN tblVisitProcedure ON tblProcedure.ProcedureID = tblVisitProcedure.ProcedureID where tblVisitProcedure.Deleted = False AND tblVisitProcedure.VisitID = " & Val(lstVisit.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temText = temText & !Procedure & ", "
                .MoveNext
            Wend
            temText = temText & vbNewLine
        End If
    End With
    With rsVisit
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitReferral where Deleted = False AND VisitID = " & Val(lstVisit.BoundText) & " order by VisitReferralID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            temText = temText & !VisitReferral & ", "
            .MoveNext
        Wend
        temText = temText & vbNewLine
    End With
    
    txtVisit.Text = temText
End Sub

Private Sub FillCombos()
    Dim Item As New clsFillCombos
    Item.FillAnyCombo cmbItem, "Item", True
    Dim Dose As New clsFillCombos
    
    ' Change1
    Dose.FillBoolCombo cmbDoseUnit, "ItemUnit", "ItemUnit", "IsDoseUnit", True
    ' Change1
    Dim Drawer As New clsFillCombos
    Drawer.FillAnyCombo cmbDrawer, "Drawer", True
    Dim Duration As New clsFillCombos
    Duration.FillBoolCombo cmbDurationUnit, "ItemUnit", "ItemUnit", "IsDurationUnit", True
    Dim Frequency As New clsFillCombos
    Frequency.FillBoolCombo cmbFrequencyUnit, "ItemUnit", "ItemUnit", "IsFrequencyUnit", True
    Dim Issue As New clsFillCombos
    Issue.FillBoolCombo cmbIssueUnit, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "Category", "Category", "IsItemCategory", True
    Dim Diag As New clsFillCombos
    Diag.FillAnyCombo cmbDiagnosis, "Diagnosis", True
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
    Dim ItemName As New clsFillCombos
    ItemName.FillBoolCombo cmbPItem, "Item", "Item", "IsItemName", True
    Dim PIssue As New clsFillCombos
    PIssue.FillBoolCombo cmbPIssueUnit, "ItemUnit", "ItemUnit", "IsIssueUnit", True
End Sub

Private Sub FillPrescreption()
    With gridPrescreption
        .Rows = 1
        .Cols = 4
        .Clear
        .row = 0
        .col = 0
        .Text = "ID"
        .col = 1
        .Text = "How to prescribe"
        .col = 2
        .Text = "Indoor Issue Quentity"
        .ColWidth(0) = 0
        .ColWidth(3) = 0
        .ColWidth(2) = 2400
        .ColWidth(1) = .Width - 3000
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblVisitItem.VisitItemID, tblVisitItem.ItemID, tblVisitItem.Indoor,  [tblItem.Item] & ' '   &  tblCategory.Category   & ' '   & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblVisitItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] & ' (' & [tblVisitItem].[IndoorOutdoor] & ')' AS Prescreption, 'Issue ' & tblVisitItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
                    "WHERE (((tblVisitItem.VisitID)=" & FormVisit.VisitID & ") AND ((tblVisitItem.Deleted)=False))" & _
                    "Order by tblVisitItem.VisitItemID"
                    
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridPrescreption.Rows = gridPrescreption.Rows + 1
            gridPrescreption.row = gridPrescreption.Rows - 1
            gridPrescreption.col = 0
            gridPrescreption.Text = ![VisitItemID]
            gridPrescreption.col = 1
            gridPrescreption.Text = !Prescreption
            gridPrescreption.col = 2
            If !Indoor = True Then
                gridPrescreption.Text = !ToIssue
            End If
            gridPrescreption.col = 3
            gridPrescreption.Text = !ItemID
            .MoveNext
        Wend
    End With
    With gridPrescreption
        If .row > 7 Then
            .TopRow = .row - 7
        End If
    End With
End Sub

Private Sub SetColours1()
    PrescreptionIndoorForeColour = RGB(255, 255, 0)
    PrescreptionOutDoorForeColour = RGB(100, 255, 100)
    
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


Private Sub FillPharmacy()
Dim temPItem As New clsItem
    Dim billTotal As Double

    Dim rsPx As New ADODB.Recordset
    With gridPharmacy
        .Clear
        
        .Rows = 1
        .Cols = 8
        
        .row = 0
        
        .col = 0
        .Text = "ID"
        
        .col = 1
        .Text = "Item"
        
        .col = 2
        .Text = "Issue Qty"
        
        .col = 3
        .Text = "VisitItemID"
        
        .col = 4
        .Text = "ItemID"
        
        .col = 5
        .Text = "Stock"
        
        .col = 6
        .Text = "Rate"
        
        .col = 7
        .Text = "Price"
        
        .ColWidth(0) = 0
        .ColWidth(1) = 3600
        .ColWidth(2) = 1200
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        
    End With
    With rsPx
        If .State = 1 Then .Close
        
        temSql = "SELECT tblVisitItem.VisitItemID, tblVisitItem.ItemID, tblVisitItem.PItemID, tblVisitItem.Indoor,  tblItem.Item, tblVisitItem.PIssueQuentity " & _
            "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.PItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
            "WHERE (((tblVisitItem.VisitID)=" & FormVisit.VisitID & ") AND ((tblVisitItem.Deleted)=False))" & _
            "Order by tblVisitItem.VisitItemID"

        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                gridPharmacy.Rows = gridPharmacy.Rows + 1
                gridPharmacy.row = gridPharmacy.Rows - 1
                
                    If !Indoor = True Then
                        gridPharmacy.col = 0
                        gridPharmacy.CellForeColor = PrescreptionIndoorForeColour
                        gridPharmacy.Text = !VisitItemID
                        
                        gridPharmacy.row = gridPharmacy.Rows - 1
                        gridPharmacy.col = 1
                        gridPharmacy.CellForeColor = RGB(200, 100, 200) 'PrescreptionIndoorForeColour
                        gridPharmacy.Text = Format(!Item, "")
                        
                        gridPharmacy.col = 2
                        gridPharmacy.CellForeColor = PrescreptionIndoorForeColour
                        gridPharmacy.Text = !PIssueQuentity
                        gridPharmacy.col = 3
                        gridPharmacy.CellForeColor = PrescreptionIndoorForeColour
                        gridPharmacy.Text = !VisitItemID
                        
                        gridPharmacy.col = 4
                        gridPharmacy.Text = !ItemID
                        
                        gridPharmacy.col = 5
                        gridPharmacy.Text = CalculateStock(!PItemID, , IndoorDepartmentID).Stock
                        
                         If AddIndoorMedicineChargesToFee = True Then
                            
                            temPItem.ID = !PItemID
                            
                            gridPharmacy.col = 6
                            gridPharmacy.Text = Format(temPItem.SalePrice, "0.00")
                            
                            gridPharmacy.col = 7
                            gridPharmacy.Text = Format(!PIssueQuentity * temPItem.SalePrice, "0.00")
                            
                            billTotal = billTotal + (!PIssueQuentity * temPItem.SalePrice)
                        End If
                        
                        
                    Else
                        gridPharmacy.col = 0
                        gridPharmacy.CellForeColor = PrescreptionOutDoorForeColour
                        gridPharmacy.Text = !VisitItemID
                        gridPharmacy.col = 1
                        gridPharmacy.CellForeColor = PrescreptionOutDoorForeColour
                        gridPharmacy.Text = Format(!Item, "")
                        gridPharmacy.col = 2
                        gridPharmacy.CellForeColor = PrescreptionOutDoorForeColour
                        gridPharmacy.Text = !PIssueQuentity
                        gridPharmacy.col = 3
                        gridPharmacy.CellForeColor = PrescreptionOutDoorForeColour
                        gridPharmacy.Text = !VisitItemID
    
                        gridPharmacy.col = 4
                        gridPharmacy.Text = !ItemID
                        
                        gridPharmacy.col = 5
                        gridPharmacy.Text = CalculateStock(!PItemID, , OutdoorDepartmentID).Stock
                        
                        If AddMedicineChargesToFee = True Then
                            
                            temPItem.ID = !PItemID
                            
                            gridPharmacy.col = 6
                            gridPharmacy.Text = Format(temPItem.SalePrice, "0.00")
                            
                            gridPharmacy.col = 7
                            gridPharmacy.Text = Format(!PIssueQuentity * temPItem.SalePrice, "0.00")
                            
                            billTotal = billTotal + (!PIssueQuentity * temPItem.SalePrice)
                        End If
                        
                    
                    End If
                                        
                .MoveNext
            
            
            Wend
            
            gridPharmacy.Rows = gridPharmacy.Rows + 1
            gridPharmacy.row = gridPharmacy.Rows - 1
            
            gridPharmacy.col = 1
            gridPharmacy.Text = "App. Bill Total"
            
            gridPharmacy.col = 7
            gridPharmacy.Text = Format(billTotal, "0.00")
            
            If AddMedicineChargesToFee = True Then
                If billTotal Mod 10 > 5 Then
                    txtMedicinesFee.Text = ((billTotal + 10) \ 10) * 10
                Else
                    txtMedicinesFee.Text = (billTotal \ 10) * 10
                End If
            End If
            
        End If
        .Close
    End With
End Sub

Private Sub FillDiagnosis()
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblVisitDiagnosis.VisitDiagnosisID, tblVisitDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblVisitDiagnosis ON tblDiagnosis.DiagnosisID = tblVisitDiagnosis.DiagnosisID where tblVisitDiagnosis.Deleted = False AND tblVisitDiagnosis.VisitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridDiagnosis.Clear
        gridDiagnosis.Rows = 1
        gridDiagnosis.Cols = 4
        gridDiagnosis.ColWidth(0) = 0
        gridDiagnosis.ColWidth(1) = 3000
        gridDiagnosis.ColWidth(2) = gridDiagnosis.Width - 3000 - 150
        gridDiagnosis.ColWidth(3) = 0
        gridDiagnosis.row = 0
        gridDiagnosis.col = 1
        gridDiagnosis.Text = "Diagnosis"
        gridDiagnosis.col = 2
        gridDiagnosis.Text = "Details"
        
        While .EOF = False
            gridDiagnosis.Rows = gridDiagnosis.Rows + 1
            gridDiagnosis.row = gridDiagnosis.Rows - 1
            gridDiagnosis.col = 0
            gridDiagnosis.Text = !VisitDiagnosisID
            gridDiagnosis.col = 1
            gridDiagnosis.Text = !Diagnosis
            gridDiagnosis.col = 2
            gridDiagnosis.Text = Format(!Comments, "")
            gridDiagnosis.col = 3
            gridDiagnosis.Text = !DiagnosisID
            .MoveNext
        Wend
    End With
    Set rsDx = Nothing
End Sub

Private Sub FillVisitDetails()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblVisit where VisitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtVisitDate.Text = Format(!VisitDate, LongDateFormat)
            txtVisitTime.Text = Format(!VisitTime, "HH:MM AMPM")
            txtConsultationFee.Text = Format(!ConsultationFee, "0.00")
            txtProcedureFee.Text = Format(!ProcedureFee, "0.00")
            txtMedicinesFee.Text = Format(!MedicineFee, "0.00")
            txtTotalCharge.Text = Format(!TotalFee, "0.00")
            txtComments.Text = Format(!Comments, "")
            If !SBP <> 0 Then txtSBP.Text = !SBP
            If !DBP <> 0 Then txtDBP.Text = !DBP
            If !VisitWeight <> 0 Then txtWeight.Text = !VisitWeight
            If IsNull(!height) = False Then
                If (!height) > 0 Then
                    txtHt.Text = !height
                End If
            End If
            If IsNull(!pr) = False Then
                If (!pr) > 0 Then
                    txtPr.Text = !pr
                End If
            End If
            Me.Caption = FormPatient.NameWithTitle & " (" & FormPatient.ID & ") - Visit on " & txtVisitDate.Text & " at " & txtVisitTime.Text
        End If
        .Close
    End With
End Sub

Private Sub cmbPrinter_Change()
    cmbPrinter_Click
End Sub

Private Sub cmbPrinter_Click()
    cmbPaper.Clear
    CsetPrinter.SetPrinterAsDefault (PrescreptionPaperName)
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
    Call SaveVisitDetails
    Call SaveSettings
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, chkClose.Name, chkClose.Value
    SaveSetting App.EXEName, Me.Name, cmbDrawer.Name, Val(cmbDrawer.BoundText)
    SaveSetting App.EXEName, Me.Name, "Printer", PrescreptionPrinterName
    SaveSetting App.EXEName, Me.Name, "Paper", PrescreptionPaperName
    SaveSetting App.EXEName, Me.Name, "ScaleHeight", Val(txtScaleHeight.Text)
    SaveSetting App.EXEName, Me.Name, "ScaleWidth", Val(txtScaleWidth.Text)
    SaveCommonSettings Me
    
End Sub

Private Sub gridDiagnosis_Click()
    Dim temGridRow As Long
    With gridDiagnosis
        temGridRow = .row
        .col = 0
        txtDxDeleteID.Text = .Text
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub gridDiagnosis_DblClick()
    Dim temGridRow As Long
    With gridDiagnosis
        temGridRow = .row
        .col = 0
        txtDxEditID.Text = .Text
        .ColSel = .Cols - 1
        .col = 3
        cmbDiagnosis.BoundText = Val(.Text)
        .col = 2
        txtDiagnosis.Text = .Text
    End With
End Sub

Private Sub btnDxAdd_Click()
    If IsNumeric(cmbDiagnosis.BoundText) = False Then
        MsgBox "Please select a diagnosis"
        cmbDiagnosis.SetFocus
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtDxEditID.Text) <> 0 Then
            temSql = "Select * from tblVisitDiagnosis where VisitDiagnosisID = " & Val(txtDxEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !DiagnosisID = Val(cmbDiagnosis.BoundText)
                !Comments = txtDiagnosis.Text
                .Update
            Else
                .AddNew
                !VisitID = FormVisit.VisitID
                !PatientID = FormPatient.ID
                !DiagnosisID = Val(cmbDiagnosis.BoundText)
                !Comments = txtDiagnosis.Text
                .Update
            End If
        Else
            temSql = "Select * from tblVisitDiagnosis where VisitDiagnosisID = 0"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !VisitID = FormVisit.VisitID
            !PatientID = FormPatient.ID
            !DiagnosisID = Val(cmbDiagnosis.BoundText)
            !Comments = txtDiagnosis.Text
            .Update
        End If
        .Close
    End With
    Call FillDiagnosis
    cmbDiagnosis.Text = Empty
    txtDiagnosis.Text = Empty
    txtDxEditID.Text = Empty
    cmbDiagnosis.SetFocus
    On Error Resume Next
    On Error Resume Next: SendKeys "{Escape}"
End Sub

Private Sub btnDxDelete_Click()
    If IsNumeric(txtDxDeleteID.Text) = False Then
        MsgBox "Please select a diagnosis to delete"
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtDxDeleteID.Text) <> 0 Then
            temSql = "Select * from tblVisitDiagnosis where VisitDiagnosisID = " & Val(txtDxDeleteID.Text)
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedTime = Now
                !DeletedUserID = UserID
                .Update
            End If
        End If
        .Close
    End With
    Call FillDiagnosis
    cmbDiagnosis.Text = Empty
    txtDiagnosis.Text = Empty
    txtDxEditID.Text = Empty
    cmbDiagnosis.SetFocus
    On Error Resume Next: SendKeys "{Escape}"
End Sub











Private Sub gridIx_Click()
    Dim temGridRow As Long
    With gridIx
        temGridRow = .row
        .col = 3
        txtIxID.Text = .Text
        If IsNumeric(.Text) = False Then
            MsgBox "Please select an Investigation"
            gridIx.SetFocus
            Exit Sub
        End If
        .col = 0
        txtIxDeleteID.Text = .Text
        If IsNumeric(.Text) = False Then
            MsgBox "Please select an Investigation"
            gridIx.SetFocus
            Exit Sub
        End If
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub gridIx_DblClick()
    Dim temGridRow As Long
    Dim temPatientIXID As Long
    With gridIx
        temGridRow = .row
        .col = 0
        txtIxEditID.Text = .Text
        .ColSel = .Cols - 1
        .col = 3
        cmbIx.BoundText = Val(.Text)
    End With
End Sub

Private Sub btnIxAdd_Click()
    If IsNumeric(cmbIx.BoundText) = False Then
        MsgBox "Please select a Ix"
        cmbIx.SetFocus
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtIxEditID.Text) <> 0 Then
            temSql = "Select * from tblPatientIx where PatientIxID = " & Val(txtIxEditID.Text)
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !IxID = Val(cmbIx.BoundText)
                .Update
            Else
                .AddNew
                !VisitID = FormVisit.VisitID
                !PatientID = FormPatient.ID
                !IxID = Val(cmbIx.BoundText)
                !Requested = True
                !RequestedDate = Date
                !RequestedTime = Time
                !RequestedDateTime = Now
                !RequestedUserID = UserID
                .Update
            End If
        Else
            If .State = 1 Then .Close
            temSql = "Select * from tblPatientIx where PatientIxID = 0"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !VisitID = FormVisit.VisitID
            !PatientID = FormPatient.ID
            !IxID = Val(cmbIx.BoundText)
            !Requested = True
            !RequestedDate = Date
            !RequestedTime = Time
            !RequestedDateTime = Now
            !RequestedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FillIx
    cmbIx.Text = Empty
    txtIxEditID.Text = Empty
    txtIxID.Text = Empty
    cmbIx.SetFocus
    On Error Resume Next
    On Error Resume Next: SendKeys "{Escape}"
End Sub

Private Sub btnIxDelete_Click()
    If IsNumeric(txtIxDeleteID.Text) = False Then
        MsgBox "Please select a Ix to delete"
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtIxDeleteID.Text) <> 0 Then
            temSql = "Select * from tblPatientIx where PatientIxID = " & Val(txtIxDeleteID.Text)
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedTime = Now
                !DeletedUserID = UserID
                .Update
            End If
        End If
        .Close
    End With
    Call FillIx
    cmbIx.Text = Empty
    txtIxEditID.Text = Empty
    txtIxID.Text = Empty
    txtIxDeleteID.Text = Empty
    cmbIx.SetFocus
    On Error Resume Next: SendKeys "{Escape}"
End Sub


Private Sub FillIx()
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientIx.PatientIxID, tblIx.Ix, tblIx.IxID FROM tblIx RIGHT JOIN tblPatientIx ON tblIx.IxID = tblPatientIx.IxID where tblPatientIx.Deleted = False AND tblPatientIx.VisitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridIx.Clear
        gridIx.Rows = 1
        gridIx.Cols = 4
        gridIx.ColWidth(0) = 0
        gridIx.ColWidth(1) = gridIx.Width - 175
        gridIx.ColWidth(2) = 0
        gridIx.ColWidth(3) = 0
        gridIx.row = 0
        gridIx.col = 1
        gridIx.Text = "Ix"
        
        While .EOF = False
            gridIx.Rows = gridIx.Rows + 1
            gridIx.row = gridIx.Rows - 1
            gridIx.col = 0
            gridIx.Text = !PatientIxID
            gridIx.col = 1
            gridIx.Text = !Ix
            gridIx.col = 3
            gridIx.Text = !IxID
            .MoveNext
        Wend
    End With
    Set rsDx = Nothing
End Sub





Private Sub gridPrescreption_Click()
    
    Dim temRow As Integer
    With gridPrescreption
        temRow = .row
        txtDeletedID.Text = .TextMatrix(temRow, 0)
    End With
End Sub

Private Sub gridPrescreption_DblClick()
    
    Call PharmacyPrescreptionShift(True)
    
    Dim temRow As Integer
    Dim EditID As Long
    
    optPrescreption.Value = True
    
    With gridPrescreption
        temRow = .row
        EditID = Val(.TextMatrix(temRow, 0))
    End With
    Call ClearItemValues
    txtAddID.Text = EditID
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitItem where VIsitItemID = " & EditID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbItem.BoundText = !ItemID
            cmbCategory.BoundText = !ItemCategoryID
            cmbDoseUnit.BoundText = !DoseUnitID
            cmbIssueUnit.BoundText = !IssueUnitID
            cmbFrequencyUnit.BoundText = !FrequencyUnitID
            cmbDurationUnit.BoundText = !DurationUnitID
            txtDose.Text = Val(Format(!Dose, "0.00"))
            txtDuration.Text = Val(Format(!Duration, "0.00"))
            txtIssueQty.Text = !IssueQuentity
            If !Indoor = True Then
                chkIndoor.Value = 1
            Else
                chkIndoor.Value = 0
            End If
            Dim MyPItem As New clsItem
            MyPItem.ID = !PItemID
            cmbPItem.BoundText = !PItemID
            cmbPIssueUnit.BoundText = MyPItem.IssueUnitID
            txtPIssueQty.Text = !PIssueQuentity
        End If
        .Close
    End With
End Sub

Private Sub CalculateTotalFee()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub


'Private Sub SSTab1_Click(PreviousTab As Integer)
'    If SSTab1.Tab = 0 Then
'        Call ClearItemValues
'        Call PharmacyPrescreptionShift(True)
'    Else
'        Call ClearItemValues
'        Call PharmacyPrescreptionShift(False)
'    End If
'End Sub

Private Sub txtConsultationFee_LostFocus()
    If Val(txtConsultationFee.Text) <> FormVisit.ConsultationFee Then
        Dim rsVisit As New ADODB.Recordset
        With rsVisit
            If .State = 1 Then .Close
            temSql = "Select * from tblVisit where visitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !ConsultationFee = Val(txtConsultationFee.Text)
                !TotalFee = !ProcedureFee + !ConsultationFee + !MedicineFee
                !TotalCost = !ProcedureCost + !MedicineCost + !ConsultationCost
                .Update
            End If
            .Close
        End With
    End If
    Call DisplayVisitCharges
End Sub

Private Sub txtDBP_Change()
    On Error Resume Next
    If Len(Trim(txtDBP.Text)) >= 3 Then btnAddBP.SetFocus
    If Len(Trim(txtDBP.Text)) = 2 And Val(txtDBP.Text) > 20 Then btnAddBP.SetFocus
End Sub

Private Sub txtDose_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTem As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        If cmbDoseUnit.Enabled = True Then
            KeyCode = Empty
            cmbDoseUnit.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtDose.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        With rsTem
            If .State = 1 Then .Close
            temSql = "SELECT tblVisitItem.Dose "
            temSql = temSql & "From tblVisitItem "
            If IsNumeric(cmbDoseUnit.BoundText) = True Then
                temSql = temSql & "Where (((tblVisitItem.Dose) > " & Val(txtDose.Text) & ") And ((tblVisitItem.ItemID) = " & Val(cmbItem.BoundText) & ") And ((tblVisitItem.DoseUnitID) = " & Val(cmbDoseUnit.BoundText) & ") And ((tblVisitItem.Deleted) = False)) "
            Else
                temSql = temSql & "Where (((tblVisitItem.Dose) > " & Val(txtDose.Text) & ") And ((tblVisitItem.ItemID) = " & Val(cmbItem.BoundText) & ") And ((tblVisitItem.Deleted) = False)) "
            End If
            temSql = temSql & "ORDER BY tblVisitItem.Dose"
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If IsNull(!Dose) = False Then
                    txtDose.Text = !Dose
                End If
            End If
            .Close
        End With
    ElseIf KeyCode = vbKeyDown Then
        With rsTem
            If .State = 1 Then .Close
            temSql = "SELECT tblVisitItem.Dose "
            temSql = temSql & "From tblVisitItem "
            If IsNumeric(cmbDoseUnit.BoundText) = True Then
                temSql = temSql & "Where (((tblVisitItem.Dose) < " & Val(txtDose.Text) & ") And ((tblVisitItem.ItemID) = " & Val(cmbItem.BoundText) & ") And ((tblVisitItem.DoseUnitID) = " & Val(cmbDoseUnit.BoundText) & ") And ((tblVisitItem.Deleted) = False)) "
            Else
                temSql = temSql & "Where (((tblVisitItem.Dose) < " & Val(txtDose.Text) & ") And ((tblVisitItem.ItemID) = " & Val(cmbItem.BoundText) & ") And ((tblVisitItem.Deleted) = False)) "
            End If
            temSql = temSql & "ORDER BY tblVisitItem.Dose DESC"
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If IsNull(!Dose) = False Then
                    txtDose.Text = !Dose
                End If
            End If
            .Close
        End With
    ElseIf KeyCode > vbKeyA And KeyCode < vbKeyZ Then
        Dim TemKeyCode As Integer
        TemKeyCode = KeyCode
        KeyCode = Empty
        cmbDoseUnit.SetFocus
        SendKeys KeyCodeToString(KeyCode, Shift)
    End If
End Sub

Private Sub txtDuration_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbDurationUnit.Enabled = False Then
            cmbDurationUnit.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtDuration.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtDuration.Text = Val(txtDuration.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtDuration.Text) > 0 Then
            txtDuration.Text = Val(txtDuration.Text) - 1
        End If
    End If
End Sub

Private Sub txtHt_Change()
    txtBmi.Text = calculateBmi(Val(txtHt.Text), Val(txtWeight.Text))
End Sub

Private Sub txtIssueQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbIssueUnit.Enabled = True Then
            cmbIssueUnit.SetFocus
        Else
            btnRxAdd_Click
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtIssueQty.Text = Empty
    ElseIf KeyCode = vbKeyUp Then
        txtIssueQty.Text = Val(txtIssueQty.Text) + 1
    ElseIf KeyCode = vbKeyDown Then
        If Val(txtIssueQty.Text) > 0 Then
            txtIssueQty.Text = Val(txtIssueQty.Text) - 1
        End If
    End If
    
End Sub

Private Sub txtMedicinesFee_Change()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub

Private Sub txtProcedureFee_Change()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub

Private Sub txtSBP_Change()
    On Error Resume Next
    If Len(txtSBP.Text) >= 3 Then txtDBP.SetFocus
    If Len(txtSBP.Text) = 2 And Val(txtSBP.Text) > 30 Then txtDBP.SetFocus
End Sub

Private Sub txtTotalCharge_Change()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub

Private Sub btnDelete_Click()
    
    Dim i As Integer
    Dim rsTem As New ADODB.Recordset
    
    i = MsgBox("Are you sure you want to delete the visit?", vbYesNo)
    If i = vbNo Then Exit Sub
    i = MsgBox("When you delete a visit, the Referrals, Certificates, Investigations all will also be deleted. The Medicines will NOT be added back to stocks. Are you  still sure you want to delete the visit?", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim MyTable(8) As String
    
    chkDelete.Value = 1
    
    MyTable(0) = "Visit"
    MyTable(1) = "VisitIx"
    MyTable(2) = "VisitCertificate"
    MyTable(3) = "VisitDiagnosis"
    MyTable(4) = "VisitItem"
    MyTable(5) = "VisitIx"
    MyTable(6) = "VisitProcedure"
    MyTable(7) = "VisitReferral"
    MyTable(8) = "VisitSystemFinding"
    
    txtTotalCharge.Text = 0
    
    For i = 0 To UBound(MyTable)
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tbl" & MyTable(i) & " where VisitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            While .EOF = False
                !Deleted = True
                !DeletedDate = Date
                !DeletedTime = Time
                !DeletedUserID = UserID
                If MyTable(i) = "Visit" Then
                    ChangeDrawerBalances !DrawerID, !PaymentMethodID, !TotalFee, False, True, False
                    !TotalFee = 0
                    !MedicineFee = 0
                    !ProcedureFee = 0
                    !ConsultationFee = 0
                    !MedicineCost = 0
                    !ProcedureCost = 0
                    !ConsultationCost = 0
                End If
                .Update
                .MoveNext
            Wend
            .Close
        End With
    Next
    Unload Me
End Sub

Private Sub txtWeight_Change()
    txtBmi.Text = calculateBmi(Val(txtHt.Text), Val(txtWeight.Text))
End Sub
