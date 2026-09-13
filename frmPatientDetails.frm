VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPatientDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Details"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   13035
   Begin VB.TextBox txtTem 
      Height          =   360
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   93
      Text            =   "frmPatientDetails.frx":058A
      Top             =   6600
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3840
      TabIndex        =   92
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3495
      Left            =   120
      TabIndex        =   69
      Top             =   4200
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Queue"
      TabPicture(0)   =   "frmPatientDetails.frx":0590
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label18"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label21"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbDoctor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbSecession"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtWeight"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSBP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDBP"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmPatientDetails.frx":05AC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).Control(1)=   "gridPatient"
      Tab(1).Control(2)=   "optMale"
      Tab(1).Control(3)=   "optFemale"
      Tab(1).Control(4)=   "optAll"
      Tab(1).ControlCount=   5
      Begin TabDlg.SSTab SSTab3 
         Height          =   855
         Left            =   -74880
         TabIndex        =   83
         Top             =   360
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   1508
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Name"
         TabPicture(0)   =   "frmPatientDetails.frx":05C8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtName"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Address"
         TabPicture(1)   =   "frmPatientDetails.frx":05E4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtByAddress"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Birthday"
         TabPicture(2)   =   "frmPatientDetails.frx":0600
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dtpSearchDOB"
         Tab(2).ControlCount=   1
         Begin VB.TextBox txtName 
            Height          =   360
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtByAddress 
            Height          =   360
            Left            =   -74880
            TabIndex        =   84
            Top             =   360
            Width           =   3735
         End
         Begin MSComCtl2.DTPicker dtpSearchDOB 
            Height          =   375
            Left            =   -74880
            TabIndex        =   86
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   128581635
            CurrentDate     =   39984
         End
      End
      Begin VB.TextBox txtDBP 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   72
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtSBP 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   71
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtWeight 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   70
         Top             =   1920
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   960
         TabIndex        =   73
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy"
         Format          =   128581635
         CurrentDate     =   39859
      End
      Begin MSDataListLib.DataCombo cmbSecession 
         Height          =   315
         Left            =   960
         TabIndex        =   74
         Top             =   1440
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbDoctor 
         Height          =   315
         Left            =   960
         TabIndex        =   75
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridPatient 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   82
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3413
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton optMale 
         Caption         =   "&Male"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   89
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optFemale 
         Caption         =   "&Female"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73320
         TabIndex        =   90
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optAll 
         Caption         =   "&All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71640
         TabIndex        =   91
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "&Weight"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "&BP                                    mmHg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   79
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Da&te"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Docto&r"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "&Secession"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.TextBox txtSearchID 
      Height          =   360
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   4560
      TabIndex        =   8
      Top             =   4800
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Past Visits"
      TabPicture(0)   =   "frmPatientDetails.frx":061C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtVisit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstVisit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&History"
      TabPicture(1)   =   "frmPatientDetails.frx":0638
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label26"
      Tab(1).Control(1)=   "Label27"
      Tab(1).Control(2)=   "lstDrugs"
      Tab(1).Control(3)=   "lstDx"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "A&llergies"
      TabPicture(2)   =   "frmPatientDetails.frx":0654
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label28"
      Tab(2).Control(1)=   "Label29"
      Tab(2).Control(2)=   "lstFoodAllergies"
      Tab(2).Control(3)=   "lstAllergies"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Investigations"
      TabPicture(3)   =   "frmPatientDetails.frx":0670
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label30"
      Tab(3).Control(1)=   "Label31"
      Tab(3).Control(2)=   "lstIx"
      Tab(3).Control(3)=   "txtIx"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "I&mages"
      TabPicture(4)   =   "frmPatientDetails.frx":068C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtImage"
      Tab(4).Control(1)=   "lstImages"
      Tab(4).Control(2)=   "Label33"
      Tab(4).Control(3)=   "Label32"
      Tab(4).Control(4)=   "imgPatient"
      Tab(4).ControlCount=   5
      Begin VB.TextBox txtImage 
         Height          =   2055
         Left            =   -72600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtIx 
         Height          =   1980
         Left            =   -72000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   735
         Width           =   5175
      End
      Begin MSDataListLib.DataList lstVisit 
         Height          =   1980
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3493
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin VB.TextBox txtVisit 
         Height          =   2025
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   720
         Width           =   4815
      End
      Begin MSDataListLib.DataList lstDx 
         Height          =   1980
         Left            =   -74880
         TabIndex        =   14
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3493
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstDrugs 
         Height          =   1980
         Left            =   -70920
         TabIndex        =   16
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3493
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstAllergies 
         Height          =   2100
         Left            =   -74880
         TabIndex        =   18
         Top             =   660
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3704
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstIx 
         Height          =   1980
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3493
         _Version        =   393216
      End
      Begin MSDataListLib.DataList lstImages 
         Height          =   2025
         Left            =   -74880
         TabIndex        =   26
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3572
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstFoodAllergies 
         Height          =   2100
         Left            =   -70920
         TabIndex        =   20
         Top             =   660
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3704
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Image &Details"
         Height          =   255
         Left            =   -72600
         TabIndex        =   27
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "I&mages"
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "&Results"
         Height          =   255
         Left            =   -70800
         TabIndex        =   23
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "&Investigation"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "&Food Allergies"
         Height          =   255
         Left            =   -69840
         TabIndex        =   19
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "&Drug Allergies"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Drug &History"
         Height          =   255
         Left            =   -70560
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "&Diagnoses"
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Visit &Details"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "&Visit"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image imgPatient 
         Height          =   2475
         Left            =   -69840
         Top             =   360
         Width           =   3000
      End
   End
   Begin VB.TextBox txtCode 
      Height          =   360
      Left            =   8400
      TabIndex        =   51
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtID 
      Height          =   360
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtMobile 
      Height          =   360
      Left            =   8520
      TabIndex        =   62
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txteMail 
      Height          =   360
      Left            =   10920
      TabIndex        =   64
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtPhone 
      Height          =   360
      Left            =   6000
      TabIndex        =   60
      Top             =   3840
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpDOB 
      Height          =   375
      Left            =   6000
      TabIndex        =   43
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   90701827
      CurrentDate     =   39845
   End
   Begin VB.TextBox txtAge 
      Height          =   360
      Left            =   8400
      TabIndex        =   45
      Top             =   1920
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo cmbTitle 
      Height          =   360
      Left            =   6000
      TabIndex        =   38
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtAddress 
      Height          =   345
      Left            =   6000
      TabIndex        =   55
      Top             =   2880
      Width           =   6735
   End
   Begin VB.TextBox txtComments 
      Height          =   360
      Left            =   6000
      TabIndex        =   57
      Top             =   3360
      Width           =   6735
   End
   Begin VB.TextBox txtFullName 
      Height          =   360
      Left            =   6000
      TabIndex        =   36
      Top             =   960
      Width           =   6735
   End
   Begin VB.TextBox txtPatient 
      Height          =   360
      Left            =   6000
      TabIndex        =   34
      Top             =   480
      Width           =   6735
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin MSDataListLib.DataCombo cmbCivilStatus 
      Height          =   360
      Left            =   10920
      TabIndex        =   41
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSex 
      Height          =   360
      Left            =   8400
      TabIndex        =   39
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbRace 
      Height          =   360
      Left            =   10920
      TabIndex        =   47
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   11640
      TabIndex        =   32
      Top             =   7800
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   11640
      TabIndex        =   68
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   4560
      TabIndex        =   65
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin MSDataListLib.DataCombo cmbBloodGroup 
      Height          =   360
      Left            =   6000
      TabIndex        =   49
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnSaveNewVisit 
      Height          =   375
      Left            =   6120
      TabIndex        =   66
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save And &New Visit"
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
   Begin btButtonEx.ButtonEx btnSaveQueue 
      Height          =   375
      Left            =   8880
      TabIndex        =   67
      Top             =   4320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save And &Queue"
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
   Begin btButtonEx.ButtonEx btnIx 
      Height          =   375
      Left            =   6120
      TabIndex        =   30
      Top             =   7800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Investigations"
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
   Begin btButtonEx.ButtonEx btnPastHistory 
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   7800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&History"
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
   Begin btButtonEx.ButtonEx btnImage 
      Height          =   375
      Left            =   7920
      TabIndex        =   31
      Top             =   7800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "I&mages"
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
   Begin btButtonEx.ButtonEx btnSearch 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Sea&rch"
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
   Begin MSDataListLib.DataCombo cmbPatient 
      Height          =   3060
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5398
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnToQueue 
      Height          =   375
      Left            =   2280
      TabIndex        =   87
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Add To &Queue"
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
   Begin btButtonEx.ButtonEx btnNewVisit 
      Height          =   375
      Left            =   120
      TabIndex        =   88
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "New &Visit"
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
   Begin btButtonEx.ButtonEx btnFamilyHistory 
      Height          =   375
      Left            =   9480
      TabIndex        =   94
      Top             =   7800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Family History"
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
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "&ID"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "&Blood Group"
      Height          =   255
      Left            =   4680
      TabIndex        =   48
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   255
      Left            =   10080
      TabIndex        =   52
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "C&ode"
      Height          =   255
      Left            =   7920
      TabIndex        =   50
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "&Mobile"
      Height          =   255
      Left            =   7800
      TabIndex        =   61
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "&eMail"
      Height          =   255
      Left            =   10080
      TabIndex        =   63
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Phone"
      Height          =   255
      Left            =   4680
      TabIndex        =   59
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "A&ddress"
      Height          =   255
      Left            =   4680
      TabIndex        =   54
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Age"
      Height          =   255
      Left            =   7920
      TabIndex        =   44
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Date of Birth"
      Height          =   255
      Left            =   4680
      TabIndex        =   42
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Race"
      Height          =   255
      Left            =   10080
      TabIndex        =   46
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ci&vil St."
      Height          =   255
      Left            =   10080
      TabIndex        =   40
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Se&x"
      Height          =   255
      Left            =   7920
      TabIndex        =   58
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Title"
      Height          =   255
      Left            =   4680
      TabIndex        =   37
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   4680
      TabIndex        =   56
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Full Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   35
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Na&me"
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Patient"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatientDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim SearchCount As Long
    Dim rsPastVisits As New ADODB.Recordset
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbPatient.BoundText) = False Then
        temText = cmbPatient.Text
    Else
        temText = Empty
    End If
    cmbPatient.Text = Empty
    Call EditMode
    txtPatient.Text = temText
    txtPatient.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbPatient.Text = Empty
    cmbPatient.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbPatient.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
            MsgBox "Deleted"
        Else
            MsgBox "Nothing to Delete"
        End If
        .Close
    End With
    Set rsTem = Nothing
    Call FillCombos
    cmbPatient.SetFocus
    cmbPatient.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbPatient.BoundText) = False Then Exit Sub
    Call EditMode
    txtPatient.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnFamilyHistory_Click()
    frmFamilyTree.Show
    frmFamilyTree.ZOrder 0
End Sub

Private Sub btnImage_Click()
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
    
    frmImages.Show
    frmImages.ZOrder 0
End Sub

Private Sub btnIx_Click()
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
    
    frmPatientIx.Show
    frmPatientIx.ZOrder 0
End Sub

Private Sub btnNewVisit_Click()
    On Error Resume Next
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbDoctor.BoundText) = False Then
        MsgBox "Please select a doctor"
        cmbDoctor.SetFocus
        Exit Sub
    End If
    currentDoctorId = Val(cmbDoctor.BoundText)
    Unload frmNewVisit
    Dim NewVisitID As Long
    NewVisitID = CurrentVisit.NewVisitID(Val(cmbPatient.BoundText), Date, UserID)
    frmNewVisit.Show
    frmNewVisit.ZOrder 0
End Sub

Private Sub btnPastHistory_Click()
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient first"
        cmbPatient.SetFocus
        Exit Sub
    End If
    CurrentPatient.ID = cmbPatient.BoundText
    Unload frmPastHistory
    frmPastHistory.Show
    frmPastHistory.ZOrder 0
End Sub

Private Sub btnSave_Click()
    Dim temPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        temPatientID = cmbPatient.BoundText
        Call SaveOld
    Else
        temPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.SetFocus
    cmbPatient.BoundText = temPatientID
End Sub

Private Sub btnSaveNewVisit_Click()
    Dim temPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        temPatientID = cmbPatient.BoundText
        Call SaveOld
    Else
        temPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.BoundText = temPatientID
    
    Unload frmNewVisit
    Dim NewVisitID As Long
    NewVisitID = CurrentVisit.NewVisitID(Val(cmbPatient.BoundText), Date, UserID)
    frmNewVisit.Show
    frmNewVisit.ZOrder 0
    
End Sub

Private Sub btnSaveQueue_Click()
    Dim temPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        Call SaveOld
        temPatientID = cmbPatient.BoundText
    Else
        temPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.BoundText = temPatientID
    Dim rsQueue As New ADODB.Recordset
    If IsNumeric(cmbDoctor.BoundText) = False Then
        MsgBox "Please select a doctor"
        cmbDoctor.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSecession.BoundText) = False Then
        MsgBox "Please select a secession"
        cmbSecession.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    i = MsgBox("Patient" & vbTab & ":" & vbTab & CurrentPatient.NameWithTitle & vbNewLine & "Age   " & vbTab & ":" & vbTab & CurrentPatient.AgeInWords & vbNewLine & "Date    " & vbTab & ":" & vbTab & Format(dtpDate.Value, "dd MMMM yyyy") & vbNewLine & "Secession" & vbTab & ":" & vbTab & cmbSecession.Text & vbNewLine & "Number" & vbTab & ":" & vbTab & GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText)) & vbNewLine & vbNewLine & "Add to Queue?", vbYesNo)
    If i = vbNo Then Exit Sub
    With rsQueue
        temSql = "Select * from tblQueue"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PatientID = Val(cmbPatient.BoundText)
        !SecessionID = Val(cmbSecession.BoundText)
        !QueueDate = dtpDate.Value
        !DoctorID = Val(cmbDoctor.BoundText)
        !QueueNo = GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText))
        !BookedUserID = UserID
        !BookedAt = Now
        !weight = Val(txtWeight.Text)
        !SBP = Val(txtSBP.Text)
        !DBP = Val(txtDBP.Text)
        !Age = CurrentPatient.AgeInWords
        .Update
'        MsgBox "Added to Queue. Number is " & !QueueNo
        .Close
    End With
    txtWeight.Text = Empty
    txtSBP.Text = Empty
    txtDBP.Text = Empty
End Sub

Private Sub btnSearch_Click()
    frmSearchPatient.Show
    frmSearchPatient.FormName = Me.Name
    frmSearchPatient.ZOrder 0
End Sub

Private Sub btnToQueue_Click()
    Dim rsQueue As New ADODB.Recordset
    If IsNumeric(cmbDoctor.BoundText) = False Then
        MsgBox "Please select a doctor"
        cmbDoctor.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSecession.BoundText) = False Then
        MsgBox "Please select a secession"
        cmbSecession.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
    Dim i As Integer
'    i = MsgBox("Patient" & vbTab & ":" & vbTab & CurrentPatient.NameWithTitle & vbNewLine & "Age   " & vbTab & ":" & vbTab & CurrentPatient.AgeInWords & vbNewLine & "Date    " & vbTab & ":" & vbTab & Format(dtpDate.Value, "dd MMMM yyyy") & vbNewLine & "Secession" & vbTab & ":" & vbTab & cmbSecession.Text & vbNewLine & "Number" & vbTab & ":" & vbTab & GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText)) & vbNewLine & vbNewLine & "Add to Queue?", vbYesNo)
    i = MsgBox("Patient" & vbTab & ":" & vbTab & CurrentPatient.NameWithTitle & vbNewLine & "Age   " & vbTab & ":" & vbTab & CurrentPatient.AgeInWords & vbNewLine & vbNewLine & "Doctor  " & vbTab & ":" & vbTab & cmbDoctor.Text & vbNewLine & "Date    " & vbTab & ":" & vbTab & Format(dtpDate.Value, "dd MMMM yyyy") & vbNewLine & "Secession" & vbTab & ":" & vbTab & cmbSecession.Text & vbNewLine & "Number" & vbTab & ":" & vbTab & GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText)) & vbNewLine & vbNewLine & "Add to Queue?", vbYesNo)
    
    If i = vbNo Then Exit Sub
    
    With rsQueue
        temSql = "Select * from tblQueue"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PatientID = Val(cmbPatient.BoundText)
        !SecessionID = Val(cmbSecession.BoundText)
        !QueueDate = dtpDate.Value
        !DoctorID = Val(cmbDoctor.BoundText)
        !QueueNo = GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText))
        !BookedUserID = UserID
        !BookedAt = Now
        !weight = Val(txtWeight.Text)
        !SBP = Val(txtSBP.Text)
        !DBP = Val(txtDBP.Text)
        !Age = CurrentPatient.AgeInWords
        .Update
'        MsgBox "Added to Queue. Number is " & !QueueNo
        .Close
    End With
    txtWeight.Text = Empty
    txtSBP.Text = Empty
    txtDBP.Text = Empty
End Sub


Private Sub cmbPatient_Change()
    
    On Error Resume Next
    Call ClearValues
    CurrentPatient.ID = Val(cmbPatient.BoundText)
    'If IsNumeric(cmbPatient.BoundText) = True Then
        Call DisplayDetails
        txtAge.Text = CurrentPatient.AgeInYears
        Call ListVisits
        Call ListHistory
        Call ListIx
        Call ListAllergies
        Call ListImages
    'End If
'    txtName.Text = cmbPatient.Text
    
End Sub

Private Sub ListHistory()
    Dim rsDx As New ADODB.Recordset
    Dim rsIt As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientDiagnosis.PatientDiagnosisID, tblPatientDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblPatientDiagnosis ON tblDiagnosis.DiagnosisID = tblPatientDiagnosis.DiagnosisID where tblPatientDiagnosis.Deleted = False AND tblPatientDiagnosis.PatientID = " & CurrentPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstDx
        Set .RowSource = rsDx
        .ListField = "Diagnosis"
        .BoundColumn = "PatientDiagnosisID"
    End With
    With rsIt
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientItem.PatientItemID, [tblItem.Item] & ' ' & [tblPatientItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblPatientItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] & ' (' & [tblPatientItem].[IndoorOutdoor] & ')' AS Prescreption, 'Issue ' & tblPatientItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM (((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblPatientItem ON tblIssueUnit.ItemUnitID = tblPatientItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblPatientItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblPatientItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblPatientItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblPatientItem.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblPatientItem.PatientID)=" & CurrentPatient.ID & ") AND ((tblPatientItem.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstDrugs
        Set .RowSource = rsIt
        .ListField = "Prescreption"
        .BoundColumn = "PatientItemID"
    End With
End Sub

Private Sub ListIx()
    Dim temID As Long
    Dim rsPtIx As New ADODB.Recordset
    With rsPtIx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientIX.PatientIXID, tblIx.Ix, tblIx.IxID FROM tblIx RIGHT JOIN tblPatientIX ON tblIx.IxID = tblPatientIX.IxID where tblPatientIX.Deleted = False AND tblPatientIX.PatientID = " & CurrentPatient.ID & " Order by PatientIXID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temID = !PatientIxID
        End If
    End With
    With lstIx
        Set .RowSource = rsPtIx
        .ListField = "Ix"
        .BoundColumn = "PatientIXID"
        .BoundText = temID
        If temID <> 0 Then
            lstIx_Click
        End If
    End With
End Sub

Private Sub ListAllergies()
    Dim rsDA As New ADODB.Recordset
    Dim rsFA As New ADODB.Recordset
    
    With rsDA
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientAllergy.PatientAllergyID, tblPatientAllergy.Comments, tblItem.Item, tblItem.ItemID FROM tblItem RIGHT JOIN tblPatientAllergy ON tblItem.ItemID = tblPatientAllergy.ItemID where tblPatientAllergy.ItemID <> 0 AND tblPatientAllergy.Deleted = False AND tblPatientAllergy.PatientID = " & CurrentPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstAllergies
        Set .RowSource = rsDA
        .ListField = "Item"
        .BoundColumn = "PatientAllergyID"
    End With
    
    With rsFA
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientAllergy.PatientAllergyID, tblPatientAllergy.Comments, tblFoodItem.FoodItem, tblFoodItem.FoodItemID FROM tblFoodItem RIGHT JOIN tblPatientAllergy ON tblFoodItem.FoodItemID = tblPatientAllergy.FoodItemID where tblPatientAllergy.FoodItemID <> 0 AND tblPatientAllergy.Deleted = False AND tblPatientAllergy.PatientID = " & CurrentPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstFoodAllergies
        Set .RowSource = rsFA
        .ListField = "FoodItem"
        .BoundColumn = "PatientAllergyID"
    End With
End Sub

Private Sub ListImages()
    Dim rsImg As New ADODB.Recordset
    Dim temID As Long
    
    With rsImg
        temSql = "Select * from tblImage where PatientID <> 0 AND PatientID = " & CurrentPatient.ID & " AND Deleted = False  Order by PatientDefault, ImageID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temID = !ImageID
            imgPatient.Stretch = True
            FillImage rsImg, "ImageData", "ImageSize", imgPatient
            txtImage.Text = !Image & vbNewLine & !Comments
        Else
            temID = 0
            imgPatient.Picture = LoadPicture()
        End If
    End With
    With lstImages
        Set .RowSource = rsImg
        .ListField = "Image"
        .BoundColumn = "ImageID"
        .BoundText = temID
    End With
End Sub

Private Sub ListVisits()
    Dim temID As Long
    With rsPastVisits
        If .State = 1 Then .Close
        temSql = "Select tblVisit.VisitID, format$(tblVisit.VisitDate, 'dd/MMM/yyyy') & ' - ' & SBP & '/' & DBP & ' - ' & VisitWeight & ' kg' as VisitDisplay from tblVisit where PatientID = " & Val(cmbPatient.BoundText) & " And Deleted = FALSE Order by VisitID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temID = !VisitID
        Else
            temID = 0
        End If
    End With
    With lstVisit
        Set .RowSource = rsPastVisits
        .ListField = "VisitDisplay"
        .BoundColumn = "VisitID"
        .BoundText = temID
        Call lstVisit_Click
    End With
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


Private Sub cmbPatient_Click(Area As Integer)
    Call cmbPatient_Change
End Sub


Private Sub cmbPatient_KeyUp(KeyCode As Integer, Shift As Integer)
    txtName.Text = cmbPatient.Text

End Sub

Private Sub Command1_Click()
    Dim MyCtrl As Control
    txtTem.Text = Empty
    For Each MyCtrl In Controls
'        If TypeOf MyCtrl Is Textbox Then
'            txtTem.Text = txtTem.Text & vbNewLine
'            txtTem.Text = txtTem.Text & "Private Sub " & MyCtrl.Name & "_GotFocus()" & vbNewLine
'            txtTem.Text = txtTem.Text & "SendKeys '{home}+{end}'" & vbNewLine
'            txtTem.Text = txtTem.Text & "End Sub" & vbNewLine
'        End If
        If TypeOf MyCtrl Is Textbox Then
            txtTem.Text = txtTem.Text & vbNewLine
            txtTem.Text = txtTem.Text & "Private Sub " & MyCtrl.Name & "_KeyDown(KeyCode As Integer, Shift As Integer)" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & "If KeyCode = vbKeyReturn Then" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "KeyCode = Empty" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "If Shift = 1 Then" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & vbTab & ".SETFOCUS" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "Else" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & vbTab & ".setfocus" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "End if" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & "ElseIf KeyCode = vbKeyEscape Then" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & MyCtrl.Name & ".text = empty" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & "End if" & vbNewLine
            txtTem.Text = txtTem.Text & "End Sub" & vbNewLine
        End If
        If TypeOf MyCtrl Is DataCombo Then
            txtTem.Text = txtTem.Text & vbNewLine
            txtTem.Text = txtTem.Text & "Private Sub " & MyCtrl.Name & "_KeyDown(KeyCode As Integer, Shift As Integer)" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & "If KeyCode = vbKeyReturn Then" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "KeyCode = Empty" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "If Shift = 1 Then" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & vbTab & ".SETFOCUS" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "Else" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & vbTab & ".setfocus" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & "End if" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & "ElseIf KeyCode = vbKeyEscape Then" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & vbTab & MyCtrl.Name & ".text = empty" & vbNewLine
            txtTem.Text = txtTem.Text & vbTab & "End if" & vbNewLine
            txtTem.Text = txtTem.Text & "End Sub" & vbNewLine
        End If
        
    Next
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtComments.SetFocus
        Else
            txtComments.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtComments.Text = Empty
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbDoctor.SetFocus
        Else
            cmbSecession.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        dtpDate.Value = Date
    End If
End Sub

Private Sub dtpDOB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbCivilStatus.SetFocus
        Else
            txtAge.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        dtpDOB.Value = Date
    End If
End Sub

Private Sub txtName_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtByAddress_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtDBP_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtPatient_Change()
    txtFullName.Text = txtPatient.Text
End Sub

Private Sub txtSBP_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtWeight_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtSearchID_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtImage_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtIx_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtVisit_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtCode_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtID_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtMobile_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txteMail_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtPhone_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtAge_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtAddress_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtComments_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtFullName_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtPatient_GotFocus()
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub txtSBP_Change()
    On Error Resume Next
    If Len(txtSBP.Text) >= 3 Then txtDBP.SetFocus
    If Len(txtSBP.Text) = 2 And Val(txtSBP.Text) > 30 Then txtDBP.SetFocus
End Sub

Private Sub txtDBP_Change()
    On Error Resume Next
    If Len(Trim(txtDBP.Text)) >= 3 Then btnToQueue.SetFocus
    If Len(Trim(txtDBP.Text)) = 2 And Val(txtDBP.Text) > 20 Then btnToQueue.SetFocus
End Sub


Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            btnCancel.SetFocus
        Else
            txtFullName.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtName.Text = Empty
    End If
End Sub

Private Sub txtByAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtCode.SetFocus
        Else
            txtComments.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtByAddress.Text = Empty
    End If
End Sub

Private Sub txtDBP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtSBP.SetFocus
        Else
            btnToQueue.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtDBP.Text = Empty
    End If
End Sub

Private Sub txtSBP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtWeight.SetFocus
        Else
            txtDBP.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtSBP.Text = Empty
    End If
End Sub

Private Sub txtWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbSecession.SetFocus
        Else
            txtSBP.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtWeight.Text = Empty
    End If
End Sub

Private Sub cmbSecession_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            dtpDate.SetFocus
        Else
            txtWeight.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbSecession.Text = Empty
    End If
End Sub

Private Sub cmbDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbPatient.SetFocus
        Else
            cmbSecession.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbDoctor.Text = Empty
    End If
End Sub

Private Sub txtSearchID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbPatient.SetFocus
        Else
            btnSearch.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtSearchID.Text = Empty
    End If
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbBloodGroup.SetFocus
        Else
            txtAddress.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtCode.Text = Empty
    End If
End Sub

Private Sub txtMobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtPhone.SetFocus
        Else
            txteMail.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtMobile.Text = Empty
    End If
End Sub

Private Sub txteMail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtMobile.SetFocus
        Else
            btnSave.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txteMail.Text = Empty
    End If
End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtComments.SetFocus
        Else
            txtMobile.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtPhone.Text = Empty
    End If
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            dtpDOB.SetFocus
        Else
            cmbRace.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtAge.Text = Empty
    End If
End Sub

Private Sub cmbTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtFullName.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
        Else
            cmbSex.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbTitle.Text = Empty
    End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtCode.SetFocus
        Else
            txtComments.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtAddress.Text = Empty
    End If
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtAddress.SetFocus
        Else
            txtPhone.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtComments.Text = Empty
    End If
End Sub

Private Sub txtFullName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtPatient.SetFocus
        Else
            cmbTitle.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtFullName.Text = Empty
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            btnCancel.SetFocus
        Else
            txtFullName.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtPatient.Text = Empty
    End If
End Sub

Private Sub cmbCivilStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbSex.SetFocus
        Else
            dtpDOB.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbCivilStatus.Text = Empty
    End If
End Sub

Private Sub cmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbTitle.SetFocus
        Else
            cmbCivilStatus.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbSex.Text = Empty
    End If
End Sub

Private Sub cmbRace_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtAge.SetFocus
        Else
            cmbBloodGroup.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbRace.Text = Empty
    End If
End Sub

Private Sub cmbBloodGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbRace.SetFocus
        Else
            txtCode.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbBloodGroup.Text = Empty
    End If
End Sub

Private Sub cmbPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtSearchID.SetFocus
        Else
            btnAdd.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbPatient.Text = Empty
    End If
End Sub





Private Sub dtpSearchDOB_Change()
    Call PatientSearch
    SearchCount = SearchCount + 1

End Sub

Private Sub Form_Load()
    Call SetColours
    Call SelectMode
    Call ClearValues
    Call FillCombos
    Call FillOnceCombo
    Call GetFormSettings
    If AutoSecession = True Then Call SetAutoSecession
End Sub

Private Sub SetAutoSecession()
'     If AutoSecession = True Then Call SetAutoSecession
    
    If Time > #2:00:00 PM# Then
        cmbSecession.Text = "Evening"
    Else
        cmbSecession.Text = "Morning"
    End If
End Sub


Private Sub FillOnceCombo()
    Dim Title As New clsFillCombos
    Title.FillAnyCombo cmbTitle, "Title", True
    Dim Sex As New clsFillCombos
    Sex.FillAnyCombo cmbSex, "Sex", True
    Dim Race As New clsFillCombos
    Race.FillAnyCombo cmbRace, "Race", True
    Dim CivilStatus As New clsFillCombos
    CivilStatus.FillAnyCombo cmbCivilStatus, "CivilStatus", True
    Dim BP As New clsFillCombos
    BP.FillAnyCombo cmbBloodGroup, "BloodGroup", True
    Dim Doc As New clsFillCombos
    Doc.FillBoolCombo cmbDoctor, "staff", "staff", "IsADoctor", True
    Dim Secession As New clsFillCombos
    Secession.FillBoolCombo cmbSecession, "Secession", "Secession", "IsVisitSecession", True
End Sub

Private Sub GetFormSettings()
    cmbDoctor.BoundText = Val(GetSetting(App.EXEName, Me.Name, "DoctorID", 0))
    cmbSecession.BoundText = Val(GetSetting(App.EXEName, Me.Name, "SecessionID", 0))
    GetCommonSettings Me
    
    dtpDate.Value = Date
    SSTab1.Tab = Val(GetSetting(App.EXEName, Me.Name, SSTab1.Name, 0))
    SSTab2.Tab = Val(GetSetting(App.EXEName, Me.Name, SSTab2.Name, 0))
    SSTab3.Tab = Val(GetSetting(App.EXEName, Me.Name, SSTab3.Name, 0))
    optAll.Value = CBool(GetSetting(App.EXEName, Me.Name, optAll.Name, True))
    optMale.Value = CBool(GetSetting(App.EXEName, Me.Name, optMale.Name, True))
    optFemale.Value = CBool(GetSetting(App.EXEName, Me.Name, optFemale.Name, True))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetFormSettings
End Sub

Private Sub SetFormSettings()
    SaveSetting App.EXEName, Me.Name, "DoctorID", Val(cmbDoctor.BoundText)
    SaveSetting App.EXEName, Me.Name, "SecessionID", Val(cmbSecession.BoundText)
    SaveSetting App.EXEName, Me.Name, SSTab1.Name, SSTab1.Tab
    SaveSetting App.EXEName, Me.Name, SSTab2.Name, SSTab2.Tab
    SaveSetting App.EXEName, Me.Name, SSTab3.Name, SSTab3.Tab
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, Me.Name, optAll.Name, optAll.Value
    SaveSetting App.EXEName, Me.Name, optMale.Name, optMale.Value
    SaveSetting App.EXEName, Me.Name, optFemale.Name, optFemale.Value
End Sub

Private Sub gridPatient_DblClick()
    Dim temPatientID As Long
    temPatientID = Val(gridPatient.TextMatrix(gridPatient.row, 0))
    cmbPatient.BoundText = temPatientID
End Sub

Private Sub imgPatient_DblClick()
    Dim MyFrm As Form
    Set MyFrm = New frmPatientImage
    MyFrm.Show
    MyFrm.ZOrder 0
    MyFrm.Caption = lstImages.Text
End Sub

Private Sub optAll_Click()
    Call PatientSearch
End Sub

Private Sub optFemale_Click()
    Call PatientSearch
End Sub

Private Sub optMale_Click()
    Call PatientSearch
End Sub

Private Sub txtByAddress_Change()
    Call PatientSearch
    SearchCount = SearchCount + 1
End Sub

Private Sub txtName_Change()
    SearchCount = SearchCount + 1
    Call PatientSearch
End Sub

Private Sub PatientSearch()
    Select Case SSTab3.Tab
        Case 0: NameSearch
        Case 1: AddressSearch
        Case 2: DOBSearch
    End Select
End Sub

Private Sub AddressSearch()
    On Error Resume Next
    Dim OldSearchCount As Long
    Call FormatGrid
    OldSearchCount = SearchCount
    If Trim(txtByAddress.Text) = "" Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    temSql = "Select * from tblPatient where "
    If optMale.Value = True Then
        temSql = temSql & " SexID = 1 AND "
    ElseIf optFemale.Value = True Then
        temSql = temSql & " SexID = 2 AND "
    Else
        
    End If
    temSql = temSql & " Address like '%" & Trim(txtByAddress.Text) & "%' order by Patient"
    With rsTem
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        gridPatient.Visible = False
        If .RecordCount > 0 Then
            While .EOF = False
                DoEvents
                If SearchCount <> OldSearchCount Then
                    gridPatient.Visible = True
                    Exit Sub
                End If
                gridPatient.Rows = gridPatient.Rows + 1
                gridPatient.row = gridPatient.Rows - 1
                gridPatient.col = 0
                gridPatient.Text = !PatientID
                gridPatient.col = 1
                gridPatient.Text = !Patient
                gridPatient.col = 2
                gridPatient.Text = CalculateAgeInWords(!DateOfBirth)
                .MoveNext
            Wend
        End If
        gridPatient.Visible = True
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub DOBSearch()
    On Error Resume Next

    Dim OldSearchCount As Long
    Call FormatGrid
    OldSearchCount = SearchCount
    Dim rsTem As New ADODB.Recordset
    temSql = "Select * from tblPatient where "
    If optMale.Value = True Then
        temSql = temSql & " SexID = 1 AND "
    ElseIf optFemale.Value = True Then
        temSql = temSql & " SexID = 2 AND "
    Else
        
    End If
    temSql = temSql & " DateOfBirth = #" & Format(dtpDOB.Value, "dd MMMM yyyy") & "# order by Patient"
    With rsTem
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        gridPatient.Visible = False
        If .RecordCount > 0 Then
            While .EOF = False
                DoEvents
                If SearchCount <> OldSearchCount Then
                    gridPatient.Visible = True
                    Exit Sub
                End If
                gridPatient.Rows = gridPatient.Rows + 1
                gridPatient.row = gridPatient.Rows - 1
                gridPatient.col = 0
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = !PatientID
                gridPatient.col = 1
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = !Patient
                gridPatient.col = 2
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = CalculateAgeInWords(!DateOfBirth)
                .MoveNext
            Wend
        End If
        gridPatient.Visible = True
        .Close
    End With
    Set rsTem = Nothing
End Sub


Private Sub NameSearch()
    On Error Resume Next

    Dim OldSearchCount As Long
    Call FormatGrid
    OldSearchCount = SearchCount
    If Trim(txtName.Text) = "" Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    temSql = "Select * from tblPatient where "
    If optMale.Value = True Then
        temSql = temSql & " SexID = 1 AND "
    ElseIf optFemale.Value = True Then
        temSql = temSql & " SexID = 2 AND "
    Else
        
    End If
    temSql = temSql & " Patient like '%" & Trim(txtName.Text) & "%' order by Patient"
    With rsTem
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        gridPatient.Visible = False
        If .RecordCount > 0 Then
            While .EOF = False
                DoEvents
                If SearchCount <> OldSearchCount Then
                    gridPatient.Visible = True
                    Exit Sub
                End If
                gridPatient.Rows = gridPatient.Rows + 1
                gridPatient.row = gridPatient.Rows - 1
                gridPatient.col = 0
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                
                gridPatient.Text = !PatientID
                gridPatient.col = 1
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = !Patient
                gridPatient.col = 2
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = CalculateAgeInWords(!DateOfBirth)
                .MoveNext
            Wend
        End If
        gridPatient.Visible = True
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FormatGrid()
    With gridPatient
        .Clear
        .Rows = 1
        .Cols = 3
        .col = 0
        .Text = "ID"
        .col = 1
        .Text = "Name"
        .col = 2
        .Text = "Age"
        .ColWidth(0) = 700
        .ColWidth(2) = 1800
        .ColWidth(1) = 3000
    End With
End Sub

Private Sub lstImages_Click()
    Dim rsImg As New ADODB.Recordset
    Dim temID As Long
    With rsImg
        temSql = "Select * from tblImage where ImageID = " & Val(lstImages.BoundText) & " AND Deleted = False  Order by PatientDefault, ImageID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temID = !ImageID
            imgPatient.Stretch = True
            FillImage rsImg, "ImageData", "ImageSize", imgPatient
            txtImage.Text = !Image & vbNewLine & !Comments
        Else
            temID = 0
            imgPatient.Picture = LoadPicture()
        End If
    End With
End Sub

Private Sub lstIx_Click()
    Dim rsTemIxItem As New ADODB.Recordset
    Dim rsTemIxValue As New ADODB.Recordset
    
    Dim temPatientIXID As Long
    Dim temPtIxID As Long
    Dim temIxID As Long
    Dim temPtID As Long
    Dim temText As String
    
    temPatientIXID = Val(lstIx.BoundText)
    temPtID = CurrentPatient.ID
    
    
    With rsTemIxItem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientIX where PatientIXID = " & temPatientIXID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temIxID = !IxID
        End If
        
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientIx where PatientIXID = " & temPatientIXID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temPtIxID = !PatientIxID
        End If
        
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem where Deleted = false AND (IsValue = true or IsCalc = true or IsFlag = true  ) AND IxID = " & temIxID & " Order by format(X1, '0.0' ) , Y1 "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temText = temText & !IxItem & " - "
                
                If rsTemIxValue.State = 1 Then rsTemIxValue.Close
                temSql = "SELECT * from tblPatientIxItem where PatientIxId = " & temPtIxID & " AND IxItemID = " & !IxItemID
                rsTemIxValue.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If rsTemIxValue.RecordCount > 0 Then
                    temText = temText & rsTemIxValue!Value & vbNewLine
                Else
                    temText = temText & vbNewLine
                End If
                
                .MoveNext
            Wend
        End If
        .Close
    End With
    txtIx.Text = temText
End Sub

Private Sub lstVisit_Click()
    Dim temText As String
    Dim rsVisit As New ADODB.Recordset
    If Val(lstVisit.BoundText) = 0 Then Exit Sub
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
    
    txtVisit.Text = temText
End Sub

Private Sub txtAge_LostFocus()
    If Abs(DateDiff("yyyy", DateSerial(Year(Date) - Val(txtAge.Text), 1, 1), dtpDOB.Value)) > 2 Then
        dtpDOB.Value = DateSerial(Year(Date) - Val(txtAge.Text), 1, 1)
    End If
End Sub

Private Sub dtpDOB_LostFocus()
    If Abs(DateDiff("y", DateSerial(Year(Date) - Val(txtAge.Text), 6, 31), dtpDOB.Value)) > 2 Then
        txtAge.Text = Abs(DateDiff("yyyy", dtpDOB.Value, Date))
    End If

End Sub

Private Sub lstVisit_DblClick()
    CurrentVisit.VisitID = Val(lstVisit.BoundText)
    currentDoctorId = Val(cmbDoctor.BoundText)
    Unload frmNewVisit
    frmNewVisit.Show
    frmNewVisit.ZOrder 0
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbPatient.Enabled = False
    
    SSTab1.Enabled = False
    SSTab2.Enabled = False
    btnNewVisit.Enabled = False
    btnToQueue.Enabled = False
    btnPastHistory.Enabled = False
    btnImage.Enabled = False
    btnIx.Enabled = False
    
    txtPatient.Enabled = True
    txtComments.Enabled = True
    txtFullName.Enabled = True
    txtAddress.Enabled = True
    txtPhone.Enabled = True
    txtCode.Enabled = True
    txtMobile.Enabled = True
    txtAge.Enabled = True
    txteMail.Enabled = True
    txtID.Enabled = True
    cmbTitle.Enabled = True
    cmbSex.Enabled = True
    cmbRace.Enabled = True
    cmbBloodGroup.Enabled = True
    cmbCivilStatus.Enabled = True
    dtpDOB.Enabled = True
    
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    btnSaveNewVisit.Enabled = True
    btnSaveQueue.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbPatient.Enabled = True

    SSTab1.Enabled = True
    SSTab2.Enabled = True
    btnNewVisit.Enabled = True
    btnToQueue.Enabled = True
    btnPastHistory.Enabled = True
    btnImage.Enabled = True
    btnIx.Enabled = True
    
    txtPatient.Enabled = False
    txtComments.Enabled = False
    txtCode.Enabled = False
    txtFullName.Enabled = False
    txtAddress.Enabled = False
    txtPhone.Enabled = False
    txtMobile.Enabled = False
    txtAge.Enabled = False
    txteMail.Enabled = False
    txtID.Enabled = False
    cmbTitle.Enabled = False
    cmbSex.Enabled = False
    cmbRace.Enabled = False
    cmbBloodGroup.Enabled = False
    cmbCivilStatus.Enabled = False
    dtpDOB.Enabled = False
    
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
    btnSaveNewVisit.Enabled = False
    btnSaveQueue.Enabled = False
    
End Sub

Private Sub ClearValues()
    txtPatient.Text = Empty
    txtComments.Text = Empty
    txtFullName.Text = Empty
    txtAddress.Text = Empty
    txtCode.Text = Empty
    txtPhone.Text = Empty
    txtMobile.Text = Empty
    txtAge.Text = Empty
    txteMail.Text = Empty
    cmbTitle.Text = Empty
    cmbSex.Text = Empty
    cmbRace.Text = Empty
    cmbBloodGroup.Text = Empty
    cmbCivilStatus.Text = Empty
    lblID.Visible = False
    txtID.Visible = False
    imgPatient.Picture = LoadPicture()
    txtVisit.Text = Empty
    txtImage.Text = Empty
    dtpDOB.Value = Date
    
    txtImage.Text = Empty
    txtIx.Text = Empty
    txtVisit.Text = Empty
    
    
End Sub

Private Function SaveNew() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatient"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Patient = txtPatient.Text
        !Comments = txtComments.Text
        !FullName = txtFullName.Text
        !Address = txtAddress.Text
        !PatientCode = txtCode.Text
        !Telephone = txtPhone.Text
        !Mobile = txtMobile.Text
        !eMail = txteMail.Text
        !TitleID = Val(cmbTitle.BoundText)
        !SexID = Val(cmbSex.BoundText)
        !RaceID = Val(cmbRace.BoundText)
        !BloodGroupID = Val(cmbBloodGroup.BoundText)
        !CivilStatusID = Val(cmbCivilStatus.BoundText)
        !DateOfBirth = dtpDOB.Value
        .Update
        CurrentPatient.ID = !PatientID
        SaveNew = !PatientID
        .Close
    End With
    Set rsTem = Nothing
End Function

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Patient = txtPatient.Text
            !Comments = txtComments.Text
            !FullName = txtFullName.Text
            !Address = txtAddress.Text
            !PatientCode = txtCode.Text
            !Telephone = txtPhone.Text
            !Mobile = txtMobile.Text
            !eMail = txteMail.Text
            !TitleID = Val(cmbTitle.BoundText)
            !SexID = Val(cmbSex.BoundText)
            !RaceID = Val(cmbRace.BoundText)
            !BloodGroupID = Val(cmbBloodGroup.BoundText)
            !CivilStatusID = Val(cmbCivilStatus.BoundText)
            !DateOfBirth = dtpDOB.Value
            .Update
        End If
        CurrentPatient.ID = !PatientID
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim Pt As New clsFillCombos
    Pt.FillAnyCombo cmbPatient, "Patient", True
End Sub

Private Sub DisplayDetails()
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
            txtPatient.Text = CurrentPatient.Patient ' !Patient
            txtComments.Text = CurrentPatient.Comments   'Format(!FamilyName, "")
            txtFullName.Text = CurrentPatient.FullName ' Format(!GivenName, "")
            txtAddress.Text = CurrentPatient.Address ' Format(!Address, "")
            txtPhone.Text = CurrentPatient.Telephone ' Format(!Telephone, "")
            txtMobile.Text = CurrentPatient.Mobile ' Format(!Mobile, "")
            txteMail.Text = CurrentPatient.eMail ' Format(!Email, "")
            txtCode.Text = CurrentPatient.Code ' Format(!PatientCode, "")
            cmbTitle.BoundText = CurrentPatient.TitleID ' !TitleID
            cmbSex.BoundText = CurrentPatient.SexID ' !SexID
            cmbRace.BoundText = CurrentPatient.RaceID '  !RaceID
            cmbBloodGroup.BoundText = CurrentPatient.BloodGroupID   '!BloodGroupID
            cmbCivilStatus.BoundText = CurrentPatient.CivilStatusID ' !CivilStatusID
            txtComments.Text = CurrentPatient.Comments
            'If IsNull(!DateOfBirth) = False Then dtpDOB.Value = !DateOfBirth
            dtpDOB.Value = CurrentPatient.DateOfBirth
            lblID.Visible = True
            txtID.Visible = True
            txtID.Text = CurrentPatient.ID '  !PatientID
'        End If
'        .Close
'    End With
'    Set rsTem = Nothing
End Sub

Private Sub txtSearchID_Change()
    cmbPatient.BoundText = Val(txtSearchID.Text)
End Sub
