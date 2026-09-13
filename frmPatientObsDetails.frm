VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPatientObsDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Details"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12840
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
   ScaleHeight     =   9330
   ScaleWidth      =   12840
   Begin VB.TextBox txtWeight 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox txtSBP 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1200
      TabIndex        =   16
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox txtDBP 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1920
      TabIndex        =   17
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox txtSearchID 
      Height          =   360
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   3360
      TabIndex        =   20
      Top             =   5400
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Past Visits"
      TabPicture(0)   =   "frmPatientObsDetails.frx":0000
      Tab(0).ControlEnabled=   0   'False
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
      TabPicture(1)   =   "frmPatientObsDetails.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label26"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label27"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lstDrugs"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstDx"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "A&llergies"
      TabPicture(2)   =   "frmPatientObsDetails.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label28"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label29"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lstFoodAllergies"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lstAllergies"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Investigations"
      TabPicture(3)   =   "frmPatientObsDetails.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label30"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label31"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lstIx"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtIx"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "&mages"
      TabPicture(4)   =   "frmPatientObsDetails.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "imgPatient"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label32"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label33"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lstImages"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txtImage"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).ControlCount=   5
      Begin VB.TextBox txtImage 
         Height          =   2055
         Left            =   3600
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtIx 
         Height          =   1980
         Left            =   -70800
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   735
         Width           =   5055
      End
      Begin MSDataListLib.DataList lstVisit 
         Height          =   1980
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3493
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin VB.TextBox txtVisit 
         Height          =   2025
         Left            =   -71640
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   720
         Width           =   5895
      End
      Begin MSDataListLib.DataList lstDx 
         Height          =   1980
         Left            =   -74880
         TabIndex        =   26
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3493
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstDrugs 
         Height          =   1980
         Left            =   -70560
         TabIndex        =   28
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3493
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstAllergies 
         Height          =   2100
         Left            =   -74880
         TabIndex        =   30
         Top             =   660
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3704
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstIx 
         Height          =   1980
         Left            =   -74880
         TabIndex        =   34
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3493
         _Version        =   393216
      End
      Begin MSDataListLib.DataList lstImages 
         Height          =   2025
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3572
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin MSDataListLib.DataList lstFoodAllergies 
         Height          =   2100
         Left            =   -69840
         TabIndex        =   32
         Top             =   660
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3704
         _Version        =   393216
         IntegralHeight  =   0   'False
      End
      Begin VB.Label Label33 
         Caption         =   "Image &Details"
         Height          =   255
         Left            =   3600
         TabIndex        =   39
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label32 
         Caption         =   "I&mages"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "&Results"
         Height          =   255
         Left            =   -70800
         TabIndex        =   35
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label30 
         Caption         =   "&Investigation"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "&Food Allergies"
         Height          =   255
         Left            =   -69840
         TabIndex        =   31
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label28 
         Caption         =   "&Drug Allergies"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label27 
         Caption         =   "Drug &History"
         Height          =   255
         Left            =   -70560
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "&Diagnoses"
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Visit &Details"
         Height          =   255
         Left            =   -71640
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "&Visit"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image imgPatient 
         Height          =   2475
         Left            =   6360
         Top             =   360
         Width           =   3000
      End
   End
   Begin VB.TextBox txtCode 
      Height          =   360
      Left            =   8760
      TabIndex        =   66
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtID 
      Height          =   360
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtMobile 
      Height          =   360
      Left            =   9000
      TabIndex        =   74
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txteMail 
      Height          =   360
      Left            =   11280
      TabIndex        =   76
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtPhone 
      Height          =   360
      Left            =   6480
      TabIndex        =   72
      Top             =   4320
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpDOB 
      Height          =   375
      Left            =   6480
      TabIndex        =   58
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   74121219
      CurrentDate     =   39845
   End
   Begin VB.TextBox txtAge 
      Height          =   360
      Left            =   8760
      TabIndex        =   60
      Top             =   2400
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo cmbTitle 
      Height          =   360
      Left            =   6480
      TabIndex        =   52
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtAddress 
      Height          =   840
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   70
      Top             =   3360
      Width           =   6255
   End
   Begin VB.TextBox txtFamilyName 
      Height          =   360
      Left            =   6480
      TabIndex        =   50
      Top             =   1440
      Width           =   6255
   End
   Begin VB.TextBox txtGivenName 
      Height          =   360
      Left            =   6480
      TabIndex        =   48
      Top             =   960
      Width           =   6255
   End
   Begin VB.TextBox txtPatient 
      Height          =   360
      Left            =   6480
      TabIndex        =   46
      Top             =   480
      Width           =   6255
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   1800
      TabIndex        =   3
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   3360
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataListLib.DataCombo cmbPatient 
      Height          =   4260
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7514
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCivilStatus 
      Height          =   360
      Left            =   11160
      TabIndex        =   56
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSex 
      Height          =   360
      Left            =   8760
      TabIndex        =   54
      Top             =   1920
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
      Left            =   11160
      TabIndex        =   62
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   11520
      TabIndex        =   44
      Top             =   8880
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
      Left            =   11520
      TabIndex        =   80
      Top             =   4920
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
      Left            =   5160
      TabIndex        =   77
      Top             =   4920
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
      Left            =   6480
      TabIndex        =   64
      Top             =   2880
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
      Left            =   6480
      TabIndex        =   78
      Top             =   4920
      Width           =   2295
      _ExtentX        =   4048
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
      TabIndex        =   79
      Top             =   4920
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
      Left            =   5400
      TabIndex        =   42
      Top             =   8400
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   3360
      TabIndex        =   41
      Top             =   8400
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   7440
      TabIndex        =   43
      Top             =   8400
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   74121219
      CurrentDate     =   39859
   End
   Begin MSDataListLib.DataCombo cmbSecession 
      Height          =   360
      Left            =   1200
      TabIndex        =   12
      Top             =   6720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDoctor 
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnToQueue 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   8160
      Width           =   3135
      _ExtentX        =   5530
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
      TabIndex        =   19
      Top             =   8640
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.Label Label23 
      Caption         =   "&ID"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "&Secession"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Docto&r"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Da&te"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "kg"
      Height          =   255
      Left            =   2760
      TabIndex        =   81
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "&BP                               mmHg"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Label Label21 
      Caption         =   "&Weight"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "&Blood Group"
      Height          =   255
      Left            =   5160
      TabIndex        =   63
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblID 
      Caption         =   "ID"
      Height          =   255
      Left            =   10320
      TabIndex        =   67
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "C&ode"
      Height          =   255
      Left            =   8280
      TabIndex        =   65
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "&Mobile"
      Height          =   255
      Left            =   8280
      TabIndex        =   73
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "&eMail"
      Height          =   255
      Left            =   10440
      TabIndex        =   75
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "&Phone"
      Height          =   255
      Left            =   5160
      TabIndex        =   71
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "A&ddress"
      Height          =   255
      Left            =   5160
      TabIndex        =   69
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "&Age"
      Height          =   255
      Left            =   8280
      TabIndex        =   59
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "&Date of Birth"
      Height          =   255
      Left            =   5160
      TabIndex        =   57
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "&Race"
      Height          =   255
      Left            =   10320
      TabIndex        =   61
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Ci&vil St."
      Height          =   255
      Left            =   10320
      TabIndex        =   55
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Se&x"
      Height          =   255
      Left            =   8280
      TabIndex        =   53
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "&Title"
      Height          =   255
      Left            =   5160
      TabIndex        =   51
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "&Family Name"
      Height          =   255
      Left            =   5160
      TabIndex        =   49
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "&Given Name"
      Height          =   255
      Left            =   5160
      TabIndex        =   47
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Na&me"
      Height          =   255
      Left            =   5160
      TabIndex        =   45
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Patient"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatientObsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
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
    SendKeys "{home}+{end}"
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
    SendKeys "{home}+{end}"
End Sub

Private Sub btnImage_Click()
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
    If IsNumeric(cmbDoctor.BoundText) = False Then
        MsgBox "Please select a doctor"
        cmbDoctor.SetFocus
        Exit Sub
    End If
    
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
    Dim TemPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        TemPatientID = cmbPatient.BoundText
        Call SaveOld
    Else
        TemPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.SetFocus
    cmbPatient.BoundText = TemPatientID
End Sub

Private Sub btnSaveNewVisit_Click()
    Dim TemPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        TemPatientID = cmbPatient.BoundText
        Call SaveOld
    Else
        TemPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.BoundText = TemPatientID
    
    Unload frmNewVisit
    Dim NewVisitID As Long
    NewVisitID = CurrentVisit.NewVisitID(Val(cmbPatient.BoundText), Date, UserID)
    frmNewVisit.Show
    frmNewVisit.ZOrder 0
    
End Sub

Private Sub btnSaveQueue_Click()
    Dim TemPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        Call SaveOld
        TemPatientID = cmbPatient.BoundText
    Else
        TemPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.BoundText = TemPatientID
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
        !Weight = Val(txtWeight.Text)
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
        !Weight = Val(txtWeight.Text)
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
    Call ClearValues
    CurrentPatient.ID = Val(cmbPatient.BoundText)
    If IsNumeric(cmbPatient.BoundText) = True Then
        Call DisplayDetails
        txtAge.Text = CurrentPatient.AgeInYears
        Call ListVisits
        Call ListImages
    End If
End Sub

Private Sub ListImages()
    Dim rsImg As New ADODB.Recordset
    Dim TemID As Long
    
    With rsImg
        temSql = "Select * from tblImage where PatientID = " & CurrentPatient.ID & " AND Deleted = False  Order by PatientDefault, ImageID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            TemID = !ImageID
            imgPatient.Stretch = True
            FillImage rsImg, "ImageData", "ImageSize", imgPatient
            txtImage.Text = !Image & vbNewLine & !Comments
        Else
            TemID = 0
            imgPatient.Picture = LoadPicture()
        End If
    End With
    With lstImages
        Set .RowSource = rsImg
        .ListField = "Image"
        .BoundColumn = "ImageID"
        .BoundText = TemID
    End With
End Sub

Private Sub ListVisits()
    Dim TemID As Long
    With rsPastVisits
        If .State = 1 Then .Close
        temSql = "Select tblVisit.VisitID, tblVisit.VisitDate & ' - ' & SBP & '/' & DBP as VisitDisplay from tblVisit where PatientID = " & Val(cmbPatient.BoundText) & " And Deleted = FALSE Order by VisitID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            TemID = !VisitID
        Else
            TemID = 0
        End If
    End With
    With lstVisit
        Set .RowSource = rsPastVisits
        .ListField = "VisitDisplay"
        .BoundColumn = "VisitID"
        .BoundText = TemID
        Call lstVisit_Click
    End With
End Sub

Private Sub Form_Load()
    Call SelectMode
    Call ClearValues
    Call FillCombos
    Call FillOnceCombo
    Call GetFormSettings
    dtpDate.Value = Date
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetFormSettings
End Sub

Private Sub SetFormSettings()
    SaveSetting App.EXEName, Me.Name, "DoctorID", Val(cmbDoctor.BoundText)
    SaveSetting App.EXEName, Me.Name, "SecessionID", Val(cmbSecession.BoundText)
End Sub


Private Sub imgPatient_DblClick()
    Dim MyFrm As Form
    Set MyFrm = New frmPatientImage
    MyFrm.Show
    MyFrm.ZOrder 0
    MyFrm.Caption = lstImages.Text
End Sub

Private Sub lstImages_Click()
    Dim rsImg As New ADODB.Recordset
    Dim TemID As Long
    With rsImg
        temSql = "Select * from tblImage where ImageID = " & Val(lstImages.BoundText) & " AND Deleted = False  Order by PatientDefault, ImageID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            TemID = !ImageID
            imgPatient.Stretch = True
            FillImage rsImg, "ImageData", "ImageSize", imgPatient
            txtImage.Text = !Image & vbNewLine & !Comments
        Else
            TemID = 0
            imgPatient.Picture = LoadPicture()
        End If
    End With
End Sub

Private Sub lstVisit_Click()
    Dim temText As String
    Dim rsVisit As New ADODB.Recordset
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

Private Sub lstVisit_DblClick()
    CurrentVisit.VisitID = Val(lstVisit.BoundText)
    Unload frmNewVisit
    frmNewVisit.Show
    frmNewVisit.ZOrder 0
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpDOB.Value = DateSerial(Year(Date) - Val(txtAge.Text), 6, 31)
    End If
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbPatient.Enabled = False
    
    SSTab1.Enabled = False
    
    txtPatient.Enabled = True
    txtFamilyName.Enabled = True
    txtGivenName.Enabled = True
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
    
    txtPatient.Enabled = False
    txtFamilyName.Enabled = False
    txtCode.Enabled = False
    txtGivenName.Enabled = False
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
    txtFamilyName.Text = Empty
    txtGivenName.Text = Empty
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
End Sub

Private Function SaveNew() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatient"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Patient = txtPatient.Text
        !FamilyName = txtFamilyName.Text
        !GivenName = txtGivenName.Text
        !Address = txtAddress.Text
        !PatientCode = txtCode.Text
        !Telephone = txtPhone.Text
        !Mobile = txtMobile.Text
        !Email = txteMail.Text
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
            !FamilyName = txtFamilyName.Text
            !GivenName = txtGivenName.Text
            !Address = txtAddress.Text
            !PatientCode = txtCode.Text
            !Telephone = txtPhone.Text
            !Mobile = txtMobile.Text
            !Email = txteMail.Text
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
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtPatient.Text = !Patient
            txtFamilyName.Text = Format(!FamilyName, "")
            txtGivenName.Text = Format(!GivenName, "")
            txtAddress.Text = Format(!Address, "")
            txtPhone.Text = Format(!Telephone, "")
            txtMobile.Text = Format(!Mobile, "")
            txteMail.Text = Format(!Email, "")
            txtCode.Text = Format(!PatientCode, "")
            cmbTitle.BoundText = !TitleID
            cmbSex.BoundText = !SexID
            cmbRace.BoundText = !RaceID
            cmbBloodGroup.BoundText = !BloodGroupID
            cmbCivilStatus.BoundText = !CivilStatusID
            dtpDOB.Value = !DateOfBirth
            lblID.Visible = True
            txtID.Visible = True
            txtID.Text = !PatientID
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub txtSearchID_Change()
    cmbPatient.BoundText = Val(txtSearchID.Text)
End Sub
