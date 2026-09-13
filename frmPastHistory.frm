VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPastHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12300
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
   ScaleHeight     =   6990
   ScaleWidth      =   12300
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Diagnoses"
      TabPicture(0)   =   "frmPastHistory.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnIxGraph"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnWeightGraph"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnBPGraph"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnDxDelete"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "gridDiagnosis"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbDiagnosis"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDiagnosis"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btnDxAdd"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDxID"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDxID1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Medications"
      TabPicture(1)   =   "frmPastHistory.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDeletedID"
      Tab(1).Control(1)=   "txtAddID"
      Tab(1).Control(2)=   "chkIndoor"
      Tab(1).Control(3)=   "txtIssueQty"
      Tab(1).Control(4)=   "txtDuration"
      Tab(1).Control(5)=   "txtDose"
      Tab(1).Control(6)=   "cmbItem"
      Tab(1).Control(7)=   "cmbDoseUnit"
      Tab(1).Control(8)=   "cmbFrequencyUnit"
      Tab(1).Control(9)=   "cmbDurationUnit"
      Tab(1).Control(10)=   "cmbIssueUnit"
      Tab(1).Control(11)=   "gridItem"
      Tab(1).Control(12)=   "btnRxAdd"
      Tab(1).Control(13)=   "btnRxDelete"
      Tab(1).Control(14)=   "cmbCategory"
      Tab(1).Control(15)=   "Label2"
      Tab(1).Control(16)=   "Label8"
      Tab(1).Control(17)=   "Label7"
      Tab(1).Control(18)=   "Label6"
      Tab(1).Control(19)=   "Label5"
      Tab(1).Control(20)=   "Label4"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Allergies"
      TabPicture(2)   =   "frmPastHistory.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtFoodItem"
      Tab(2).Control(1)=   "txtAllergy"
      Tab(2).Control(2)=   "cmbAllergy"
      Tab(2).Control(3)=   "gridAllergy"
      Tab(2).Control(4)=   "btnAxAdd"
      Tab(2).Control(5)=   "btnAxDelete"
      Tab(2).Control(6)=   "cmbFoodItem"
      Tab(2).Control(7)=   "gridFoodItem"
      Tab(2).Control(8)=   "btnFoodItemAdd"
      Tab(2).Control(9)=   "btnFoodItemDelete"
      Tab(2).Control(10)=   "txtAxID"
      Tab(2).Control(11)=   "txtAxID1"
      Tab(2).Control(12)=   "txtFoodItemAddID"
      Tab(2).Control(13)=   "txtFoodItemDeleteID"
      Tab(2).Control(14)=   "Label9"
      Tab(2).Control(15)=   "Drugs"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Procedures"
      TabPicture(3)   =   "frmPastHistory.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtProcedure"
      Tab(3).Control(1)=   "cmbProcedure"
      Tab(3).Control(2)=   "MSFlexGrid3"
      Tab(3).Control(3)=   "btnPxAdd"
      Tab(3).Control(4)=   "btnPxDelete"
      Tab(3).Control(5)=   "Label3"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Obstetric"
      TabPicture(4)   =   "frmPastHistory.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "gridObs"
      Tab(4).Control(1)=   "btnObsAdd"
      Tab(4).Control(2)=   "btnObsDelete"
      Tab(4).Control(3)=   "frameObs"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Family History"
      TabPicture(5)   =   "frmPastHistory.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label10"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "btnFHDxDelete"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "gridFHDiagnosis"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmbFHDiagnosis"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "btnFHDxAdd"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "txtFHDxID1"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "txtFHDxID"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "txtFHDiagnosis"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).ControlCount=   8
      Begin VB.TextBox txtFHDiagnosis 
         Height          =   855
         Left            =   -73800
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   960
         Width           =   9255
      End
      Begin VB.TextBox txtFHDxID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -64440
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtFHDxID1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -64440
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid gridObs 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   59
         Top             =   2880
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5530
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin btButtonEx.ButtonEx btnObsAdd 
         Height          =   375
         Left            =   -64320
         TabIndex        =   57
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
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
      Begin btButtonEx.ButtonEx btnObsDelete 
         Height          =   375
         Left            =   -64320
         TabIndex        =   58
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.Frame frameObs 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   56
         Top             =   360
         Width           =   11655
         Begin VB.TextBox txtObsDelID 
            Height          =   495
            Left            =   7680
            TabIndex        =   61
            Top             =   1200
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtObsEditID 
            Height          =   495
            Left            =   7680
            TabIndex        =   60
            Top             =   1680
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.TextBox txtFoodItem 
         Height          =   375
         Left            =   -70680
         TabIndex        =   37
         Top             =   3120
         Width           =   6255
      End
      Begin VB.TextBox txtDeletedID 
         Height          =   360
         Left            =   -64200
         TabIndex        =   53
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAddID 
         Height          =   360
         Left            =   -64200
         TabIndex        =   52
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkIndoor 
         Caption         =   "Indoor"
         Height          =   255
         Left            =   -64320
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtDxID1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9240
         TabIndex        =   49
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDxID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9240
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtIssueQty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -66600
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDuration 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -69480
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDose 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -74400
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtProcedure 
         Height          =   375
         Left            =   -69720
         TabIndex        =   43
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtAllergy 
         Height          =   375
         Left            =   -70680
         TabIndex        =   31
         Top             =   480
         Width           =   6135
      End
      Begin btButtonEx.ButtonEx btnDxAdd 
         Height          =   375
         Left            =   10680
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.TextBox txtDiagnosis 
         Height          =   855
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   9135
      End
      Begin MSDataListLib.DataCombo cmbDiagnosis 
         Height          =   360
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid gridDiagnosis 
         Height          =   3615
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6376
         _Version        =   393216
         WordWrap        =   -1  'True
      End
      Begin btButtonEx.ButtonEx btnDxDelete 
         Height          =   375
         Left            =   10680
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSDataListLib.DataCombo cmbAllergy 
         Height          =   360
         Left            =   -73800
         TabIndex        =   30
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid gridAllergy 
         Height          =   1935
         Left            =   -73800
         TabIndex        =   33
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3413
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx btnAxAdd 
         Height          =   375
         Left            =   -64320
         TabIndex        =   32
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
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
      Begin btButtonEx.ButtonEx btnAxDelete 
         Height          =   375
         Left            =   -64320
         TabIndex        =   34
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSDataListLib.DataCombo cmbProcedure 
         Height          =   360
         Left            =   -73800
         TabIndex        =   42
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   45
         Top             =   960
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8070
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx btnPxAdd 
         Height          =   375
         Left            =   -64200
         TabIndex        =   44
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
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
      Begin btButtonEx.ButtonEx btnPxDelete 
         Height          =   375
         Left            =   -64200
         TabIndex        =   46
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
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
      Begin MSDataListLib.DataCombo cmbItem 
         Height          =   360
         Left            =   -74400
         TabIndex        =   11
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDoseUnit 
         Height          =   360
         Left            =   -73680
         TabIndex        =   16
         Top             =   960
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
         Left            =   -71640
         TabIndex        =   18
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDurationUnit 
         Height          =   360
         Left            =   -68760
         TabIndex        =   21
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbIssueUnit 
         Height          =   360
         Left            =   -65880
         TabIndex        =   24
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid gridItem 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   27
         Top             =   1440
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx btnRxAdd 
         Height          =   375
         Left            =   -64320
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
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
      Begin btButtonEx.ButtonEx btnRxDelete 
         Height          =   375
         Left            =   -64320
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSDataListLib.DataCombo cmbCategory 
         Height          =   360
         Left            =   -66600
         TabIndex        =   13
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx btnBPGraph 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Blood Pressure Graph"
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
      Begin btButtonEx.ButtonEx btnWeightGraph 
         Height          =   375
         Left            =   5400
         TabIndex        =   9
         Top             =   5640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Weight Graph"
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
      Begin MSDataListLib.DataCombo cmbFoodItem 
         Height          =   360
         Left            =   -73800
         TabIndex        =   36
         Top             =   3120
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid gridFoodItem 
         Height          =   1935
         Left            =   -73800
         TabIndex        =   39
         Top             =   3600
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3413
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx btnFoodItemAdd 
         Height          =   375
         Left            =   -64320
         TabIndex        =   38
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
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
      Begin btButtonEx.ButtonEx btnFoodItemDelete 
         Height          =   375
         Left            =   -64320
         TabIndex        =   40
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.TextBox txtAxID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -65760
         TabIndex        =   50
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAxID1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -65760
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtFoodItemAddID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -65760
         TabIndex        =   54
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtFoodItemDeleteID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -65760
         TabIndex        =   55
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin btButtonEx.ButtonEx btnIxGraph 
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   5640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Investigaion Graph"
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
      Begin btButtonEx.ButtonEx btnFHDxAdd 
         Height          =   375
         Left            =   -64440
         TabIndex        =   64
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSDataListLib.DataCombo cmbFHDiagnosis 
         Height          =   360
         Left            =   -73800
         TabIndex        =   66
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid gridFHDiagnosis 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   67
         Top             =   1920
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6376
         _Version        =   393216
         WordWrap        =   -1  'True
      End
      Begin btButtonEx.ButtonEx btnFHDxDelete 
         Height          =   375
         Left            =   -64440
         TabIndex        =   68
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&Diagnosis"
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Food Item"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cat."
         Height          =   255
         Left            =   -67200
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Issue"
         Height          =   255
         Left            =   -67200
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Refills"
         Height          =   255
         Left            =   -70080
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Freq."
         Height          =   255
         Left            =   -72120
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dose"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Procedure"
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Drugs 
         BackStyle       =   0  'Transparent
         Caption         =   "Drugs"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Diagnosis"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11040
      TabIndex        =   47
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
End
Attribute VB_Name = "frmPastHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
    Dim AppExcel As Excel.Application
    
    Dim myworkbook As Excel.Workbook
    
    Dim myWorkSheet1 As Excel.Worksheet
    Dim myWorkSheet2 As Excel.Worksheet
    Dim myWorkSheet3 As Excel.Worksheet
    
    Dim FilledOnce As Boolean
    Dim ObsFilledOnce As Boolean
    
    Dim FormPatient As New clsPatient
        
Private Sub btnBPGraph_Click()
    frmBPGraph.Show
    frmBPGraph.ZOrder 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDxAdd_Click()
    If IsNumeric(cmbDiagnosis.BoundText) = False Then
        MsgBox "Please select a diagnosis"
        cmbDiagnosis.SetFocus
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtDxID.Text) <> 0 Then
            temSql = "Select * from tblPatientDiagnosis where PatientDiagnosisID = " & Val(txtDxID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !DiagnosisID = Val(cmbDiagnosis.BoundText)
                !Comments = txtDiagnosis.Text
                .Update
            Else
                .AddNew
                !PatientID = FormPatient.ID
                !DiagnosisID = Val(cmbDiagnosis.BoundText)
                !Comments = txtDiagnosis.Text
                .Update
            End If
        Else
            temSql = "Select * from tblPatientDiagnosis"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
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
    txtDxID.Text = Empty
    cmbDiagnosis.SetFocus
End Sub

Private Sub FillDiagnosis()
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientDiagnosis.PatientDiagnosisID, tblPatientDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblPatientDiagnosis ON tblDiagnosis.DiagnosisID = tblPatientDiagnosis.DiagnosisID where tblPatientDiagnosis.Deleted = False AND tblPatientDiagnosis.PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridDiagnosis.Clear
        gridDiagnosis.Rows = 1
        gridDiagnosis.Cols = 4
        gridDiagnosis.ColWidth(0) = 0
        gridDiagnosis.ColWidth(1) = 3000
        gridDiagnosis.ColWidth(2) = gridDiagnosis.Width - 3000 - 150
        gridDiagnosis.ColWidth(3) = 0
        gridDiagnosis.Row = 0
        gridDiagnosis.Col = 1
        gridDiagnosis.Text = "Diagnosis"
        gridDiagnosis.Col = 2
        gridDiagnosis.Text = "Details"
        
        While .EOF = False
            gridDiagnosis.Rows = gridDiagnosis.Rows + 1
            gridDiagnosis.Row = gridDiagnosis.Rows - 1
            gridDiagnosis.Col = 0
            gridDiagnosis.Text = !PatientDiagnosisID
            gridDiagnosis.Col = 1
            gridDiagnosis.Text = !Diagnosis
            gridDiagnosis.Col = 2
            gridDiagnosis.Text = Format(!Comments, "")
            gridDiagnosis.Col = 3
            gridDiagnosis.Text = !DiagnosisID
            .MoveNext
        Wend
    End With
    Set rsDx = Nothing
End Sub

Private Sub btnDxDelete_Click()
    If IsNumeric(txtDxID1.Text) = False Then
        MsgBox "Please select a diagnosis to delete"
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtDxID1.Text) <> 0 Then
            temSql = "Select * from tblPatientDiagnosis where PatientDiagnosisID = " & Val(txtDxID1.Text)
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
    txtDxID.Text = Empty
    cmbDiagnosis.SetFocus
End Sub

Private Sub btnIxGraph_Click()
    Unload frmIxGraph
    frmIxGraph.Show
    frmIxGraph.ZOrder 0
End Sub

Private Sub btnObsAdd_Click()
    Dim MyControl As Control
    Dim temObsHistoryValueID  As Long
    Dim temObsHistoryID As Long
    Dim temSystemFindingID As Long
    Dim rstemOBS As New ADODB.Recordset
    
    If IsNumeric(txtObsEditID.Text) = True Then
        temObsHistoryID = Val(txtObsEditID.Text)
    Else
        With rstemOBS
            If .State = 1 Then .Close
            temSql = "Select * from tblObsHistory"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = FormPatient.ID
            .Update
            temObsHistoryID = !ObsHistoryID
            .Close
        End With
    End If
    

    For Each MyControl In Controls
        If Left(MyControl.Name, 6) = "cmbObs" Then
            temSystemFindingID = Val(Right(MyControl.Name, Len(MyControl.Name) - 6))
            
            With rstemOBS
                If .State = 1 Then .Close
                temSql = "Select * from tblObsHistoryValue where PatientID = " & FormPatient.ID & " AND Deleted = False AND SystemFindingID = " & temSystemFindingID & " AND ObsHistoryID = " & temObsHistoryID
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                
                Else
                    .AddNew
                    !PatientID = FormPatient.ID
                    !ObsHistoryID = temObsHistoryID
                    !SystemFindingID = temSystemFindingID
                    
                End If
                !ObsHistoryValue = MyControl.Text
                .Update
                .Close
            End With
        End If
    Next
    
    Call ClearObsValues
    Call FillObs
    
    
End Sub

Private Sub FillObs()
    Dim ObsSF() As String
    Dim ObsSFID() As Long
    Dim ObsSFCount As Long
    Dim temRow As Long
    Dim rsSF As New ADODB.Recordset
    Dim rsObsHx As New ADODB.Recordset
    Dim i As Integer
    
    With rsSF
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding where IsObsHistory = true and Deleted = False order by OrderNO"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            ObsSFCount = .RecordCount
            .MoveFirst
            ReDim ObsSF(ObsSFCount)
            ReDim ObsSFID(ObsSFCount)
            While .EOF = False
                ObsSF(i) = !SystemFinding
                ObsSFID(i) = !SystemFindingID
                .MoveNext
                i = i + 1
            Wend
        Else
            ObsSFCount = 0
            ReDim ObsSF(ObsSFCount)
            ReDim ObsSFID(ObsSFCount)
        End If
        .Close
    End With
    With gridObs
        .Clear
        .Rows = 1
        .Cols = ObsSFCount + 1
        .FixedCols = 0
        .Row = 0
        .TextMatrix(0, ObsSFCount) = "ID"
        .ColWidth(ObsSFCount) = 0
    End With
    For i = 0 To ObsSFCount - 1
        With gridObs
            .Col = i
            .Text = ObsSF(i)
        End With
    Next
    
    With rsObsHx
        If .State = 1 Then .Close
        temSql = "Select * from tblObsHistory where Deleted = False AND PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temRow = 1
            While .EOF = False
                gridObs.Rows = gridObs.Rows + 1
                For i = 0 To ObsSFCount - 1
                    If rsSF.State = 1 Then rsSF.Close
                    temSql = "Select * from tblObsHistoryValue where Deleted = False AND PatientID = " & FormPatient.ID & " AND SystemFindingID = " & ObsSFID(i) & " AND ObsHistoryID = " & !ObsHistoryID
                    rsSF.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    If rsSF.RecordCount > 0 Then
                        gridObs.TextMatrix(temRow, i) = rsSF!ObsHistoryValue
                    End If
                    rsSF.Close
                Next i
                gridObs.TextMatrix(temRow, ObsSFCount) = !ObsHistoryID
                .MoveNext
                temRow = temRow + 1
            Wend
        End If
        .Close
    End With
End Sub


Private Sub ClearObsValues()
    Dim MyControl As Control
    For Each MyControl In Controls
        If Left(MyControl.Name, 6) = "cmbObs" Then
            MyControl.Text = Empty
        End If
    Next
End Sub


Private Sub btnObsDelete_Click()
    If IsNumeric(txtObsDelID.Text) = False Then
        MsgBox "Please select an entry to delete"
        gridObs.SetFocus
        Exit Sub
    End If
    Dim rsObs As New ADODB.Recordset
    With rsObs
        If .State = 1 Then .Close
        temSql = "Select * from tblObsHistoryValue where Deleted = False AND ObsHistoryID = " & Val(txtObsDelID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedTime = Now
            .Update
            .MoveNext
        Wend
        If .State = 1 Then .Close
        temSql = "Select * from tblObsHistory where Deleted = False AND ObsHistoryID = " & Val(txtObsDelID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedTime = Now
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Call ClearObsValues
    Call FillObs
End Sub

Private Sub btnRxAdd_Click()
    Dim ThisItem As New clsItem
    Dim i As Integer
    Dim rsTem As New ADODB.Recordset
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Please select an Item"
        cmbItem.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtDose.Text) = False Then
        MsgBox "Please select a Dose"
        txtDose.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbDoseUnit.BoundText) = False Then
        MsgBox "Please select a Dose Unit"
        cmbDoseUnit.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbFrequencyUnit.BoundText) = False Then
        MsgBox "Please select a frequency"
        cmbFrequencyUnit.SetFocus
        Exit Sub
    End If
    If IsAllergic(FormPatient.ID, cmbItem.BoundText) = True Then
        i = MsgBox("The medicine you are going to prescribe May be allergic to this patient. Do you still want to Prescribe this medicine?", vbYesNo)
        If i <> vbYes Then
            frmPastHistory.SSTab1.Tab = 2
            Exit Sub
        End If
    End If
    ThisItem.ID = Val(cmbItem.BoundText)
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtAddID.Text) = True Then
            temSql = "Select * from tblPatientItem where PatientItemID = " & Val(txtAddID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        Else
            temSql = "Select * from tblPatientItem"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = FormPatient.ID
        End If
        !UserID = UserID
        !ItemID = Val(cmbItem.BoundText)
        !ItemCategoryID = Val(cmbCategory.BoundText)
        !Dose = txtDose.Text
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

        .Update
    End With
    Call FillItems
    Call ClearItemValues
    cmbItem.SetFocus

End Sub

Private Sub FillItems()
    Dim rsTem As New ADODB.Recordset
    With GridItem
        .Rows = 1
        .Cols = 3
        .Clear
        .Row = 0
        .Col = 0
        .Text = "ID"
        .Col = 1
        .Text = "How to prescribe"
        .Col = 2
        .Text = "Indoor Issue Quentity"
        .ColWidth(0) = 0
        .ColWidth(2) = 2500
        .ColWidth(1) = .Width - 2650
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientItem.PatientItemID, [tblItem.Item] & ' ' & [tblPatientItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblPatientItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] & ' (' & [tblPatientItem].[IndoorOutdoor] & ')' AS Prescreption, 'Issue ' & tblPatientItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM (((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblPatientItem ON tblIssueUnit.ItemUnitID = tblPatientItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblPatientItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblPatientItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblPatientItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblPatientItem.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblPatientItem.PatientID)=" & FormPatient.ID & ") AND ((tblPatientItem.Deleted)=False))"
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            GridItem.Rows = GridItem.Rows + 1
            GridItem.Row = GridItem.Rows - 1
            GridItem.Col = 0
            GridItem.Text = ![PatientItemID]
            GridItem.Col = 1
            GridItem.Text = !Prescreption
            GridItem.Col = 2
            GridItem.Text = !ToIssue
            .MoveNext
        Wend
    End With
End Sub

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
    chkIndoor.Value = 0
End Sub

Private Sub btnRxDelete_Click()
    Dim rsTem As New ADODB.Recordset
    If IsNumeric(txtDeletedID.Text) = False Then
        MsgBox "Please select an Item to delete"
        Exit Sub
    End If
    With rsTem
        If Val(txtDeletedID.Text) <> 0 Then
            temSql = "Select * from tblPatientItem where PatientItemID = " & Val(txtDeletedID.Text)
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
    Call FillItems
    cmbItem.SetFocus
End Sub

Private Sub btnWeightGraph_Click()
    Unload frmWeightGraph
    
    frmWeightGraph.Show
    frmWeightGraph.ZOrder 0
End Sub

Private Sub Form_Activate()
    Call FillObsHx
    If ObsFilledOnce = False Then
        Call FillObs
        ObsFilledOnce = True
    End If
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
        End If
    Next
End Sub

Private Sub Form_Load()
        
    FormPatient.ID = CurrentPatient.ID
    
' ************** Tem

    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(4) = False
    SSTab1.TabsPerRow = 4

' **************
    
    
    Call FillCombos
    Call SetColours
    Call FillDiagnosis
    Call FillFHDiagnosis
    Call FillAllergy
    Call FillItems
    
    Call FillFoodAllergy
    
    Call GetSettings
    
    FilledOnce = False
    ObsFilledOnce = False
    
End Sub

Private Sub GetSettings()
    On Error Resume Next
    SSTab1.Tab = Val(GetSetting(App.EXEName, Me.Name, SSTab1.Name, 1))
    GetCommonSettings Me
    
    
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, SSTab1.Name, SSTab1.Tab
    SaveCommonSettings Me
    

End Sub

Private Sub FillObsHx()
    
    If FilledOnce = True Then
        Exit Sub
    Else
        FilledOnce = True
    End If

    Dim rsTem As New ADODB.Recordset
    Dim rsSF As New ADODB.Recordset
    Dim rsValue As New ADODB.Recordset
    Dim rsVV As New ADODB.Recordset
    
    Dim i As Integer
    
    Dim MyLabel As Label
    Dim MyCombo As ComboBox
    
    Dim LabelX As Long
    Dim LabelWidth As Long
    Dim LabelHeight As Long
    
    Dim ComboX As Long
    Dim ComboWidth As Long
    Dim ComboHeight As Long
    
    Dim LabelComboY As Long
    Dim LabelComboWidth As Long
    Dim LabelComboHeight As Long
    
    If rsSF.State = 1 Then rsSF.Close
    temSql = "Select * from tblSystemFinding where IsObsHistory = true and Deleted = False order by OrderNO"
    rsSF.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                
                
    LabelX = 100
    LabelWidth = 1000
    LabelHeight = 360
    
    ComboX = 1400
    ComboWidth = 4000
    ComboHeight = 360
    
    LabelComboY = 400
    LabelComboWidth = 2400
    LabelComboHeight = 500
                
    While rsSF.EOF = False
        
        If LabelComboY + LabelComboHeight > frameObs.Height Then
            LabelComboY = 400
            LabelX = LabelX + LabelComboWidth
            ComboX = ComboX + LabelComboWidth
        End If
        
        Set MyLabel = Me.Controls.Add("VB.label", "lblObs" & rsSF!SystemFindingID, frameObs)

        MyLabel.Height = LabelHeight
        MyLabel.Width = LabelWidth
        MyLabel.Top = LabelComboY
        MyLabel.Left = LabelX
        MyLabel.Caption = rsSF!SystemFinding
        MyLabel.Visible = True
        
        Set MyCombo = Me.Controls.Add("VB.ComboBox", "cmbObs" & rsSF!SystemFindingID, frameObs)

        MyCombo.Width = ComboWidth
        MyCombo.Top = LabelComboY
        MyCombo.Left = ComboX
        MyCombo.Visible = True
        
        If rsValue.State = 1 Then rsValue.Close
        temSql = "Select * from tblSystemFindingValue where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False order by SystemFindingValue"
        rsValue.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While rsValue.EOF = False
            MyCombo.AddItem rsValue!SystemFindingValue
            If rsValue!IsSystemFindingDefault = True Then MyCombo.Text = rsValue!SystemFindingValue
            rsValue.MoveNext
        Wend
        rsValue.Close

        LabelComboY = LabelComboY + LabelComboHeight
        
        rsSF.MoveNext
    Wend
                
End Sub


Private Sub FillCombos()
    Dim Dx As New clsFillCombos
    Dx.FillAnyCombo cmbDiagnosis, "Diagnosis", True
    Dim Px As New clsFillCombos
    Px.FillAnyCombo cmbProcedure, "Procedure", True
    Dim Allergen As New clsFillCombos
    Allergen.FillAnyCombo cmbAllergy, "Item", True
    Dim HxDx As New clsFillCombos
    HxDx.FillAnyCombo cmbFHDiagnosis, "Diagnosis", True
    
    ' Change1
    Dim DoseU As New clsFillCombos
    DoseU.FillBoolCombo cmbDoseUnit, "ItemUnit", "ItemUnit", "IsDoseUnit", True
    ' Change1
    
    Dim FreqU As New clsFillCombos
    FreqU.FillBoolCombo cmbFrequencyUnit, "ItemUnit", "ItemUnit", "IsFrequencyUnit", True
    
    Dim Duration As New clsFillCombos
    Duration.FillBoolCombo cmbDurationUnit, "ItemUnit", "ItemUnit", "IsDurationUnit", True
    
    Dim IssueUnit As New clsFillCombos
    IssueUnit.FillBoolCombo cmbIssueUnit, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "Category", "Category", "IsItemCategory", True
    
    Dim Item As New clsFillCombos
    Item.FillAnyCombo cmbItem, "Item", True
    
    Dim FI As New clsFillCombos
    FI.FillAnyCombo cmbFoodItem, "FoodItem", True
    
End Sub

Private Sub cmbItem_Change()
    Dim TemItem As New clsItem
    TemItem.ID = Val(cmbItem.BoundText)
    If TemItem.IsItem = True Then
        cmbCategory.Enabled = False
        cmbCategory.BoundText = TemItem.CategoryID
    Else
        cmbCategory.Enabled = True
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub gridDiagnosis_Click()
    Dim temGridRow As Long
    With gridDiagnosis
        temGridRow = .Row
        .Col = 0
        txtDxID1.Text = .Text
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub gridDiagnosis_DblClick()
    Dim temGridRow As Long
    With gridDiagnosis
        temGridRow = .Row
        .Col = 0
        txtDxID.Text = .Text
        .ColSel = .Cols - 1
        .Col = 3
        cmbDiagnosis.BoundText = Val(.Text)
        .Col = 2
        txtDiagnosis.Text = .Text
    End With
End Sub

Private Sub FillAllergy()
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientAllergy.PatientAllergyID, tblPatientAllergy.Comments, tblItem.Item, tblItem.ItemID FROM tblItem RIGHT JOIN tblPatientAllergy ON tblItem.ItemID = tblPatientAllergy.ItemID where tblPatientAllergy.ItemID <> 0 AND tblPatientAllergy.Deleted = False AND tblPatientAllergy.PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridAllergy.Clear
        gridAllergy.Rows = 1
        gridAllergy.Cols = 4
        gridAllergy.ColWidth(0) = 0
        gridAllergy.ColWidth(1) = 3000
        gridAllergy.ColWidth(2) = gridAllergy.Width - 3000 - 150
        gridAllergy.ColWidth(3) = 0
        gridAllergy.Row = 0
        gridAllergy.Col = 1
        gridAllergy.Text = "Allergy"
        gridAllergy.Col = 2
        gridAllergy.Text = "Details"
        
        While .EOF = False
            gridAllergy.Rows = gridAllergy.Rows + 1
            gridAllergy.Row = gridAllergy.Rows - 1
            gridAllergy.Col = 0
            gridAllergy.Text = !PatientAllergyID
            gridAllergy.Col = 1
            If IsNull(!Item) = False Then
                gridAllergy.Text = !Item
            End If
            gridAllergy.Col = 2
            gridAllergy.Text = Format(!Comments, "")
            gridAllergy.Col = 3
            If IsNull(!ItemID) = False Then
                gridAllergy.Text = !ItemID
            End If
            .MoveNext
        Wend
    End With
    Set rsDx = Nothing
End Sub

Private Sub gridAllergy_Click()
    Dim temGridRow As Long
    With gridAllergy
        temGridRow = .Row
        .Col = 0
        txtAxID1.Text = .Text
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub gridAllergy_DblClick()
    Dim temGridRow As Long
    With gridAllergy
        temGridRow = .Row
        .Col = 0
        txtAxID.Text = .Text
        .ColSel = .Cols - 1
        .Col = 3
        cmbAllergy.BoundText = Val(.Text)
        .Col = 2
        txtAllergy.Text = .Text
    End With
End Sub

Private Sub btnAxAdd_Click()
    If IsNumeric(cmbAllergy.BoundText) = False Then
        MsgBox "Please select a Allergy"
        cmbAllergy.SetFocus
        Exit Sub
    End If
    Dim rsAx As New ADODB.Recordset
    With rsAx
        If Val(txtAxID.Text) <> 0 Then
            temSql = "Select * from tblPatientAllergy where PatientAllergyID = " & Val(txtAxID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !ItemID = Val(cmbAllergy.BoundText)
                !Comments = txtAllergy.Text
                .Update
            Else
                .AddNew
                !PatientID = FormPatient.ID
                !ItemID = Val(cmbAllergy.BoundText)
                !Comments = txtAllergy.Text
                .Update
            End If
        Else
            temSql = "Select * from tblPatientAllergy"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = FormPatient.ID
            !ItemID = Val(cmbAllergy.BoundText)
            !Comments = txtAllergy.Text
            .Update
        End If
        .Close
    End With
    Call FillAllergy
    cmbAllergy.Text = Empty
    txtAllergy.Text = Empty
    txtAxID.Text = Empty
    cmbAllergy.SetFocus
End Sub

Private Sub btnAxDelete_Click()
    If IsNumeric(txtAxID1.Text) = False Then
        MsgBox "Please select a Allergy to delete"
        Exit Sub
    End If
    Dim rsAx As New ADODB.Recordset
    With rsAx
        If Val(txtAxID1.Text) <> 0 Then
            temSql = "Select * from tblPatientAllergy where PatientAllergyID = " & Val(txtAxID1.Text)
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
    Call FillAllergy
    cmbAllergy.Text = Empty
    txtAllergy.Text = Empty
    txtAxID.Text = Empty
    cmbAllergy.SetFocus
End Sub






























Private Sub FillFoodAllergy()
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientAllergy.PatientAllergyID, tblPatientAllergy.Comments, tblFoodItem.FoodItem, tblFoodItem.FoodItemID FROM tblFoodItem RIGHT JOIN tblPatientAllergy ON tblFoodItem.FoodItemID = tblPatientAllergy.FoodItemID where tblPatientAllergy.FoodItemID <> 0 AND tblPatientAllergy.Deleted = False AND tblPatientAllergy.PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridFoodItem.Clear
        gridFoodItem.Rows = 1
        gridFoodItem.Cols = 4
        gridFoodItem.ColWidth(0) = 0
        gridFoodItem.ColWidth(1) = 3000
        gridFoodItem.ColWidth(2) = gridFoodItem.Width - 3000 - 150
        gridFoodItem.ColWidth(3) = 0
        gridFoodItem.Row = 0
        gridFoodItem.Col = 1
        gridFoodItem.Text = "Food Allergy"
        gridFoodItem.Col = 2
        gridFoodItem.Text = "Details"
        
        While .EOF = False
            gridFoodItem.Rows = gridFoodItem.Rows + 1
            gridFoodItem.Row = gridFoodItem.Rows - 1
            gridFoodItem.Col = 0
            gridFoodItem.Text = !PatientAllergyID
            gridFoodItem.Col = 1
            If IsNull(!FoodItem) = False Then
                gridFoodItem.Text = !FoodItem
            End If
            gridFoodItem.Col = 2
            gridFoodItem.Text = Format(!Comments, "")
            gridFoodItem.Col = 3
            If IsNull(!FoodItemID) = False Then
                gridFoodItem.Text = !FoodItemID
            End If
            .MoveNext
        Wend
    End With
    Set rsDx = Nothing
End Sub





Private Sub gridFoodItem_Click()
    Dim temGridRow As Long
    With gridFoodItem
        temGridRow = .Row
        .Col = 0
        txtFoodItemDeleteID.Text = .Text
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub gridFoodItem_DblClick()
    Dim temGridRow As Long
    With gridFoodItem
        temGridRow = .Row
        .Col = 0
        txtFoodItemAddID.Text = .Text
        .ColSel = .Cols - 1
        .Col = 3
        cmbFoodItem.BoundText = Val(.Text)
        .Col = 2
        txtFoodItem.Text = .Text
    End With
End Sub

Private Sub btnFoodItemAdd_Click()
    If IsNumeric(cmbFoodItem.BoundText) = False Then
        MsgBox "Please select a Food Item"
        cmbFoodItem.SetFocus
        Exit Sub
    End If
    Dim rsAx As New ADODB.Recordset
    With rsAx
        If Val(txtFoodItemAddID.Text) <> 0 Then
            temSql = "Select * from tblPatientAllergy where PatientAllergyID = " & Val(txtFoodItemAddID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !FoodItemID = Val(cmbFoodItem.BoundText)
                !Comments = txtFoodItem.Text
                .Update
            Else
                .AddNew
                !PatientID = FormPatient.ID
                !FoodItemID = Val(cmbFoodItem.BoundText)
                !Comments = txtFoodItem.Text
                .Update
            End If
        Else
            temSql = "Select * from tblPatientAllergy"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = FormPatient.ID
            !FoodItemID = Val(cmbFoodItem.BoundText)
            !Comments = txtFoodItem.Text
            .Update
        End If
        .Close
    End With
    Call FillFoodAllergy
    cmbFoodItem.Text = Empty
    txtFoodItem.Text = Empty
    txtFoodItemAddID.Text = Empty
    txtFoodItemDeleteID.Text = Empty
    cmbFoodItem.SetFocus
End Sub

Private Sub btnFoodItemDelete_Click()
    If IsNumeric(txtFoodItemDeleteID.Text) = False Then
        MsgBox "Please select a Food Item to delete"
        Exit Sub
    End If
    Dim rsAx As New ADODB.Recordset
    With rsAx
        If Val(txtFoodItemDeleteID.Text) <> 0 Then
            temSql = "Select * from tblPatientAllergy where PatientAllergyID = " & Val(txtFoodItemDeleteID.Text)
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
    Call FillFoodAllergy
    cmbFoodItem.Text = Empty
    txtFoodItem.Text = Empty
    txtFoodItemAddID.Text = Empty
    txtFoodItemDeleteID.Text = Empty
    cmbFoodItem.SetFocus
End Sub






















Private Sub GridItem_Click()
    Dim temRow As Integer
    With GridItem
        temRow = .Row
        txtDeletedID.Text = .TextMatrix(temRow, 0)
    End With
End Sub

Private Sub GridItem_DblClick()
    Dim temRow As Integer
    Dim EditID As Long
    Dim rsTem As New ADODB.Recordset
    With GridItem
        temRow = .Row
        EditID = .TextMatrix(temRow, 0)
    End With
    Call ClearItemValues
    txtAddID.Text = EditID
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientItem where PatientItemID = " & EditID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbItem.BoundText = !ItemID
            cmbCategory.BoundText = !ItemCategoryID
            cmbDoseUnit.BoundText = !DoseUnitID
            cmbIssueUnit.BoundText = !IssueUnitID
            cmbFrequencyUnit.BoundText = !FrequencyUnitID
            cmbDurationUnit.BoundText = !DurationUnitID
            If IsNull(!Dose) = False Then
                txtDose.Text = !Dose
            Else
                txtDose.Text = Empty
            End If
            txtDuration.Text = !Duration
            txtIssueQty.Text = !IssueQuentity
            If !Indoor = True Then
                chkIndoor.Value = 1
            Else
                chkIndoor.Value = 0
            End If
        End If
        .Close
    End With

End Sub

Private Sub gridObs_Click()
    Dim ObsHxID As Long
    With gridObs
        ObsHxID = Val(.TextMatrix(.Row, .Cols - 1))
    End With
    txtObsEditID.Text = Empty
    txtObsDelID.Text = ObsHxID
End Sub

Private Sub gridObs_DblClick()
    Dim ObsHxID As Long
    Dim SFID As Long
    Dim rsObs As New ADODB.Recordset
    Dim MyControl As Control
    
    With gridObs
        ObsHxID = Val(.TextMatrix(.Row, .Cols - 1))
    End With
    txtObsEditID.Text = ObsHxID
    txtObsDelID.Text = ObsHxID
    
    For Each MyControl In Controls
        
        If Left(MyControl.Name, 6) = "cmbObs" Then
            SFID = Val(Right(MyControl.Name, Len(MyControl.Name) - 6))
            With rsObs
                If .State = 1 Then .Close
                temSql = "Select * from tblObsHistoryValue where ObsHistoryID = " & ObsHxID & " AND SystemFindingID = " & SFID & " AND Deleted = False"
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    MyControl.Text = !ObsHistoryValue
                End If
                .Close
            End With
    
        End If
    
    Next
End Sub


Private Sub btnFHDxAdd_Click()
    If IsNumeric(cmbFHDiagnosis.BoundText) = False Then
        MsgBox "Please select a diagnosis"
        cmbDiagnosis.SetFocus
        Exit Sub
    End If
    Dim rsFHDx As New ADODB.Recordset
    With rsFHDx
        If Val(txtFHDxID.Text) <> 0 Then
            temSql = "Select * from tblFamilyDiagnosis where FamilyDiagnosisID = " & Val(txtFHDxID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !DiagnosisID = Val(cmbFHDiagnosis.BoundText)
                !Comments = txtFHDiagnosis.Text
                .Update
            Else
                .AddNew
                !PatientID = FormPatient.ID
                !DiagnosisID = Val(cmbFHDiagnosis.BoundText)
                !Comments = txtFHDiagnosis.Text
                .Update
            End If
        Else
            temSql = "Select * from tblFamilyDiagnosis"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = FormPatient.ID
            !DiagnosisID = Val(cmbFHDiagnosis.BoundText)
            !Comments = txtFHDiagnosis.Text
            .Update
        End If
        .Close
    End With
    Call FillFHDiagnosis
    cmbFHDiagnosis.Text = Empty
    txtFHDiagnosis.Text = Empty
    txtFHDxID.Text = Empty
    cmbFHDiagnosis.SetFocus
End Sub

Private Sub FillFHDiagnosis()
    Dim rsFHDx As New ADODB.Recordset
    With rsFHDx
        If .State = 1 Then .Close
        temSql = "SELECT tblFamilyDiagnosis.FamilyDiagnosisID, tblFamilyDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblFamilyDiagnosis ON tblDiagnosis.DiagnosisID = tblFamilyDiagnosis.DiagnosisID where tblFamilyDiagnosis.Deleted = False AND tblFamilyDiagnosis.PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridFHDiagnosis.Clear
        gridFHDiagnosis.Rows = 1
        gridFHDiagnosis.Cols = 4
        gridFHDiagnosis.ColWidth(0) = 0
        gridFHDiagnosis.ColWidth(1) = 3000
        gridFHDiagnosis.ColWidth(2) = gridFHDiagnosis.Width - 3000 - 150
        gridFHDiagnosis.ColWidth(3) = 0
        gridFHDiagnosis.Row = 0
        gridFHDiagnosis.Col = 1
        gridFHDiagnosis.Text = "Diagnosis"
        gridFHDiagnosis.Col = 2
        gridFHDiagnosis.Text = "Details"
        
        While .EOF = False
            gridFHDiagnosis.Rows = gridFHDiagnosis.Rows + 1
            gridFHDiagnosis.Row = gridFHDiagnosis.Rows - 1
            gridFHDiagnosis.Col = 0
            gridFHDiagnosis.Text = !FamilyDiagnosisID
            gridFHDiagnosis.Col = 1
            gridFHDiagnosis.Text = !Diagnosis
            gridFHDiagnosis.Col = 2
            gridFHDiagnosis.Text = Format(!Comments, "")
            gridFHDiagnosis.Col = 3
            gridFHDiagnosis.Text = !DiagnosisID
            .MoveNext
        Wend
    End With
    Set rsFHDx = Nothing
End Sub

Private Sub btnFHDxDelete_Click()
    If IsNumeric(txtFHDxID1.Text) = False Then
        MsgBox "Please select a diagnosis to delete"
        Exit Sub
    End If
    Dim rsFHDx As New ADODB.Recordset
    With rsFHDx
        If Val(txtFHDxID1.Text) <> 0 Then
            temSql = "Select * from tblFamilyDiagnosis where FamilyDiagnosisID = " & Val(txtFHDxID1.Text)
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
    Call FillFHDiagnosis
    cmbFHDiagnosis.Text = Empty
    txtFHDiagnosis.Text = Empty
    txtFHDxID.Text = Empty
    cmbFHDiagnosis.SetFocus
End Sub

Private Sub gridFHDiagnosis_Click()
    Dim temGridRow As Long
    With gridFHDiagnosis
        temGridRow = .Row
        .Col = 0
        txtFHDxID1.Text = .Text
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub gridFHDiagnosis_DblClick()
    Dim temGridRow As Long
    With gridFHDiagnosis
        temGridRow = .Row
        .Col = 0
        txtFHDxID.Text = .Text
        .ColSel = .Cols - 1
        .Col = 3
        cmbFHDiagnosis.BoundText = Val(.Text)
        .Col = 2
        txtFHDiagnosis.Text = .Text
    End With
End Sub
