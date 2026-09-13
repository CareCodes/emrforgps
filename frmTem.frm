VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTem 
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   15075
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   1080
      TabIndex        =   8
      Top             =   360
      Width           =   13215
   End
   Begin VB.TextBox txtFHDiagnosis 
      Height          =   855
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   9135
   End
   Begin VB.TextBox txtFHDxID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFHDxID1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10800
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin btButtonEx.ButtonEx btnFHDxAdd 
      Height          =   375
      Left            =   10800
      TabIndex        =   2
      Top             =   1680
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
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridFHDiagnosis 
      Height          =   3615
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6376
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin btButtonEx.ButtonEx btnFHDxDelete 
      Height          =   375
      Left            =   10800
      TabIndex        =   6
      Top             =   2160
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Diagnosis"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

