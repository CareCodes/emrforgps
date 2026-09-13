VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIxFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Format"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10980
   Begin VB.TextBox txtPrintItemID 
      Height          =   360
      Left            =   4920
      TabIndex        =   73
      Top             =   4575
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRow 
      Height          =   360
      Left            =   240
      TabIndex        =   72
      Top             =   5280
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   5160
      TabIndex        =   54
      Top             =   8040
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtItemForeColour 
         Height          =   330
         Left            =   3360
         TabIndex        =   62
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtItemBackColour 
         Height          =   360
         Left            =   2640
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox chkFontUnderline 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2400
         TabIndex        =   60
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkFontStrikeThrough 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   59
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkFontItalic 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   58
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkFontBold 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   57
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtFontSize 
         Height          =   330
         Left            =   840
         TabIndex        =   56
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtFontName 
         Height          =   360
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frameEdit 
      Height          =   3375
      Left            =   120
      TabIndex        =   24
      Top             =   5760
      Width           =   4815
      Begin VB.CheckBox chkMultiline 
         Caption         =   "Multiline"
         Height          =   240
         Left            =   120
         TabIndex        =   71
         Top             =   2760
         Width           =   1815
      End
      Begin VB.OptionButton optLeft 
         Caption         =   "Left"
         Height          =   255
         Left            =   1920
         TabIndex        =   70
         Top             =   2760
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optRight 
         Caption         =   "Right"
         Height          =   255
         Left            =   1920
         TabIndex        =   69
         Top             =   3000
         Width           =   855
      End
      Begin VB.OptionButton optCenter 
         Caption         =   "Center"
         Height          =   255
         Left            =   3000
         TabIndex        =   68
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton optJustified 
         Caption         =   "Justified"
         Height          =   255
         Left            =   3000
         TabIndex        =   67
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox chkCalc 
         Height          =   240
         Left            =   4200
         TabIndex        =   66
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin btButtonEx.ButtonEx btnBackColour 
         Height          =   375
         Left            =   3120
         TabIndex        =   63
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Back Colour"
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
      Begin VB.TextBox txtBackColour 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   50
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtForeColour 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   49
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkCircle 
         Height          =   240
         Left            =   3720
         TabIndex        =   48
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkRectangle 
         Height          =   240
         Left            =   4200
         TabIndex        =   47
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkLine 
         Height          =   240
         Left            =   3960
         TabIndex        =   46
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkValue 
         Height          =   240
         Left            =   3960
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkLabel 
         Height          =   240
         Left            =   3720
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   43
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtRadius 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtY1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   40
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtX2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtX1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Move"
         Height          =   1695
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1335
         Begin btButtonEx.ButtonEx btnMoveUp 
            Height          =   375
            Left            =   480
            TabIndex        =   29
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "5"
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
         Begin btButtonEx.ButtonEx btnMoveRight 
            Height          =   375
            Left            =   840
            TabIndex        =   30
            Top             =   720
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "4"
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
         Begin btButtonEx.ButtonEx btnMoveLeft 
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
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
         Begin btButtonEx.ButtonEx btnMoveDown 
            Height          =   375
            Left            =   480
            TabIndex        =   32
            Top             =   1080
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "6"
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
      End
      Begin VB.Frame Frame3 
         Caption         =   "Size"
         Height          =   1695
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   1335
         Begin btButtonEx.ButtonEx btnSizeUp 
            Height          =   375
            Left            =   480
            TabIndex        =   26
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "5"
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
         Begin btButtonEx.ButtonEx btnSizeDown 
            Height          =   375
            Left            =   480
            TabIndex        =   27
            Top             =   1080
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "5"
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
         Begin btButtonEx.ButtonEx btnSizeRight 
            Height          =   375
            Left            =   840
            TabIndex        =   51
            Top             =   720
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
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
         Begin btButtonEx.ButtonEx btnSizeLeft 
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Appearance      =   3
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
      End
      Begin btButtonEx.ButtonEx btnFont 
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Font"
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
      Begin btButtonEx.ButtonEx btnForeColour 
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Fore Colour"
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
      Begin VB.TextBox txtY2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame frameLabel 
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   1920
         Width           =   4455
         Begin VB.TextBox txtReportItemID 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2880
            TabIndex        =   53
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtLabel 
            Height          =   360
            Left            =   840
            TabIndex        =   36
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Text"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame frameSelect 
      Height          =   5655
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtIxID 
         Height          =   375
         Left            =   4200
         TabIndex        =   65
         Top             =   240
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid GridItem 
         Height          =   4335
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7646
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin btButtonEx.ButtonEx btnAdd 
         Height          =   375
         Left            =   1080
         TabIndex        =   20
         Top             =   5160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Add"
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
         Left            =   3480
         TabIndex        =   21
         Top             =   5160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Delete"
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
      Begin btButtonEx.ButtonEx btnEdit 
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   5160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Edit"
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
      Begin MSDataListLib.DataCombo cmbIx 
         Height          =   360
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
   End
   Begin VB.PictureBox pbxItem1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   5160
      ScaleHeight     =   8355
      ScaleWidth      =   5715
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   5160
      TabIndex        =   8
      Top             =   8160
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox chkTextFontUnderline 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkTextFontStrikeThrough 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkTextFontItalic 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkTextFontBold 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtTextFontSize 
         Height          =   330
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtTextFontName 
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   8040
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox chkLabelFontUnderline 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkLabelFontStrikeThrough 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkLabelFontItalic 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkLabelFontBold 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtLabelFontSize 
         Height          =   330
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtLabelFontName 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox pbxItem 
      AutoRedraw      =   -1  'True
      Height          =   8415
      Left            =   5160
      ScaleHeight     =   8355
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   8760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save"
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   8760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cancel"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9720
      TabIndex        =   23
      Top             =   8760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
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
End
Attribute VB_Name = "frmIxFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim i As Integer
    Dim IncreaseValue As Double
    Dim rsReportItem As New ADODB.Recordset
    Public FormIxID As Long
    
Private Sub btnAdd_Click()
    If IsNumeric(cmbIx.BoundText) = False Then
        MsgBox "Please select an investigation"
        cmbIx.SetFocus
        Exit Sub
    Else
        FormIxID = Val(cmbIx.BoundText)
    End If
    frmAddIxItemWizard.Show 1
    Call FormatGrid
    Call FillGrid
    pbxItem.Cls
    pbxItem1.Cls
    Call DrawGraphics(0, 0, pbxItem1)
    pbxItem.Picture = pbxItem1.Image
End Sub

Private Sub btnBackColour_Click()
    CommonDialog1.ShowColor
    txtItemBackColour.Text = CommonDialog1.color
End Sub

Private Sub btnCancel_Click()
    Call FormatGrid
    Call FillGrid
    pbxItem.Cls
    pbxItem1.Cls
    Call DrawGraphics(0, 0, pbxItem)
    Call SelectMode
End Sub



Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim temRow As Integer
    Dim temID As Long
    With GridItem
        temRow = .row
        temID = Val(.TextMatrix(temRow, 0))
        If temID = 0 Then Exit Sub
    End With
    Dim rstemReportItems As New ADODB.Recordset
    With rstemReportItems
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem  where IxItemID = " & temID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedDate = Date
            !DeletedTime = Time
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    pbxItem.Cls
    pbxItem1.Cls
    Call DrawGraphics(0, 0, pbxItem1)
    pbxItem.Picture = pbxItem1.Image
End Sub

Private Sub btnEdit_Click()
    
    txtX1.Text = "0.00"
    txtX2.Text = "0.00"
    txtY1.Text = "0.00"
    txtY2.Text = "0.00"
    txtRadius.Text = "0.00"
    
    chkLabel.Value = 0
    chkValue.Value = 0
    chkCalc.Value = 0
    chkLine.Value = 0
    chkCircle.Value = 0
    chkRectangle.Value = 0
    
    txtLabel.Text = Empty
    
    txtFontName.Text = Empty
    txtFontSize.Text = Empty
    chkFontBold.Value = 0
    chkFontItalic.Value = 0
    chkFontStrikeThrough.Value = 0
    chkFontUnderline.Value = 0

    txtItemBackColour.Text = Empty
    txtItemForeColour.Text = Empty
    
    Dim temRow As Integer
    Dim temID As Long
    With GridItem
        temRow = .row
        temID = Val(.TextMatrix(temRow, 0))
        If temID = 0 Then Exit Sub
    End With
    
    Call EditMode
    
    pbxItem.Cls
    pbxItem1.Cls
    
    Call DrawGraphics(temID, 0, pbxItem1)
    
    pbxItem.Picture = pbxItem1.Image
    
    Dim rstemReportItems As New ADODB.Recordset
    With rstemReportItems
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem  where IxItemID = " & temID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtX1.Text = !X1
            txtX2.Text = !X2
            txtY1.Text = !Y1
            txtY2.Text = !Y2
            txtRadius.Text = !CircleRadius
            txtForeColour.Text = !ForeColour
            txtBackColour.Text = !BackColour
            txtPrintItemID.Text = !IxItemID
            
            If IsNull(!FontName) = False Then txtFontName.Text = !FontName
            txtFontSize.Text = !FontSize
            If !FontBold = True Then
                chkFontBold.Value = 1
            Else
                chkFontBold.Value = 0
            End If
            If !FontItalic = True Then
                chkFontItalic.Value = 1
            Else
                chkFontItalic.Value = 1
            End If
            If !FontStrikeThrough = True Then
                chkFontStrikeThrough.Value = 1
            Else
                chkFontStrikeThrough.Value = 0
            End If
            If !FontUnderline = True Then
                chkFontUnderline.Value = 1
            Else
                chkFontUnderline.Value = 0
            End If
            
            txtItemBackColour.Text = !BackColour
            txtItemForeColour.Text = !ForeColour
            
            If !IsLabel = True Then
                txtLabel.Text = Format(!ItemText, "")
                chkLabel.Value = 1
            Else
                txtLabel.Text = !IxItem
            End If
            If !IsLine = True Then chkLine.Value = 1
            If !IsRectangle = True Then chkRectangle.Value = 1
            If !IsCircle = True Then chkCircle.Value = 1
            If !IsValue = True Then chkValue.Value = 1
            If !IsCalc = True Then chkCalc.Value = 1
            
            If !AllowMultiline = True Then
                chkMultiline.Value = 1
            Else
                chkMultiline.Value = 0
            End If
            
            Select Case !TextAlignment
                Case TextAlignment.CentreAlign: optCenter.Value = True
                Case TextAlignment.JustifiedAlign: optJustified.Value = True
                Case TextAlignment.LeftAlign: optLeft.Value = True
                Case TextAlignment.RightAlign: optRight.Value = True
            End Select
            
            
        End If
        .Close
    End With
    Call DrawChanged
End Sub

Private Sub btnFont_Click()
    Dim MyDefaults As New clsReportDefault
    If chkCalc.Value = 1 Or chkValue.Value = 1 Then
        CommonDialog1.FontName = MyDefaults.TextFontName
        CommonDialog1.FontSize = MyDefaults.TextFontSize
        CommonDialog1.FontBold = MyDefaults.TextFontBold
        CommonDialog1.FontItalic = MyDefaults.TextFontItalic
        CommonDialog1.FontUnderline = MyDefaults.TextFontUnderline
        CommonDialog1.FontStrikethru = MyDefaults.TextFontStrikeThrough
    ElseIf chkLabel.Value = 1 Then
        CommonDialog1.FontName = MyDefaults.LabelFontName
        CommonDialog1.FontSize = MyDefaults.LabelFontSize
        CommonDialog1.FontBold = MyDefaults.LabelFontBold
        CommonDialog1.FontItalic = MyDefaults.LabelFontItalic
        CommonDialog1.FontUnderline = MyDefaults.LabelFontUnderline
        CommonDialog1.FontStrikethru = MyDefaults.LabelFontStrikeThrough
    End If
    CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
    CommonDialog1.ShowFont
    If CommonDialog1.FontBold = True Then
        chkFontBold.Value = 1
    Else
        chkFontBold.Value = 0
    End If
    If CommonDialog1.FontItalic = True Then
        chkFontItalic.Value = 1
    Else
        chkFontItalic.Value = 0
    End If
    If CommonDialog1.FontUnderline = True Then
        chkFontUnderline.Value = 1
    Else
        chkFontUnderline.Value = 0
    End If
    If CommonDialog1.FontStrikethru = True Then
        chkFontStrikeThrough.Value = 1
    Else
        chkFontStrikeThrough.Value = 0
    End If
    txtFontName.Text = CommonDialog1.FontName
    txtFontSize.Text = CommonDialog1.FontSize
End Sub

Private Sub btnForeColour_Click()
    CommonDialog1.ShowColor
    txtItemForeColour.Text = CommonDialog1.color
End Sub

Private Sub btnMoveDown_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtY1.Text = CDbl(txtY1.Text) + IncreaseValue
    txtY2.Text = CDbl(txtY2.Text) + IncreaseValue
    DrawChanged

End Sub

Private Sub btnMoveLeft_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtX1.Text = CDbl(txtX1.Text) - IncreaseValue
    txtX2.Text = CDbl(txtX2.Text) - IncreaseValue
    DrawChanged

End Sub

Private Sub btnMoveRight_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtX1.Text = CDbl(txtX1.Text) + IncreaseValue
    txtX2.Text = CDbl(txtX2.Text) + IncreaseValue
    DrawChanged

End Sub

Private Sub btnMoveUp_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtY1.Text = CDbl(txtY1.Text) - IncreaseValue
    txtY2.Text = CDbl(txtY2.Text) - IncreaseValue
    DrawChanged
End Sub

Private Sub DrawChanged()
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long
    
    TemX1 = pbxItem.ScaleWidth * txtX1.Text
    TemX2 = pbxItem.ScaleWidth * txtX2.Text
    TemY1 = pbxItem.ScaleHeight * txtY1.Text
    TemY2 = pbxItem.ScaleHeight * txtY2.Text
    temRadius = pbxItem.ScaleWidth * txtRadius.Text
    
    
    If chkCircle.Value = 1 Then
        pbxItem.Circle (TemX1, TemY1), temRadius
    ElseIf chkLabel.Value = 1 Or chkValue.Value = 1 Or chkCalc.Value = 1 Then
        pbxItem.Line (TemX1, TemY1)-(TemX2, TemY2), vbYellow, BF
        If optCenter.Value = True Then
                pbxItem.CurrentX = ((TemX2 + TemX1) / 2) - (pbxItem.TextWidth(txtLabel.Text) / 2)
        ElseIf optJustified.Value = True Then
                pbxItem.CurrentX = ((TemX2 + TemX1) / 2) - (pbxItem.TextWidth(txtLabel.Text) / 2)
        ElseIf optLeft.Value = True Then
                pbxItem.CurrentX = TemX1
        ElseIf optRight.Value = True Then
                pbxItem.CurrentX = TemX2 - (pbxItem.TextWidth(txtLabel.Text))
        End If
        pbxItem.CurrentY = TemY1
        pbxItem.Print txtLabel.Text
    ElseIf chkLine.Value = 1 Then
        pbxItem.Line (TemX1, TemY1)-(TemX2, TemY2)
    ElseIf chkRectangle.Value = 1 Then
        pbxItem.Line (TemX1, TemY1)-(TemX2, TemY2), , B
    End If
    
End Sub

Private Sub btnSave_Click()
    Dim MyReportDefault As New clsReportDefault
    With rsReportItem
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem where IxItemID = " & Val(txtPrintItemID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        
        !IxItem = txtLabel.Text
        
        If chkCircle.Value = 1 Then
            !IsCircle = True
            !CircleRadius = Abs(CDbl(txtRadius.Text))
        ElseIf chkLabel.Value = 1 Then
            !IsLabel = True
            !ItemText = txtLabel.Text
        ElseIf chkLine.Value = 1 Then
            !IsLine = True
        ElseIf chkRectangle.Value = 1 Then
            !IsRectangle = True
        ElseIf chkValue.Value = 1 Then
            !IsValue = True
        ElseIf chkCalc.Value = True Then
            !IsCalc = True
        End If
        
        !X1 = CDbl(txtX1.Text)
        !X2 = CDbl(txtX2.Text)
        !Y1 = CDbl(txtY1.Text)
        !Y2 = CDbl(txtY2.Text)
        !CircleRadius = CDbl(txtRadius.Text)
        
        If chkLabel.Value = 1 Then
            
            If Trim(txtFontName.Text) = "" Then
                !FontName = MyReportDefault.LabelFontName
            Else
                !FontName = txtFontName.Text
            End If
            If Val(txtFontSize.Text) = 0 Then
                !FontSize = MyReportDefault.LabelFontSize
            Else
                !FontSize = Val(txtFontSize.Text)
            End If
            !FontBold = chkFontBold.Value
            !FontItalic = chkFontItalic.Value
            !FontStrikeThrough = chkFontStrikeThrough.Value
            !FontUnderline = chkFontUnderline.Value
            If Trim(txtItemBackColour.Text) = "" Then
                !BackColour = MyReportDefault.LabelBackColour
            Else
                !BackColour = Val(txtItemBackColour.Text)
            End If
            If Trim(txtItemForeColour.Text) = "" Then
                !ForeColour = MyReportDefault.LabelForeColour
            Else
                !ForeColour = Val(txtItemForeColour.Text)
            End If
         ElseIf chkValue.Value = 1 Or chkCalc.Value = 1 Then
            If Trim(txtFontName.Text) = "" Then
                !FontName = MyReportDefault.TextFontName
            Else
                !FontName = txtFontName.Text
            End If
            If Val(txtFontSize.Text) = 0 Then
                !FontSize = MyReportDefault.TextFontSize
            Else
                !FontSize = Val(txtFontSize.Text)
            End If
            !FontBold = chkFontBold.Value
            !FontItalic = chkFontItalic.Value
            !FontStrikeThrough = chkFontStrikeThrough.Value
            !FontUnderline = chkFontUnderline.Value
            If Trim(txtItemBackColour.Text) = "" Then
                !BackColour = MyReportDefault.TextBackColour
            Else
                !BackColour = Val(txtItemBackColour.Text)
            End If
            If Trim(txtItemForeColour.Text) = "" Then
                !ForeColour = MyReportDefault.TextForeColour
            Else
                !ForeColour = Val(txtItemForeColour.Text)
            End If
        Else
            If Trim(txtItemBackColour.Text) = "" Then
                !BackColour = vbBlack
            Else
                !BackColour = Val(txtItemBackColour.Text)
            End If
            If Trim(txtItemForeColour.Text) = "" Then
                !ForeColour = vbWhite
            Else
                !ForeColour = Val(txtItemForeColour.Text)
            End If
        End If
        
        If chkMultiline.Value = 1 Then
            !AllowMultiline = True
        Else
            !AllowMultiline = False
        End If
        
        If optCenter.Value = True Then
            !TextAlignment = TextAlignment.CentreAlign
        ElseIf optJustified.Value = True Then
            !TextAlignment = TextAlignment.JustifiedAlign
        ElseIf optLeft.Value = True Then
            !TextAlignment = TextAlignment.LeftAlign
        ElseIf optRight.Value = True Then
            !TextAlignment = TextAlignment.RightAlign
        End If
        
        .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    pbxItem.Cls
    pbxItem1.Cls
    Call DrawGraphics(0, 0, pbxItem1)
    pbxItem.Picture = pbxItem1.Image
    Call SelectMode
End Sub

Private Sub btnSizeDown_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtY2.Text = CDbl(txtY2.Text) - IncreaseValue
    txtRadius.Text = CDbl(txtRadius.Text) - IncreaseValue
    DrawChanged
End Sub

Private Sub btnSizeLeft_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtX1.Text = CDbl(txtX1.Text) - IncreaseValue
    txtRadius.Text = CDbl(txtRadius.Text) + IncreaseValue
    DrawChanged
End Sub

Private Sub btnSizeRight_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtX2.Text = CDbl(txtX2.Text) - IncreaseValue
    txtRadius.Text = CDbl(txtRadius.Text) - IncreaseValue
    DrawChanged
End Sub

Private Sub btnSizeUp_Click()
    pbxItem.Cls
    pbxItem.Picture = pbxItem1.Image
    txtY1.Text = CDbl(txtY1.Text) - IncreaseValue
    txtRadius.Text = CDbl(txtRadius.Text) + IncreaseValue
    DrawChanged
End Sub

Private Sub FillCombos()
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
End Sub

Private Sub cmbIx_Change()
    FormIxID = Val(cmbIx.BoundText)
    txtIxID.Text = Val(cmbIx.BoundText)
    Call FormatGrid
    Call FillGrid
    pbxItem.Cls
    pbxItem1.Cls
    Call DrawGraphics(0, 0, pbxItem1)
    pbxItem.Picture = pbxItem1.Image
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

Private Sub Form_Load()
    Call SetColours
    GetCommonSettings Me
    
    
    IncreaseValue = 0.01
    
    txtX1.Text = "0.00"
    txtX2.Text = "0.00"
    txtY1.Text = "0.00"
    txtY2.Text = "0.00"
    txtRadius.Text = "0.00"
    
    Call FillCombos
    Call SelectMode
    Call FormatGrid
    
End Sub

Private Sub FormatGrid()
    With GridItem
        .Clear
        .Rows = 1
        .Cols = 2
        .ColWidth(0) = 0
        .ColWidth(1) = .Width - 150
    End With
End Sub

Private Sub SelectMode()

    frameSelect.Enabled = True
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
    
    
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    GridItem.Enabled = True
    
    btnMoveDown.Enabled = False
    btnMoveLeft.Enabled = False
    btnMoveRight.Enabled = False
    btnMoveUp.Enabled = False
    
    btnSizeDown.Enabled = False
    btnSizeLeft.Enabled = False
    btnSizeRight.Enabled = False
    btnSizeUp.Enabled = False
    
    btnFont.Enabled = False
    btnBackColour.Enabled = False
    btnForeColour.Enabled = False
    
    frameLabel.Visible = False

End Sub

Private Sub EditMode()
    frameSelect.Enabled = False
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    GridItem.Enabled = False
    
    btnMoveDown.Enabled = True
    btnMoveLeft.Enabled = True
    btnMoveRight.Enabled = True
    btnMoveUp.Enabled = True
    
    btnSizeDown.Enabled = True
    btnSizeLeft.Enabled = True
    btnSizeRight.Enabled = True
    btnSizeUp.Enabled = True
    
    btnFont.Enabled = True
    btnBackColour.Enabled = True
    btnForeColour.Enabled = True
    
    frameLabel.Visible = True
    
End Sub

Private Sub FillGrid()
    With rsReportItem
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem where Deleted = false And IxID = " & Val(cmbIx.BoundText) & " Order by IxItemID"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            Dim LabelCount As Long
            .MoveFirst
            LabelCount = .RecordCount
            Dim i As Integer
            GridItem.Rows = LabelCount + 1
            For i = 1 To .RecordCount
                GridItem.TextMatrix(i, 0) = !IxItemID
                GridItem.TextMatrix(i, 1) = !IxItem
                .MoveNext
            Next i
        End If
        .Close
    End With
    If Val(txtRow.Text) <> 0 Then
        On Error Resume Next
        GridItem.TopRow = Val(txtRow.Text)
    End If
End Sub







Private Sub ReportGraphics(ExceptID As Long, OnlyID As Long, MyPicture As PictureBox)
    Dim rstemReportItems As New ADODB.Recordset
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long
    
    
    With rstemReportItems
        If .State = 1 Then .Close
        If ExceptID <> 0 Then
            temSQL = "Select * from tblPrintItem  where Deleted = false And PrintItemCategoryID = 8 AND PrintItemID <> " & ExceptID
        ElseIf OnlyID <> 0 Then
            temSQL = "Select * from tblPrintItem  where  PrintItemCategoryID = 8 AND PrintItemID = " & OnlyID
        Else
            temSQL = "Select * from tblPrintItem  where  PrintItemCategoryID = 8 AND Deleted = false"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
        
            TemX1 = MyPicture.ScaleWidth * !X1
            TemX2 = MyPicture.ScaleWidth * !X2
            TemY1 = MyPicture.ScaleHeight * !Y1
            TemY2 = MyPicture.ScaleHeight * !Y2
            
            If IsNull(!CircleRadius) = False Then
                temRadius = MyPicture.ScaleWidth * !CircleRadius
            End If
            If !IsLine = True Then
                MyPicture.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour
            ElseIf !IsRectangle = True Then
                MyPicture.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour, B
            ElseIf !IsCircle = True Then
                MyPicture.Circle (TemX1, TemY1), temRadius, !ForeColour
            ElseIf !IsLabel = True Then
                MyPicture.Line (TemX1, TemY1)-(TemX2, TemY2), vbYellow, BF
                Select Case !TextAlignment
                    Case TextAlignment.CentreAlign:
                        MyPicture.CurrentX = ((TemX2 + TemX1) / 2) - (MyPicture.TextWidth(!ItemText) / 2)
                    Case TextAlignment.JustifiedAlign:
                        MyPicture.CurrentX = ((TemX2 + TemX1) / 2) - (MyPicture.TextWidth(!ItemText) / 2)
                    Case TextAlignment.LeftAlign:
                        MyPicture.CurrentX = TemX1
                    Case TextAlignment.RightAlign:
                        MyPicture.CurrentX = TemX2 - (MyPicture.TextWidth(!ItemText))
                End Select
                MyPicture.CurrentY = TemY1
                MyPicture.ForeColor = !ForeColour
                MyPicture.Print !ItemText
            ElseIf !IsText = True Then
                MyPicture.Line (TemX1, TemY1)-(TemX2, TemY2), vbYellow, BF
                Select Case !TextAlignment
                    Case TextAlignment.CentreAlign:
                        MyPicture.CurrentX = ((TemX2 + TemX1) / 2) - (MyPicture.TextWidth(!ItemText) / 2)
                    Case TextAlignment.JustifiedAlign:
                        MyPicture.CurrentX = ((TemX2 + TemX1) / 2) - (MyPicture.TextWidth(!ItemText) / 2)
                    Case TextAlignment.LeftAlign:
                        MyPicture.CurrentX = TemX1
                    Case TextAlignment.RightAlign:
                        MyPicture.CurrentX = TemX2 - MyPicture.TextWidth(!ItemText)
                End Select
                MyPicture.CurrentY = TemY1
                MyPicture.ForeColor = !ForeColour
                MyPicture.Print !ItemText
            End If
            .MoveNext
        Wend
        .Close
    End With

End Sub










Private Sub DrawGraphics(ExceptID As Long, OnlyID As Long, MyPicture As PictureBox)
    
    Call ReportGraphics(0, 0, MyPicture)
    
    Dim rstemReportItems As New ADODB.Recordset
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long
    
    
    With rstemReportItems
        If .State = 1 Then .Close
        If ExceptID <> 0 Then
            temSQL = "Select * from tblIxItem  where Deleted = false And IxItemID <> " & ExceptID & " AND IxID = " & Val(cmbIx.BoundText)
        ElseIf OnlyID <> 0 Then
            temSQL = "Select * from tblIxItem  where IxItemID = " & OnlyID
        Else
            temSQL = "Select * from tblIxItem  where Deleted = false  AND IxID = " & Val(cmbIx.BoundText)
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
        
            TemX1 = MyPicture.ScaleWidth * !X1
            TemX2 = MyPicture.ScaleWidth * !X2
            TemY1 = MyPicture.ScaleHeight * !Y1
            TemY2 = MyPicture.ScaleHeight * !Y2
            
            If IsNull(!CircleRadius) = False Then
                temRadius = MyPicture.ScaleWidth * !CircleRadius
            End If
            If !IsLine = True Then
                MyPicture.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour
            ElseIf !IsRectangle = True Then
                MyPicture.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour, B
            ElseIf !IsCircle = True Then
                MyPicture.Circle (TemX1, TemY1), temRadius, !ForeColour
            ElseIf !IsLabel = True Or !IsValue = True Or !IsCalc = True Then
                MyPicture.Line (TemX1, TemY1)-(TemX2, TemY2), vbYellow, BF
                Select Case !TextAlignment
                    Case TextAlignment.CentreAlign:
                        MyPicture.CurrentX = ((TemX2 + TemX1) / 2) - (MyPicture.TextWidth(!IxItem) / 2)
                    Case TextAlignment.JustifiedAlign:
                        MyPicture.CurrentX = ((TemX2 + TemX1) / 2) - (MyPicture.TextWidth(!IxItem) / 2)
                    Case TextAlignment.LeftAlign:
                        MyPicture.CurrentX = TemX1
                    Case TextAlignment.RightAlign:
                        MyPicture.CurrentX = TemX2 - (MyPicture.TextWidth(!IxItem))
                End Select
                MyPicture.CurrentY = TemY1
                MyPicture.ForeColor = !ForeColour
                MyPicture.Print !IxItem
            End If
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub GridItem_Click()
    txtRow.Text = GridItem.TopRow
End Sub

Private Sub txtIxID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbIx.BoundText = Val(txtIxID.Text)
    End If
End Sub
