VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNewIxFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Investigation Format"
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6720
      TabIndex        =   19
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
      TabIndex        =   20
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
      Left            =   7080
      TabIndex        =   63
      Top             =   8760
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
      Left            =   6360
      TabIndex        =   62
      Top             =   8760
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
      Left            =   6720
      TabIndex        =   61
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtIxItemID 
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
      Left            =   5880
      TabIndex        =   60
      Top             =   8760
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
      Left            =   5160
      TabIndex        =   59
      Top             =   8040
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
      Left            =   5160
      TabIndex        =   58
      Top             =   8400
      Visible         =   0   'False
      Width           =   375
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
      Left            =   5520
      TabIndex        =   57
      Top             =   8400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkLabel 
      Height          =   240
      Left            =   5880
      TabIndex        =   56
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkText 
      Height          =   240
      Left            =   6120
      TabIndex        =   55
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkLine 
      Height          =   240
      Left            =   6120
      TabIndex        =   54
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkRectangle 
      Height          =   240
      Left            =   6360
      TabIndex        =   53
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkCircle 
      Height          =   240
      Left            =   5880
      TabIndex        =   52
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
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
      Left            =   5160
      TabIndex        =   51
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
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
      Left            =   5520
      TabIndex        =   50
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
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
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkCalc 
         Height          =   240
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkValue 
         Height          =   240
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtItemForeColour 
         Height          =   330
         Left            =   3360
         TabIndex        =   49
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtItemBackColour 
         Height          =   360
         Left            =   2640
         TabIndex        =   48
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox chkFontUnderline 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2400
         TabIndex        =   47
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkFontStrikeThrough 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   46
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkFontItalic 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkFontBold 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtFontSize 
         Height          =   330
         Left            =   840
         TabIndex        =   43
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtFontName 
         Height          =   360
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frameEdit 
      Height          =   2415
      Left            =   120
      TabIndex        =   38
      Top             =   6720
      Width           =   4815
      Begin VB.OptionButton optJustified 
         Caption         =   "Justified"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton optCenter 
         Caption         =   "Center"
         Height          =   255
         Left            =   2200
         TabIndex        =   16
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optRight 
         Caption         =   "Right"
         Height          =   255
         Left            =   1160
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton optLeft 
         Caption         =   "Left"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CheckBox chkMultiline 
         Caption         =   "Multiline"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1815
      End
      Begin btButtonEx.ButtonEx btnBackColour 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
      Begin btButtonEx.ButtonEx btnFont 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   1680
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Frame frameLabel 
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   4575
         Begin VB.TextBox txtLabel 
            Height          =   360
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Text"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frameText 
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   4575
         Begin VB.ComboBox cmbText 
            Height          =   360
            Left            =   840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "Field"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame frameSelect 
      Height          =   6615
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtIxID 
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid GridItem 
         Height          =   5295
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   9340
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin btButtonEx.ButtonEx btnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   6120
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   3240
         TabIndex        =   5
         Top             =   6120
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   1680
         TabIndex        =   4
         Top             =   6120
         Width           =   1455
         _ExtentX        =   2566
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
         TabIndex        =   0
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
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   5160
      TabIndex        =   29
      Top             =   8160
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox chkTextFontUnderline 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2400
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkTextFontStrikeThrough 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkTextFontItalic 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkTextFontBold 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtTextFontSize 
         Height          =   330
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtTextFontName 
         Height          =   360
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   5160
      TabIndex        =   22
      Top             =   8040
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox chkLabelFontUnderline 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkLabelFontStrikeThrough 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkLabelFontItalic 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkLabelFontBold 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtLabelFontSize 
         Height          =   330
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtLabelFontName 
         Height          =   360
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox pbxItem 
      AutoRedraw      =   -1  'True
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
      TabIndex        =   18
      Top             =   240
      Width           =   5775
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9720
      TabIndex        =   21
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
   Begin EmrForGps.UCPopup ucpIxItem 
      Left            =   0
      Top             =   0
      _ExtentX        =   318
      _ExtentY        =   370
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
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
   End
End
Attribute VB_Name = "frmNewIxFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim i As Integer
    Dim IncreaseValue As Double
    Dim rsIxItem As New ADODB.Recordset
    Dim PrintFormatItems() As New clsPrintFormatItem
    Dim BorderPick As Integer
    Public FormIxID As Long
    
    Dim ctlShape As Control
    Dim temZOrder As Integer
    Dim FoundLeftBorder As Boolean
    Dim FoundRightBorder As Boolean
    Dim FoundTopBorder As Boolean
    Dim FoundBottomBorder As Boolean
    Dim FoundCentre As Boolean
    Dim TopID As Long
    Dim BotID As Long
    Dim LeftID As Long
    Dim RightID As Long
    Dim MoveClickPointX As Long
    Dim MoveClickPointY As Long
    
    Dim PictureBoxBorderLength As Long
    Dim SelectedSerialNo As Long

    Dim LineRightLow As Boolean
    
    Dim YMoreThanY2 As Boolean
    Dim XMoreThanX2 As Boolean

    Dim OneSelecetd As Boolean

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
    Call GetIxItems(0, 0, pbxItem1, Val(cmbIx.BoundText))
    Call DrawIxItems(pbxItem1, 0)
    pbxItem.Picture = pbxItem1.Image
End Sub

Private Sub GetIxItems(ExceptID As Long, OnlyID As Long, MyPicture As PictureBox, IxID As Long)
    Dim rstemIxItems As New ADODB.Recordset
    
    FoundLeftBorder = False
    FoundRightBorder = False
    FoundTopBorder = False
    FoundBottomBorder = False
    FoundCentre = False
    
    TopID = 0
    BotID = 0
    LeftID = 0
    RightID = 0
    
    MoveClickPointX = 0
    MoveClickPointY = 0
    PictureBoxBorderLength = 0
    SelectedSerialNo = 0
    
    LineRightLow = False
    YMoreThanY2 = False
    XMoreThanX2 = False
    OneSelecetd = False
    
    With rstemIxItems
        If .State = 1 Then .Close
        If ExceptID <> 0 Then
            temSQL = "Select * from tblIxItem  where Deleted = false AND IxID = " & IxID & " And IxItemID <> " & ExceptID
        ElseIf OnlyID <> 0 Then
            temSQL = "Select * from tblIxItem  where IxID = " & IxID & " And IxItemID = " & OnlyID
        Else
            temSQL = "Select * from tblIxItem  where IxID = " & IxID & " And Deleted = false"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        ReDim PrintFormatItems(.RecordCount)
        If .RecordCount > 0 Then
            i = 0
            While .EOF = False
                PrintFormatItems(i).X1 = !X1
                PrintFormatItems(i).X2 = !X2
                PrintFormatItems(i).Y1 = !Y1
                PrintFormatItems(i).Y2 = !Y2
                PrintFormatItems(i).ScaleX1 = MyPicture.ScaleWidth * !X1
                PrintFormatItems(i).ScaleX2 = MyPicture.ScaleWidth * !X2
                PrintFormatItems(i).ScaleY1 = MyPicture.ScaleHeight * !Y1
                PrintFormatItems(i).ScaleY2 = MyPicture.ScaleHeight * !Y2
                If IsNull(!CircleRadius) = False Then
                    PrintFormatItems(i).Radious = !CircleRadius
                    PrintFormatItems(i).ScaleRadious = MyPicture.ScaleWidth * !CircleRadius
                End If
                PrintFormatItems(i).TextAlignment = !TextAlignment
                If !IsLine = True Then
                    PrintFormatItems(i).IsLine = True
                ElseIf !IsRectangle = True Then
                    PrintFormatItems(i).IsRectangle = True
                ElseIf !IsCircle = True Then
                    PrintFormatItems(i).IsCircle = True
                ElseIf !IsLabel = True Then
                    PrintFormatItems(i).IsLabel = True
                ElseIf !IsValue = True Then
                    PrintFormatItems(i).IsValue = True
                ElseIf !IsCalc = True Then
                    PrintFormatItems(i).IsCalc = True
                ElseIf !IsFlag = True Then
                    PrintFormatItems(i).IsFlag = True
                End If
                
                If Not IsNull(!ItemText) Then
                    PrintFormatItems(i).ItemText = Format(!ItemText, "")
                End If
                If Not IsNull(!IxItem) Then
                    PrintFormatItems(i).ItemName = !IxItem
                End If
                PrintFormatItems(i).ID = !IxItemID
                PrintFormatItems(i).ZOrder = i
                PrintFormatItems(i).SerialNo = i
                .MoveNext
                i = i + 1
            Wend
        End If
        .Close
    End With
End Sub


Private Sub DrawIxItems(MyPicture As PictureBox, FocusControlOD As Long)
    SelectedSerialNo = 0
    For i = 0 To UBound(PrintFormatItems)
        With PrintFormatItems(i)
            If .IsLine = True Then
                MyPicture.Line (.ScaleX1, .ScaleY1)-(.ScaleX2, .ScaleY2), .ForeColour
            ElseIf .IsRectangle = True Then
                MyPicture.Line (.ScaleX1, .ScaleY1)-(.ScaleX2, .ScaleY2), .ForeColour, B
            ElseIf .IsCircle = True Then
                MyPicture.Circle (.ScaleX1, .ScaleY1), .ScaleRadious, .ForeColour
            ElseIf .IsLabel = True Or .IsCalc = True Or .IsValue = True Or .IsFlag = True Then
                MyPicture.Line (.ScaleX1 + PictureBoxBorderLength, .ScaleY1 + PictureBoxBorderLength)-(.ScaleX2 + PictureBoxBorderLength, .ScaleY2 + PictureBoxBorderLength), vbYellow, BF
                Select Case .TextAlignment
                    Case TextAlignment.CentreAlign:
                        MyPicture.CurrentX = ((.ScaleX2 + .ScaleX1) / 2) - (MyPicture.TextWidth(.ItemText) / 2)
                    Case TextAlignment.JustifiedAlign:
                        MyPicture.CurrentX = ((.ScaleX2 + .ScaleX1) / 2) - (MyPicture.TextWidth(.ItemText) / 2)
                    Case TextAlignment.LeftAlign:
                        MyPicture.CurrentX = .ScaleX1
                    Case TextAlignment.RightAlign:
                        MyPicture.CurrentX = .ScaleX2 - (MyPicture.TextWidth(.ItemText))
                End Select
                MyPicture.CurrentY = .ScaleY1
                MyPicture.ForeColor = .ForeColour
                MyPicture.Print .ItemText
            ElseIf .IsText = True Then
                MyPicture.Line (.ScaleX1, .ScaleY1)-(.ScaleX2, .ScaleY2), vbYellow, BF
                Select Case .TextAlignment
                    Case TextAlignment.CentreAlign:
                        MyPicture.CurrentX = ((.ScaleX2 + .ScaleX1) / 2) - (MyPicture.TextWidth(.ItemText) / 2)
                    Case TextAlignment.JustifiedAlign:
                        MyPicture.CurrentX = ((.ScaleX2 + .ScaleX1) / 2) - (MyPicture.TextWidth(.ItemText) / 2)
                    Case TextAlignment.LeftAlign:
                        MyPicture.CurrentX = .ScaleX1
                    Case TextAlignment.RightAlign:
                        MyPicture.CurrentX = .ScaleX2 - MyPicture.TextWidth(.ItemText)
                End Select
                MyPicture.CurrentY = .ScaleY1
                MyPicture.ForeColor = .ForeColour
                MyPicture.Print .ItemText
            End If
        End With
    Next
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
    Call DrawIxItems(pbxItem1, 0)
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
    Dim rstemIxItems As New ADODB.Recordset
    With rstemIxItems
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem  where IxItemID = " & temID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    pbxItem.Cls
    pbxItem1.Cls
    Call GetIxItems(0, 0, pbxItem1, Val(cmbIx.BoundText))
    Call DrawIxItems(pbxItem1, 0)
    pbxItem.Picture = pbxItem1.Image
End Sub

Private Sub btnEdit_Click()
    txtX1.Text = "0.00"
    txtX2.Text = "0.00"
    txtY1.Text = "0.00"
    txtY2.Text = "0.00"
    txtRadius.Text = "0.00"
    
    chkLabel.Value = 0
    chkText.Value = 0
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
    
'    pbxItem.Cls
'    pbxItem1.Cls
'
'    Call DrawGraphics(temID, 0, pbxItem1, Val(cmbIx.BoundText))
'
'    pbxItem.Picture = pbxItem1.Image
    
    Dim rstemIxItems As New ADODB.Recordset
    With rstemIxItems
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
            txtIxItemID.Text = !IxItemID
            chkCircle.Value = 0
            chkLabel.Value = 0
            chkLine.Value = 0
            chkRectangle.Value = 0
            chkText.Value = 0
            
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
            
'            If !IsLabel = True Then
'                txtLabel.Text = Format(!ItemText, "")
'                chkLabel.Value = 1
'                frameText.Visible = False
'                frameLabel.Visible = True
'            End If
'            If !IsText = True Then
'                cmbText.Text = Format(!ItemText, "")
'                chkText.Value = 1
'                frameText.Visible = True
'                frameLabel.Visible = False
'            End If
            
            txtLabel.Text = Format(!ItemText, "")
            
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
    
    Call GetIxItems(0, 0, pbxItem1, Val(cmbIx.BoundText))
    Call DrawIxItems(pbxItem1, 0)
    pbxItem.Picture = pbxItem1.Image
    
End Sub

Private Sub btnFont_Click()
    Dim MyDefaults As New clsReportDefault
    If chkText.Value = 1 Then
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

'Private Sub DrawChanged()
'    Dim TemX1 As Long
'    Dim TemY1 As Long
'    Dim TemX2 As Long
'    Dim TemY2 As Long
'    Dim temRadius As Long
'
'
'
'    TemX1 = pbxItem.ScaleWidth * txtX1.Text
'    TemX2 = pbxItem.ScaleWidth * txtX2.Text
'    TemY1 = pbxItem.ScaleHeight * txtY1.Text
'    TemY2 = pbxItem.ScaleHeight * txtY2.Text
'    temRadius = pbxItem.ScaleWidth * txtRadius.Text
'
'
'    If chkCircle.Value = 1 Then
'        pbxItem.Circle (TemX1, TemY1), temRadius
'    ElseIf chkLabel.Value = 1 Then
'        pbxItem.Line (TemX1, TemY1)-(TemX2, TemY2), vbYellow, BF
'        If optCenter.Value = True Then
'                pbxItem.CurrentX = ((TemX2 + TemX1) / 2) - (pbxItem.TextWidth(txtLabel.Text) / 2)
'        ElseIf optJustified.Value = True Then
'                pbxItem.CurrentX = ((TemX2 + TemX1) / 2) - (pbxItem.TextWidth(txtLabel.Text) / 2)
'        ElseIf optLeft.Value = True Then
'                pbxItem.CurrentX = TemX1
'        ElseIf optRight.Value = True Then
'                pbxItem.CurrentX = TemX2 - (pbxItem.TextWidth(txtLabel.Text))
'        End If
'        pbxItem.CurrentY = TemY1
'        pbxItem.Print txtLabel.Text
'    ElseIf chkLine.Value = 1 Then
'        pbxItem.Line (TemX1, TemY1)-(TemX2, TemY2)
'    ElseIf chkRectangle.Value = 1 Then
'        pbxItem.Line (TemX1, TemY1)-(TemX2, TemY2), , B
'    ElseIf chkText.Value = 1 Then
'        pbxItem.Line (TemX1, TemY1)-(TemX2, TemY2), vbYellow, BF
'        If optCenter.Value = True Then
'                pbxItem.CurrentX = ((TemX2 + TemX1) / 2) - (pbxItem.TextWidth(cmbText.Text) / 2)
'        ElseIf optJustified.Value = True Then
'                pbxItem.CurrentX = ((TemX2 + TemX1) / 2) - (pbxItem.TextWidth(cmbText.Text) / 2)
'        ElseIf optLeft.Value = True Then
'                pbxItem.CurrentX = TemX1
'        ElseIf optRight.Value = True Then
'                pbxItem.CurrentX = TemX2 - (pbxItem.TextWidth(cmbText.Text))
'        End If
'        pbxItem.CurrentY = TemY1
'        pbxItem.Print cmbText.Text
'    End If
'
'End Sub

Private Sub btnSave_Click()
    Dim MyReportDefault As New clsReportDefault
    With rsIxItem
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem where IxItemID = " & Val(txtIxItemID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        
        !IxItem = txtLabel.Text
        !ItemText = txtLabel.Text
        
        If chkCircle.Value = 1 Then
            !IsCircle = True
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
         ElseIf chkValue.Value = 1 Then
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
    Call GetIxItems(0, 0, pbxItem1, Val(cmbIx.BoundText))
    pbxItem.Cls
    pbxItem1.Cls
    Call DrawIxItems(pbxItem1, 0)
    pbxItem.Picture = pbxItem1.Image
    Call SelectMode
End Sub

Private Sub FillCombos()
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
End Sub

Private Sub cmbIx_Change()
    txtIxID.Text = Val(cmbIx.BoundText)
    Call FormatGrid
    Call FillGrid
    Call GetIxItems(0, 0, pbxItem1, Val(cmbIx.BoundText))
    pbxItem.Cls
    pbxItem1.Cls
    Call DrawIxItems(pbxItem1, 0)
    pbxItem.Picture = pbxItem1.Image
    Call SelectMode
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
    BorderPick = 20
    PictureBoxBorderLength = 50
    
    txtX1.Text = "0.00"
    txtX2.Text = "0.00"
    txtY1.Text = "0.00"
    txtY2.Text = "0.00"
    txtRadius.Text = "0.00"
    
    Set ctlShape = Me.Controls.Add("VB.PICTUREBOX", "myShape")
    With Me("myShape")
        .BorderStyle = 1
        .Visible = True
        .Left = 0
        .Top = 0
        .Width = 100
        .Height = 100
        .ZOrder 0
    End With
    
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
    
    btnFont.Enabled = False
    btnBackColour.Enabled = False
    btnForeColour.Enabled = False
    
    frameText.Visible = False
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
    
    btnFont.Enabled = True
    btnBackColour.Enabled = True
    btnForeColour.Enabled = True
    
    frameText.Visible = True
    frameLabel.Visible = True
    
End Sub

Private Sub FillGrid()
    With rsIxItem
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
                GridItem.TextMatrix(i, 1) = Format(!IxItem, "")
                .MoveNext
            Next i
        End If
        .Close
    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    
End Sub

Private Sub GridItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MouseAction(1) As String
    Dim FocusID As Long
    Dim ItemRow As Double
    Dim TopRow As Double
    
    ItemRow = Y \ GridItem.RowHeight(0)
    If ItemRow < 1 Then Exit Sub
    ItemRow = ItemRow + GridItem.TopRow - 1
    TopRow = GridItem.TopRow
    If ItemRow > GridItem.Rows - 1 Then Exit Sub
    FocusID = Val(GridItem.TextMatrix(ItemRow, 0))
    
    If Button = vbRightButton Then
        MouseAction(0) = "Move to Centre"
        MouseAction(1) = "Edit"
        ucpIxItem.Popup MouseAction()
        GridRxPopAction ucpIxItem.ReturnIndex, FocusID, ItemRow
    End If
End Sub

Private Sub GridRxPopAction(ReturnIndex, ID, ItemRow)
    Dim rsTem As New ADODB.Recordset
    Select Case ReturnIndex
        Case 0:
            With rsTem
                If .State = 1 Then .Close
                temSQL = "Select * from tblIxItem where IxItemID = " & ID
                .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !X1 = 0.4
                    !X2 = 0.6
                    !Y1 = 0.4
                    !Y2 = 0.6
                    .Update
                End If
                .Close
            End With
            Call GetIxItems(0, 0, pbxItem1, Val(cmbIx.BoundText))
            pbxItem.Cls
            pbxItem1.Cls
            Call DrawIxItems(pbxItem1, 0)
            pbxItem.Picture = pbxItem1.Image
        Case 1:
            GridItem.row = ItemRow
            btnEdit_Click
        Case Else:
    
    End Select
End Sub

Private Sub pbxItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MouseAction(7) As String
    Dim FocusID As Long
    If Button = vbRightButton Then
        MouseAction(0) = "Bring to Front"
        MouseAction(1) = "Bring to Front"
        MouseAction(2) = "-"
        MouseAction(3) = "Bring Froward"
        MouseAction(4) = "Bring Backword"
        MouseAction(5) = "-"
        MouseAction(6) = "Delete"
        MouseAction(7) = "Edit"
        ucpIxItem.Popup MouseAction()
        FocusID = GetSelectedSerialNo(X, Y)
'        RxPopAction ucpIxItem.ReturnIndex, FocusID
    ElseIf Button = vbLeftButton Then
        If SelectedSerialNo <> 0 Then
            If ctlShape Is Nothing Then Set ctlShape = Me.Controls.Add("VB.Shape", "myShape")
            With Me("myShape")
                .BorderStyle = 1
                .Visible = True
                .ZOrder 0
            End With
        Else
            Me("myShape").Visible = False
        End If
    End If
    
    With GridItem
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = PrintFormatItems(SelectedSerialNo).ID Then
                .row = i
                .col = .Cols - 1
                .ColSel = 0
            End If
        Next
    End With
End Sub

Private Function GetSelectedSerialNo(X As Single, Y As Single) As Long
    Dim temZOrder As Integer
    GetSelectedSerialNo = 0
    Dim temText As String
    temZOrder = UBound(PrintFormatItems)
    For i = 0 To UBound(PrintFormatItems)
        With PrintFormatItems(i)
            If .ScaleX1 < X And X < .ScaleX2 Then
                If .ScaleY1 < Y And Y < .ScaleY2 Then
                    If temZOrder > .ZOrder Then
                        GetSelectedSerialNo = .SerialNo
                        temText = .ItemName
                    End If
                End If
            End If
        End With
    Next
    MsgBox temText
End Function

Private Sub RxPopAction(RxPopUpReturnIndex As Long, FocusItemID As Long)
    Select Case RxPopUpReturnIndex
        Case 0:
        Case 1:
        Case 3:
        Case 4:
        Case 6:
        Case 7:
    End Select
End Sub

Private Sub pbxItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If OneSelecetd = True Then
            Me("myShape").Visible = True
        
            If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 Then
                If YMoreThanY2 = True Then
                    LineRightLow = Not LineRightLow
                    YMoreThanY2 = False
                End If
            ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                If YMoreThanY2 = True Then
                    LineRightLow = Not LineRightLow
                    YMoreThanY2 = False
                End If
            ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                If YMoreThanY2 = False Then
                    LineRightLow = Not LineRightLow
                    YMoreThanY2 = True
                End If
            End If
            
            If X < PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                If XMoreThanX2 = True Then
                    LineRightLow = Not LineRightLow
                    XMoreThanX2 = False
                End If
            
            ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                If XMoreThanX2 = True Then
                    LineRightLow = Not LineRightLow
                    XMoreThanX2 = False
                End If
            
            ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                If XMoreThanX2 = False Then
                    LineRightLow = Not LineRightLow
                    XMoreThanX2 = True
                End If
            
            End If
        
            If SelectedSerialNo >= UBound(PrintFormatItems) Or TopID >= UBound(PrintFormatItems) Or BotID >= UBound(PrintFormatItems) Or LeftID >= UBound(PrintFormatItems) Or RightID >= UBound(PrintFormatItems) Then
                
            
            Else
                
            
                If FoundLeftBorder = True And FoundTopBorder = True And (LeftID = TopID) And PrintFormatItems(LeftID).IsLine = False Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (PrintFormatItems(SelectedSerialNo).ScaleY1 - Y)
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (Y - PrintFormatItems(SelectedSerialNo).ScaleY1)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY2 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (PrintFormatItems(SelectedSerialNo).ScaleX1 - X)
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (X - PrintFormatItems(SelectedSerialNo).ScaleX1)
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX2 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (X - PrintFormatItems(SelectedSerialNo).ScaleX2)
                    End If
                ElseIf FoundLeftBorder = True And FoundBottomBorder = True And (LeftID = BotID) And PrintFormatItems(LeftID).IsLine = False Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = PrintFormatItems(SelectedSerialNo).ScaleY1 - Y
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (PrintFormatItems(SelectedSerialNo).ScaleY2 - Y)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (PrintFormatItems(SelectedSerialNo).ScaleX1 - X)
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (X - PrintFormatItems(SelectedSerialNo).ScaleX1)
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX2 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (X - PrintFormatItems(SelectedSerialNo).ScaleX2)
                    End If
                ElseIf FoundRightBorder = True And FoundBottomBorder = True And (RightID = BotID) And PrintFormatItems(RightID).IsLine = False Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = PrintFormatItems(SelectedSerialNo).ScaleY1 - Y
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (PrintFormatItems(SelectedSerialNo).ScaleY2 - Y)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX1 - X)
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (PrintFormatItems(SelectedSerialNo).ScaleX2 - X)
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (X - PrintFormatItems(SelectedSerialNo).ScaleX2)
                    End If
                ElseIf FoundRightBorder = True And FoundTopBorder = True And (RightID = TopID) And PrintFormatItems(RightID).IsLine = False Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (PrintFormatItems(SelectedSerialNo).ScaleY1 - Y)
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (Y - PrintFormatItems(SelectedSerialNo).ScaleY1)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY2 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX1 - X)
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (PrintFormatItems(SelectedSerialNo).ScaleX2 - X)
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (X - PrintFormatItems(SelectedSerialNo).ScaleX2)
                    End If
                ElseIf FoundLeftBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (PrintFormatItems(SelectedSerialNo).ScaleX1 - X)
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (X - PrintFormatItems(SelectedSerialNo).ScaleX1)
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX2 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (X - PrintFormatItems(SelectedSerialNo).ScaleX2)
                    End If
                    Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                    Me("myShape").Height = PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1
                ElseIf FoundRightBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX1 - X)
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (PrintFormatItems(SelectedSerialNo).ScaleX2 - X)
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (X - PrintFormatItems(SelectedSerialNo).ScaleX2)
                    End If
                    Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                    Me("myShape").Height = PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1
                ElseIf FoundTopBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (PrintFormatItems(SelectedSerialNo).ScaleY1 - Y)
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (Y - PrintFormatItems(SelectedSerialNo).ScaleY1)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY2 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                    Me("myShape").Width = PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1
                ElseIf FoundBottomBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = PrintFormatItems(SelectedSerialNo).ScaleY1 - Y
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (PrintFormatItems(SelectedSerialNo).ScaleY2 - Y)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                    Me("myShape").Width = PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1
                    
                ElseIf FoundCentre = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                    Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1)
                    Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1)
                    If MoveClickPointY < Y Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength + (Y - MoveClickPointY)
                    Else
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY1 + pbxItem.Top + PictureBoxBorderLength - (MoveClickPointY - Y)
                    End If
                    If MoveClickPointX < X Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength + (X - MoveClickPointX)
                    Else
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength - (MoveClickPointX - X)
                    End If
    
                ElseIf FoundRightBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = True Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (PrintFormatItems(SelectedSerialNo).ScaleY1 - Y)
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (Y - PrintFormatItems(SelectedSerialNo).ScaleY1)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY2 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX1 - X)
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (PrintFormatItems(SelectedSerialNo).ScaleX2 - X)
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX1 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = (PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (X - PrintFormatItems(SelectedSerialNo).ScaleX2)
                    End If
                ElseIf FoundLeftBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = True Then
                    If Y < PrintFormatItems(SelectedSerialNo).ScaleY1 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) + (PrintFormatItems(SelectedSerialNo).ScaleY1 - Y)
                    ElseIf Y < PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = pbxItem.Top + Y + PictureBoxBorderLength
                        Me("myShape").Height = (PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1) - (Y - PrintFormatItems(SelectedSerialNo).ScaleY1)
                    ElseIf Y >= PrintFormatItems(SelectedSerialNo).ScaleY2 Then
                        Me("myShape").Top = PrintFormatItems(SelectedSerialNo).ScaleY2 + pbxItem.Top + PictureBoxBorderLength
                        Me("myShape").Height = (Y - PrintFormatItems(SelectedSerialNo).ScaleY2)
                    End If
                    If X < PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = Abs((PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) + (PrintFormatItems(SelectedSerialNo).ScaleX1 - X))
                    ElseIf X < PrintFormatItems(SelectedSerialNo).ScaleX2 Then
                        Me("myShape").Left = X + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = Abs((PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1) - (X - PrintFormatItems(SelectedSerialNo).ScaleX1))
                    ElseIf X >= PrintFormatItems(SelectedSerialNo).ScaleX1 Then
                        Me("myShape").Left = PrintFormatItems(SelectedSerialNo).ScaleX2 + pbxItem.Left + PictureBoxBorderLength
                        Me("myShape").Width = Abs((X - PrintFormatItems(SelectedSerialNo).ScaleX2))
                    End If
                ElseIf PrintFormatItems(SelectedSerialNo).IsLine = True Then
                    Me("myShape").Height = Abs((PrintFormatItems(SelectedSerialNo).ScaleY2 - PrintFormatItems(SelectedSerialNo).ScaleY1))
                    Me("myShape").Width = Abs((PrintFormatItems(SelectedSerialNo).ScaleX2 - PrintFormatItems(SelectedSerialNo).ScaleX1))
                    Me("myShape").Top = pbxItem.Top + PictureBoxBorderLength + Y - (Me("myShape").Height / 2)
                    Me("myShape").Left = pbxItem.Left + PictureBoxBorderLength + X - (Me("myShape").Width / 2)
                End If
            End If
        Else
            Me("myShape").Visible = False
        End If
    ElseIf Button = vbRightButton Then
        Me("myShape").Visible = False
    Else
        OneSelecetd = False
        Me("myShape").Visible = False
        FoundLeftBorder = False
        FoundRightBorder = False
        FoundTopBorder = False
        FoundBottomBorder = False
        FoundCentre = False
        temZOrder = UBound(PrintFormatItems)
        For i = 0 To UBound(PrintFormatItems)
            With PrintFormatItems(i)
                If .ScaleX1 - BorderPick < X And X < .ScaleX1 + BorderPick Then
                    If .ScaleY1 < Y And Y < .ScaleY2 Then
                        If temZOrder >= .ZOrder Then
                            SelectedSerialNo = .SerialNo
                            LeftID = .SerialNo
                            OneSelecetd = True
                        End If
                        FoundLeftBorder = True
                    End If
                End If
            End With
        Next
        For i = 0 To UBound(PrintFormatItems)
            With PrintFormatItems(i)
                If .ScaleX2 - BorderPick < X And X < .ScaleX2 + BorderPick Then
                    If .ScaleY1 < Y And Y < .ScaleY2 Then
                        If temZOrder >= .ZOrder Then
                            SelectedSerialNo = .SerialNo
                            RightID = .SerialNo
                            OneSelecetd = True
                        End If
                        FoundRightBorder = True
                    End If
                End If
            End With
        Next
        For i = 0 To UBound(PrintFormatItems)
            With PrintFormatItems(i)
                If .ScaleX1 < X And X < .ScaleX2 Then
                    If .ScaleY1 - BorderPick < Y And Y < .ScaleY1 + BorderPick Then
                        If temZOrder >= .ZOrder Then
                            SelectedSerialNo = .SerialNo
                            TopID = .SerialNo
                            OneSelecetd = True
                        End If
                        FoundTopBorder = True
                    End If
                End If
            End With
        Next
        For i = 0 To UBound(PrintFormatItems)
            With PrintFormatItems(i)
                If .ScaleX1 < X And X < .ScaleX2 Then
                    If .ScaleY2 - BorderPick < Y And Y < .ScaleY2 + BorderPick Then
                        If temZOrder >= .ZOrder Then
                            SelectedSerialNo = .SerialNo
                            BotID = .SerialNo
                            OneSelecetd = True
                        End If
                        FoundBottomBorder = True
                    End If
                End If
            End With
        Next
        For i = 0 To UBound(PrintFormatItems)
            With PrintFormatItems(i)
                If .ScaleX1 + BorderPick < X And X < .ScaleX2 - BorderPick Then
                    If .ScaleY2 - BorderPick > Y And Y > .ScaleY1 + BorderPick Then
                        If temZOrder >= .ZOrder Then
                            SelectedSerialNo = .SerialNo
                            MoveClickPointX = X
                            MoveClickPointY = Y
                            OneSelecetd = True
                        End If
                        FoundCentre = True
                    End If
                End If
            End With
        Next
        
        If SelectedSerialNo >= UBound(PrintFormatItems) Or TopID >= UBound(PrintFormatItems) Or BotID >= UBound(PrintFormatItems) Or LeftID >= UBound(PrintFormatItems) Or RightID >= UBound(PrintFormatItems) Then
            
        
        Else
        
            
            If FoundLeftBorder = True And FoundTopBorder = True And (LeftID = TopID) And PrintFormatItems(LeftID).IsLine = False Then
                Me.MousePointer = 8
            ElseIf FoundLeftBorder = True And FoundBottomBorder = True And (LeftID = BotID) And PrintFormatItems(LeftID).IsLine = False Then
                Me.MousePointer = 6
            ElseIf FoundRightBorder = True And FoundBottomBorder = True And (RightID = BotID) And PrintFormatItems(RightID).IsLine = False Then
                Me.MousePointer = 8
            ElseIf FoundRightBorder = True And FoundTopBorder = True And (RightID = TopID) And PrintFormatItems(RightID).IsLine = False Then
                Me.MousePointer = 6
            ElseIf FoundLeftBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                Me.MousePointer = 9
            ElseIf FoundRightBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                Me.MousePointer = 9
            ElseIf FoundTopBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                Me.MousePointer = 7
            ElseIf FoundBottomBorder = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                Me.MousePointer = 7
            ElseIf FoundCentre = True And PrintFormatItems(SelectedSerialNo).IsLine = False Then
                Me.MousePointer = 5
            ElseIf (FoundRightBorder = True Or FoundLeftBorder = True) And PrintFormatItems(SelectedSerialNo).IsLine = True Then
                Me.MousePointer = 10
            ElseIf (FoundRightBorder = False And FoundLeftBorder = False) And PrintFormatItems(SelectedSerialNo).IsLine = True Then
                Me.MousePointer = 5
            Else
                Me.MousePointer = vbDefault
                SelectedSerialNo = 0
                OneSelecetd = False
            End If
        End If
    End If
    
End Sub

Private Sub pbxItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rsTem As New ADODB.Recordset
    
    With Me("myShape")
        If .Top = 10 And .Left = 10 And .Width = 10 And .Height = 10 Then Exit Sub
        .Visible = False
    End With
    
    
    If OneSelecetd = True Then
    
        With rsTem
            If .State = 1 Then .Close
            temSQL = "Select * from tblIxItem where IxItemID = " & PrintFormatItems(SelectedSerialNo).ID
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !X1 = (Me("myShape").Left - pbxItem.Left + (PictureBoxBorderLength * 1)) / pbxItem.Width
                !X2 = (Me("myShape").Left + Me("myShape").Width - pbxItem.Left + (PictureBoxBorderLength * 1)) / pbxItem.Width
                !Y1 = (Me("myShape").Top - pbxItem.Top - (PictureBoxBorderLength * 0)) / pbxItem.Height
                !Y2 = (Me("myShape").Top + Me("myShape").Height - pbxItem.Top - (PictureBoxBorderLength * 0)) / pbxItem.Height
                .Update
            End If
            .Close
        End With
        Call GetIxItems(0, 0, pbxItem1, Val(cmbIx.BoundText))
        pbxItem.Cls
        pbxItem1.Cls
        Call DrawIxItems(pbxItem1, 0)
        pbxItem.Picture = pbxItem1.Image
    End If
    With Me("myShape")
        If .Top = 10 And .Left = 10 And .Width = 10 And .Height = 10 Then Exit Sub
        .Visible = False
        .Top = 10
        .Left = 10
        .Width = 10
        .Height = 10
    End With
End Sub
