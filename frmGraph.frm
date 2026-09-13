VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGraph 
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmGraph.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   12660
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   255
      Left            =   11400
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Appearance      =   3
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
   Begin VB.OLE OLE1 
      Height          =   8895
      Left            =   120
      SizeMode        =   2  'AutoSize
      TabIndex        =   1
      Top             =   120
      Width           =   12375
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New FileSystemObject
    Dim temPath As String
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GetCommonSettings Me
    
    
    temPath = FSys.GetParentFolderName(Database)
    If FSys.FileExists(GraphPathTx) = False Then
        Dim tr As Integer
        tr = MsgBox("Error")
        Exit Sub
    End If
    OLE1.OLETypeAllowed = 0
    OLE1.SizeMode = 1
    OLE1.CreateLink GraphPathTx
    OLE1.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
    OLE1.Top = 100
    OLE1.Left = 100
    OLE1.Width = Me.Width - 500
    OLE1.Height = Me.Height - 1000
    bttnClose.Top = Me.Height - (bttnClose.Height * 4)
    bttnClose.Left = Me.Width - (bttnClose.Width * 2)
    Exit Sub

eh:
    Me.Width = Screen.Width * 0.75
    Me.Height = Screen.Width * 0.75
End Sub
