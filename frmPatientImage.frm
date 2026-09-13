VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPatientImage 
   Caption         =   "Image"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctImage 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmPatientImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temPath As String
    Dim FSys As New FileSystemObject

Private Sub btnSave_Click()
    Dim i As Integer
    CommonDialog1.InitDir = temPath
    CommonDialog1.Filter = "BMP|*.bmp"
    CommonDialog1.filename = Me.Caption & ".bmp"
    CommonDialog1.ShowSave
    If FSys.FileExists(CommonDialog1.filename) = True Then
        i = MsgBox("The file name already exists. Do you want to replace it", vbYesNoCancel)
        If i = vbNo Then
            btnSave_Click
            Exit Sub
        ElseIf i = vbCancel Then
            Exit Sub
        End If
    End If
    
    If CommonDialog1.CancelError = False Then
        SavePicture pctImage.Image, CommonDialog1.filename
        SaveSetting App.EXEName, Me.Name, "Path", FSys.GetParentFolderName(CommonDialog1.filename)
        temPath = FSys.GetParentFolderName(CommonDialog1.filename)
    End If
End Sub

Private Sub Form_Load()
    Me.Width = GetSetting(App.EXEName, Me.Name, "Width", Me.Width)
    GetCommonSettings Me
    
    
    
    Me.WindowState = 0
    Me.Height = Me.Width * (frmPatientDetails.imgPatient.Picture.Width / frmPatientDetails.imgPatient.Picture.Height)
    pctImage.Top = 0
    pctImage.Left = 0
    pctImage.Height = Me.Height
    pctImage.Width = Me.Width
    pctImage.AutoSize = False
    pctImage.PaintPicture frmPatientDetails.imgPatient.Picture, 0, 0, Me.Width, Me.Height
    temPath = GetSetting(App.EXEName, Me.Name, "Path", App.Path)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Width", Me.Width
    SaveCommonSettings Me
    

End Sub

Private Sub Form_Resize()
    pctImage.Top = 0
    pctImage.Left = 0
    pctImage.Height = Me.Height
    pctImage.Width = Me.Width
    pctImage.AutoSize = False
    pctImage.PaintPicture frmPatientDetails.imgPatient.Picture, 0, 0, Me.Width, Me.Height
End Sub

Private Sub pctImage_DblClick()
    If Me.WindowState = 0 Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
    End If
End Sub
