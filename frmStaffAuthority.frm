VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmStaffAuthority 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Authorities"
   ClientHeight    =   7650
   ClientLeft      =   2325
   ClientTop       =   1050
   ClientWidth     =   13740
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
   ScaleHeight     =   7650
   ScaleWidth      =   13740
   Begin btButtonEx.ButtonEx btnEnableMenus 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Enable or Disable Menus"
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
      Left            =   12240
      TabIndex        =   0
      Top             =   7200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnVisibleMenus 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Make Visible or Invisible Menus"
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
Attribute VB_Name = "frmStaffAuthority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEnableMenus_Click()
    frmAuthorityPrevilagesMenuEnable.Show
    frmAuthorityPrevilagesMenuEnable.ZOrder 0
End Sub

Private Sub btnVisibleMenus_Click()
    frmAuthorityPrevilagesMenuVisible.Show
    frmAuthorityPrevilagesMenuVisible.ZOrder 0
End Sub
