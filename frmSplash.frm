VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6855
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1680
      Top             =   3000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7080
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   6855
      Left            =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":000C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   4
      Top             =   6240
      Width           =   6855
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version V 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   5040
      Width           =   3300
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LicenseTo All    (FREEWARE)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   315
      TabIndex        =   1
      Top             =   2640
      Width           =   3510
   End
   Begin VB.Label lblcp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   4440
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   120
      Picture         =   "frmSplash.frx":00BD
      Top             =   600
      Width           =   7035
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Click()
Load frmmain
frmmain.Show
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Form_Click
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblcp.Caption = "Author : " + App.CompanyName
    Timer1.Enabled = True
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
Form_Click
End Sub

Private Sub lblCompanyProduct_Click()
Form_Click
End Sub

Private Sub Label1_Click()
Form_Click
End Sub

Private Sub lblcp_Click()
Form_Click
End Sub

Private Sub lblLicenseTo_Click()
Form_Click
End Sub

Private Sub lblPlatform_Click()
Form_Click
End Sub

Private Sub lblProductName_Click()
Form_Click
End Sub

Private Sub lblVersion_Click()
Form_Click
End Sub

Private Sub lblWarning_Click()
Form_Click
End Sub

Private Sub Timer1_Timer()
Form_Click
End Sub
