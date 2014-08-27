VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Pro 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4680
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lb 
      Alignment       =   2  'Center
      Caption         =   "0 % complete"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This process may take several minutes depending on the complexity of your circuit and it's design"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Currently solving Resolved Matrice"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function shows(percent As Double)
Pro.Value = percent
lb.Caption = CStr(percent) + " % complete "
Me.Refresh
End Function

