VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   "Virtual Board"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8790
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   6780
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   11959
      ButtonWidth     =   1270
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Resistor"
            Key             =   "r"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Battery"
            Key             =   "b"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wire"
            Key             =   "w"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Browse"
            Key             =   "d"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run"
            Key             =   "ru"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer

Private Sub MDIForm_Load()
num = 1
newdoc
End Sub

Function newdoc()
Dim newd As New frmdoc
Load newd
newd.Show
newd.Caption = "Board " + CStr(num)
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
ActiveForm.bcom
If Not Button.Index = 4 Then
mode = Button.Index
End If

If Button.Index = 4 Then
Load bro
bro.Show vbModal
End If
If Button.Index = 5 Then
frmmain.ActiveForm.runit
End If
End Sub
