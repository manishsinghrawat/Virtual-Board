VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form bro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circuit Browser"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save and EXIT"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit Characteristic"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Component Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Component Characteristics"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Component Listing"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "bro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not lv.ListItems.Count = 0 Then
inp = InputBox("Input Characteristic of element", "Input Characteristic", "1")
lv.SelectedItem.SubItems(1) = inp
lv.SetFocus
End If

End Sub

Private Sub Command3_Click()
frmmain.ActiveForm.setitup
End Sub

Private Sub Form_Load()
lv.GridLines = True
lv.ColumnHeaders(1).Width = (lv.Width / 2) - 100
lv.ColumnHeaders(2).Width = (lv.Width / 2) - 100
frmmain.ActiveForm.retcall
End Sub

Function setit(creg() As Integer, wreg() As Integer)
Dim i As Integer
For i = 0 To frmmain.ActiveForm.getccom - 1

If IsNumeric(lv.ListItems(i + 1).SubItems(1)) = False Then
creg(5, i) = 0
Else
creg(5, i) = lv.ListItems(i + 1).SubItems(1)
End If
Next

Unload Me
End Function

Function inf(creg() As Integer, wreg() As Integer)
Dim i As Integer
For i = 0 To frmmain.ActiveForm.getccom - 1
lv.ListItems.Add , , frmmain.ActiveForm.calname(i)

If creg(5, i) = 0 Then
lv.ListItems(lv.ListItems.Count).SubItems(1) = "---"
Else
lv.ListItems(lv.ListItems.Count).SubItems(1) = creg(5, i)
End If
Next


End Function
