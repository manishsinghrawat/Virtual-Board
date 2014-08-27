VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "display"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lv 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11880
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function display(preg() As Double, num As Integer, num1 As Integer)
For j = 1 To num1
lv.ColumnHeaders.Add j
Next
For i = 1 To num
lv.ListItems.Add , , preg(i - 1, 0)
For j = 1 To num1 - 1
lv.ListItems(i).SubItems(j) = preg(i - 1, j)
Next
Next

End Function
Function displayf(preg() As Integer, num As Integer, num1 As Integer)
For j = 1 To num1
lv.ColumnHeaders.Add j
Next
For i = 1 To num
lv.ListItems.Add , , preg(i - 1, 0)
For j = 1 To num1 - 1
lv.ListItems(i).SubItems(j) = preg(i - 1, j)
Next
Next

End Function
Function display1(preg() As Double, num As Integer, num1 As Integer)

lv.ColumnHeaders.Add 1

For i = 1 To num
lv.ListItems.Add , , preg(i - 1)
Next
End Function


