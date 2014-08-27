Attribute VB_Name = "mains"
Public focus As Double
Public rcol As Double
Public bcol As Double
Public wcol As Double
Public gpt As Double
Public upt As Double
Public mode As Integer

Sub main()

focus = vbRed

rcol = vbBlack
bcol = vbBlack
wcol = vbRed
gpt = &H80000000
upt = vbBlue
mode = 1

Load frmSplash
frmSplash.Show

End Sub
