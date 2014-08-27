VERSION 5.00
Begin VB.Form frmdoc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sub Form"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   7485
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmdoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim build As Boolean
Dim spcx As Integer
Dim spcy As Integer
Dim rad As Integer
Dim spacing As Integer
Dim maxlen As Integer
Dim h As Integer
Dim hlab As Integer
Dim lastx As Integer
Dim lasty As Integer
Dim curx As Integer
Dim cury As Integer
 
Dim smlh As Integer
Dim larh As Integer
Dim batspc As Integer

Dim ccom As Integer
Dim cwir As Integer

'Resistor code 1
'battery code 2

Dim creg(5, 100) As Integer
'component type            pt 1 x,y            pt2 x,y                 characteristic
Dim wreg(3, 100) As Integer
'point1  x,y           point2  x,y
Dim preg(100, 100) As Integer
'       Junction ID
Function retcall()
bro.inf creg, wreg
End Function

Function setitup()
bro.setit creg, wreg
End Function

Function runit()
run creg, ccom, wreg, cwir

End Function

Private Sub Form_Load()

build = False
ccom = 0
cwir = 0
batspc = 200
larh = 200
smlh = 100
spacing = 100
maxlen = 100
h = 100
hlab = 400
spcx = 120 * 6
spcy = 120 * 6
rad = 35

End Sub

Function iboard()
FillColor = &H80000000
FillStyle = vbSolid

crx = spcx
While crx < Me.Width
cry = spcy
While cry < Me.Height

Circle (crx, cry), rad, FillColor
cry = cry + spcy
Wend
crx = crx + spcx
Wend
End Function

Function rboard()
Cls
iboard
n = 0
While Not wreg(0, n) = 0
Line (getx(wreg(0, n)), gety(wreg(1, n)))-(getx(wreg(2, n)), gety(wreg(3, n))), wcol
n = n + 1
Wend


n = 0
While Not creg(0, n) = 0
crlab (n)
Select Case creg(0, n)
Case 1
 cres creg(1, n), creg(2, n), creg(3, n), creg(4, n)
Case 2
  cbat creg(1, n), creg(2, n), creg(3, n), creg(4, n)
  
  'more here
End Select
n = n + 1
Wend

markpt
End Function

Function getx(ptx)
getx = ptx * spcx
End Function
Function gety(pty)
gety = pty * spcy
End Function

Function markpt()
For i = 0 To ccom
FillColor = upt
Circle (creg(1, i) * spcx, creg(2, i) * spcy), rad, FillColor
Circle (creg(3, i) * spcx, creg(4, i) * spcy), rad, FillColor
Next

For i = 0 To cwir
FillColor = upt
Circle (wreg(0, i) * spcx, wreg(1, i) * spcy), rad, FillColor
Circle (wreg(2, i) * spcx, wreg(3, i) * spcy), rad, FillColor
Next

End Function

Function crlab(num As Integer)
x1 = getx(creg(1, num))
y1 = gety(creg(2, num))
x2 = getx(creg(3, num))
y2 = gety(creg(4, num))
d = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
If d = 0 Then Exit Function
'label
xx1 = (x1 + x2) / 2
yy1 = (y1 + y2) / 2
x = (hlab * (y2 - y1)) / d + xx1
If y2 - y1 = 0 Then
y = yy1 - hlab
Else
y = ((x1 - x2) / (y2 - y1)) * (x - xx1) + yy1
End If
CurrentX = x
CurrentY = y
Print calname(num)

End Function

Function cres(xa1, ya1, xa2, ya2)
ForeColor = rcol
x1 = getx(xa1)
y1 = gety(ya1)
x2 = getx(xa2)
y2 = gety(ya2)
d = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
If d = 0 Then Exit Function

nd = d - 2 * spacing 'net length
m = Int(nd / maxlen) 'no of ups and downs
nm = nd / m 'wavelength

xx = (spacing * x2 + (d - spacing) * x1) / d
yy = (spacing * y2 + (d - spacing) * y1) / d

xx1 = (spacing * x1 + (d - spacing) * x2) / d
yy1 = (spacing * y1 + (d - spacing) * y2) / d

Line (x1, y1)-(xx, yy)
Line (x2, y2)-(xx1, yy1)

r = spacing
For i = 1 To m

r = r + nm / 2
xx1 = (r * x2 + (d - r) * x1) / d
yy1 = (r * y2 + (d - r) * y1) / d
x = (h * (y2 - y1)) / d + xx1

If y2 - y1 = 0 Then
y = yy1 - h
Else
y = ((x1 - x2) / (y2 - y1)) * (x - xx1) + yy1
End If

Line (xx, yy)-(x, y)

xx = x
yy = y

r = r + nm / 2
xx1 = (r * x2 + (d - r) * x1) / d
yy1 = (r * y2 + (d - r) * y1) / d
Line (xx, yy)-(xx1, yy1)


xx = xx1
yy = yy1
Next



End Function

Function cbat(xa1, ya1, xa2, ya2)
forecol = bcol
x1 = getx(xa1)
y1 = gety(ya1)
x2 = getx(xa2)
y2 = gety(ya2)
d = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
If d = 0 Then Exit Function
xx1 = ((d - batspc) / 2 * x2 + (d + batspc) / 2 * x1) / d
yy1 = ((d - batspc) / 2 * y2 + (d + batspc) / 2 * y1) / d
Line (x1, y1)-(xx1, yy1)


x = (larh * (y2 - y1)) / d + xx1
xr = -(larh * (y2 - y1)) / d + xx1
If y2 - y1 = 0 Then
y = yy1 - larh
yr = yy1 + larh
Else
y = ((x1 - x2) / (y2 - y1)) * (x - xx1) + yy1
yr = -((x1 - x2) / (y2 - y1)) * (x - xx1) + yy1
End If

Line (xr, yr)-(x, y)


xx1 = ((d - batspc) / 2 * x1 + (d + batspc) / 2 * x2) / d
yy1 = ((d - batspc) / 2 * y1 + (d + batspc) / 2 * y2) / d
Line (x2, y2)-(xx1, yy1)

x = (smlh * (y2 - y1)) / d + xx1
xr = -(smlh * (y2 - y1)) / d + xx1
If y2 - y1 = 0 Then
y = yy1 - smlh
yr = yy1 + smlh
Else
y = ((x1 - x2) / (y2 - y1)) * (x - xx1) + yy1
yr = -((x1 - x2) / (y2 - y1)) * (x - xx1) + yy1
End If

Line (xr, yr)-(x, y)

End Function

Function getptx(x)
If (x Mod spcx) < 0.5 * spcx Then
getptx = Int(x / spcx)
Else
getptx = Int(x / spcx) + 1
End If
End Function

Function getpty(y)
If (y Mod spcy) < 0.5 * spcy Then
getpty = Int(y / spcy)
Else
getpty = Int(y / spcy) + 1
End If
End Function



Public Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If getptx(x) < 1 Then Exit Sub
If getpty(y) < 1 Then Exit Sub

If Button = vbLeftButton Then

    If build = False Then
            build = True
            ibuild x, y   ' Initialize
            rboard
    Else
            build = False
            dbuild  'done
            rboard
    End If
  

ElseIf Button = vbRightButton Then
                    build = False
                   cbuild            'cancel
                   rboard
End If


End Sub
Function dbuild()
    If mode = 3 Then
                If wreg(0, cwir) = wreg(2, cwir) And wreg(1, cwir) = wreg(3, cwir) Then
                cbuild
                Else
                cwir = cwir + 1
                End If
    ElseIf mode = 1 Or mode = 2 Then
                If creg(1, ccom) = creg(3, ccom) And creg(2, ccom) = creg(4, ccom) Then
                cbuild
                Else
                 ccom = ccom + 1
                End If
            
    End If
End Function

Function ibuild(x, y)

            If mode = 3 Then
                wreg(0, cwir) = getptx(x)
                wreg(1, cwir) = getpty(y)
                wreg(2, cwir) = getptx(x)
                wreg(3, cwir) = getpty(y)
            ElseIf mode = 1 Or mode = 2 Then
                creg(0, ccom) = mode
                creg(1, ccom) = getptx(x)
                creg(2, ccom) = getpty(y)
                creg(3, ccom) = getptx(x)
                creg(4, ccom) = getpty(y)
            End If
End Function
Function cbuild()
            If mode = 1 Or mode = 2 Then
                    creg(0, ccom) = 0
                    creg(1, ccom) = 0
                    creg(2, ccom) = 0
                    creg(3, ccom) = 0
                    creg(4, ccom) = 0
            ElseIf mode = 3 Then
                    wreg(0, cwir) = 0
                    wreg(1, cwir) = 0
                    wreg(2, cwir) = 0
                    wreg(3, cwir) = 0
            End If
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If getptx(x) < 1 Then Exit Sub
If getpty(y) < 1 Then Exit Sub


If build = True Then
If mode = 3 Then
wreg(2, cwir) = getptx(x)
wreg(3, cwir) = getpty(y)
ElseIf mode = 1 Or mode = 2 Then
creg(3, ccom) = getptx(x)
creg(4, ccom) = getpty(y)
End If
End If

FillStyle = vbSolid


If Not (lastx = getptx(x) And lasty = getpty(y)) Then
If build = True Then
rboard
End If
FillColor = gpt
Circle (lastx * spcx, lasty * spcy), rad, FillColor
markpt
FillColor = focus
Circle (getptx(x) * spcx, getpty(y) * spcy), rad, FillColor
lastx = getptx(x)
lasty = getpty(y)

End If
End Sub

Function calname(num As Integer) As String
typ = creg(0, num)
ctr = 0
For i = 0 To num
If creg(0, i) = typ Then
ctr = ctr + 1
End If
Next
If typ = 1 Then
Char = "R"
ElseIf typ = 2 Then
Char = "B"
'more here
End If

calname = Char + CStr(ctr)
End Function
Function getccom()
getccom = ccom
End Function
 Function bcom()
build = False
dbuild
End Function

Private Sub Form_Resize()
rboard
End Sub
