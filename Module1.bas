Attribute VB_Name = "extra"
Function det(mat() As Double, order)
det = 0
If order = 1 Then
det = mat(0, 0)
Exit Function
End If

Dim ms(100, 100) As Double
For i = 1 To order

For a = 1 To order - 1
    For b = 0 To order - 1
        If b < i - 1 Then
            ms(a - 1, b) = mat(a, b)
        ElseIf b > i - 1 Then
            ms(a - 1, b - 1) = mat(a, b)
            End If
Next
Next

If Not mat(0, i - 1) = 0 Then
det = det + ((-1) ^ (i + 1)) * mat(0, i - 1) * det(ms, order - 1)
End If
Next

End Function

Function solve(mat() As Double, mat1() As Double, order As Integer, sfor As Integer, deter) As Variant
d = deter
Dim m1(100, 100) As Double


For y = 0 To order - 1
    For z = 0 To order - 1
    
        If z = sfor - 1 Then
           m1(y, z) = mat1(y)
        Else
           m1(y, z) = mat(y, z)
        End If
        
    Next
Next

solve = det(m1, order) / d

End Function


Function run(creg() As Integer, ccom As Integer, wreg() As Integer, cwir As Integer)

Dim preg(3, 1000) As Double
Dim jreg(1000) As Double
Dim cregn(4, 1000) As Double
ent = 0

For ctr = 0 To ccom - 1
match = False

For i = 0 To ent - 1
If preg(0, i) = creg(1, ctr) Then
If preg(1, i) = creg(2, ctr) Then
match = True
Exit For
End If
End If
Next

If match = False Then
preg(0, ent) = creg(1, ctr)
preg(1, ent) = creg(2, ctr)
preg(2, ent) = ent + 1
ent = ent + 1
End If

match = False
For i = 0 To ent - 1    'simplify it when program is completed
If preg(0, i) = creg(3, ctr) Then
If preg(1, i) = creg(4, ctr) Then
match = True
Exit For
End If
End If
Next

If match = False Then
preg(0, ent) = creg(3, ctr)
preg(1, ent) = creg(4, ctr)
preg(2, ent) = ent + 1
ent = ent + 1
End If

Next






junc = 2


edit = 1
While (edit)
edit = 0
For i = 0 To cwir - 1
t1 = -1 + search(preg, ent, wreg(0, i), wreg(1, i))
t2 = -1 + search(preg, ent, wreg(2, i), wreg(3, i))
If Not preg(2, t1) = preg(2, t2) Then
edit = 1
If preg(2, t1) < preg(2, t2) Then
preg(2, t2) = preg(2, t1)
Else
preg(2, t1) = preg(2, t2)
End If
End If

Next
Wend

nx = 1
junc = 1
found = 1
While (bigger(preg, ent, nx))
found = 0

For i = 0 To ent - 1
If preg(2, i) = nx Then
preg(3, i) = junc
found = 1
End If
Next

If found = 1 Then
junc = junc + 1
End If
nx = nx + 1
Wend
jn = junc - 1


GoTo endf
For i = 1 To ent - 1

For j = 0 To i - 1
If slink(preg(0, i), preg(1, i), preg(0, j), preg(1, j), wreg, cwir) = 1 Then
preg(3, i) = preg(3, j)
Exit For
End If
Next

If preg(3, i) = 0 Then
preg(3, i) = junc
junc = junc + 1
End If
Next
jn = junc - 1

endf:

For i = 0 To ccom - 1
cregn(0, i) = creg(0, i)
cregn(1, i) = getjunc(creg(1, i), creg(2, i), preg, ent)
cregn(2, i) = getjunc(creg(3, i), creg(4, i), preg, ent)
cregn(3, i) = creg(5, i)
cregn(4, i) = 0
Next

Dim prob1(1000, 1000) As Double
Dim prob2(1000) As Double
Dim order As Integer
order = jn + ccom - 1

'  problem matrice format
'   currents     --------   Potentials
'  Component equations
'   Kirchoff current law equations
'------------------------
'potential code for n   =ccom+n-1-1=ccom+n-2
'since one of junction is considered to be at zero potential
'current code  for n =n-1


'-------------------------------------------
'Procedure of building matrice from raw data
'-------------------------------------------
For i = 0 To ccom - 1
If cregn(0, i) = 1 Then
    If Not cregn(2, i) = 1 Then
      prob1(i, cregn(2, i) + ccom - 2) = 1
    End If
    If Not cregn(1, i) = 1 Then
    prob1(i, cregn(1, i) + ccom - 2) = -1
    End If
    
    prob1(i, i) = -cregn(3, i)
ElseIf cregn(0, i) = 2 Then
    If Not cregn(2, i) = 1 Then
      prob1(i, cregn(2, i) + ccom - 2) = 1
    End If
    If Not cregn(1, i) = 1 Then
    prob1(i, cregn(1, i) + ccom - 2) = -1
    End If
        prob2(i) = cregn(3, i)
End If
Next
'going current negetive junction at 1 -ve


For j = 1 To jn - 1

For k = 0 To ccom - 1
If cregn(1, k) = j Then
prob1(ccom - 1 + j, k) = -1
ElseIf cregn(2, k) = j Then
prob1(ccom - 1 + j, k) = 1
End If
Next
Next


Load Form2
Form2.Show
Dim v As Double
v = 0
Form2.Refresh

deter = det(prob1, order)
v = v + (1 / (order + 1)) * 100
Form2.shows v

For i = 0 To ccom - 1
cregn(4, i) = solve(prob1, prob2, order, i + 1, deter)
v = v + (1 / (order + 1)) * 100
Form2.shows v
Next

jreg(0) = 0
For j = 2 To jn
jreg(j - 1) = solve(prob1, prob2, order, ccom + j - 1, deter)
v = v + (1 / (order + 1)) * 100
Form2.shows v
Next


Dim f0 As New Form1
Load f0
f0.Show
f0.display preg, 4, CDbl(ent)

Dim f As New Form1
Load f
f.Show
f.display cregn, 5, ccom




Dim f1 As New Form1
Load f1
f1.Show
f1.display1 jreg, CDbl(jn), 1

Dim f2 As New Form1
Load f2
f2.Show
f2.display prob1, order, order


   
    

End Function

Function getjunc(x As Integer, y As Integer, preg() As Double, ent)
For i = 0 To ent - 1
If preg(0, i) = x And preg(1, i) = y Then
getjunc = preg(3, i)
Exit Function
End If
Next
End Function

Function slink(x1, y1, x2, y2, wreg() As Integer, cwir)
slink = -1
For i = 0 To cwir - 1
If wreg(0, i) = x1 Then
If wreg(1, i) = y1 Then
If wreg(2, i) = x2 Then
If wreg(3, i) = y2 Then
slink = 1
Exit Function
End If
End If
End If
End If

If wreg(0, i) = x2 Then
If wreg(1, i) = y2 Then
If wreg(2, i) = x1 Then
If wreg(3, i) = y1 Then
slink = 1
Exit Function
End If
End If
End If
End If
Next
End Function

Function bigger(preg() As Double, ent, n)
bigger = 0
For i = 0 To ent - 1
If preg(2, i) >= n Then
bigger = True
Exit Function
End If
Next

End Function

Function search(preg() As Double, ent, x, y)
search = -1
For i = 0 To ent - 1
If CInt(preg(0, i)) = CInt(x) And CInt(preg(1, i)) = CInt(y) Then
search = i + 1

Exit Function
End If
Next
If search = -1 Then
Err.Raise 1
End If
End Function

Function matrice()
Dim mat(100, 100) As Integer
mat(0, 0) = 1
mat(0, 1) = 5
mat(0, 2) = 67
mat(0, 3) = 32
mat(1, 0) = 564
mat(1, 1) = 56
mat(1, 2) = 1554
mat(1, 3) = 1454
mat(2, 0) = 15
mat(2, 1) = 165
mat(2, 2) = 154
mat(2, 3) = 154
mat(3, 0) = 156
mat(3, 1) = 133
mat(3, 2) = 143
mat(3, 3) = 134
Load Form1
Form1.Show
Form1.display mat, 4

End Function

