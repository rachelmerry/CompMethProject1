Attribute VB_Name = "Module1"
Option Explicit
Option Base 1
Public row1 As Integer, com1 As Integer, row2 As Integer, com2 As Integer

Sub Initialize()

'####UserForm that creates the dimensions of the matrix####

Matrix.Show

' #####reminds user to enter all data####
If Matrix.txtvect1row.Text = "" Then
    MsgBox "Please enter dimensions of matrix 1."
    Unload Matrix
    Matrix.Show
ElseIf Matrix.txtvect1com.Text = "" Then
    MsgBox "Please enter dimensions of matrix 1."
    Unload Matrix
    Matrix.Show
ElseIf Matrix.txtvect2row.Text = "" Then
    MsgBox "Please enter dimensions of matrix 2."
    Unload Matrix
    Matrix.Show
ElseIf Matrix.txtvect2com.Text = "" Then
    MsgBox "Please enter dimensions of matrix 2."
    Unload Matrix
    Matrix.Show
End If

'#### matrix dimensions, from Userform####
row1 = Matrix.txtvect1row.Text
com1 = Matrix.txtvect1com.Text
row2 = Matrix.txtvect2row.Text
com2 = Matrix.txtvect2com.Text

'#### values in matrix are inputted using a UserForm####
InputMatrix1.Show
InputMatrix2.Show


'#### This userform has 4 buttons that call the operation####
operation.Show

End Sub

Sub Addition()

'#### different size matrix- rows####
If row1 > row2 Then
    MsgBox "Matrices must be same dimension."
    Unload Matrix
    Matrix.Show
ElseIf row2 > row1 Then
    MsgBox "Matrices must be same dimension."
    Unload Matrix
    Matrix.Show
End If

'#### different size matrix- columns####
If com1 > com2 Then
    MsgBox "Matrices must be same dimension."
    Matrix.Show
    Unload Matrix
ElseIf com2 > com1 Then
    MsgBox "Matrices must be same dimension."
    Matrix.Show
    Unload Matrix
End If


Dim vect1() As Variant, vect2() As Variant, vect3() As Variant
ReDim vect1(row1, com1), vect2(row2, com2), vect3(row1, com1)


' ####addition loop####
Dim i As Double, j As Double
For i = 1 To row1
   For j = 1 To com1
        vect3(i, j) = vect1(i, j) + vect2(i, j)
    Next j
Next i

' ####Print to message box loop####
Dim k As Integer
Dim kk As Integer
Dim msg As String

For k = 1 To row1
    For kk = 1 To com1
    msg = msg & vect3(k, kk) & vbTab
    Next kk
    msg = msg & vbCrLf
Next k
    MsgBox msg
            
End Sub

Sub Subtraction()

'#### different size matrix- rows####
If row1 > row2 Then
    MsgBox "Matrices must be same dimension."
    Unload Matrix
    Matrix.Show
ElseIf row2 > row1 Then
    MsgBox "Matrices must be same dimension."
    Unload Matrix
    Matrix.Show
End If

'#### different size matrix- columns####
If com1 > com2 Then
    MsgBox "Matrices must be same dimension."
    Matrix.Show
    Unload Matrix
ElseIf com2 > com1 Then
    MsgBox "Matrices must be same dimension."
    Matrix.Show
    Unload Matrix
End If


Dim vect1() As Variant, vect2() As Variant, vect3() As Variant

ReDim vect1(row1, com1), vect2(row2, com2), vect3(row1, com1)

' ####Subtraction loop####
Dim i As Double, j As Double
For i = 1 To row1
   For j = 1 To com1
        vect3(i, j) = vect1(i, j) - vect2(i, j)
    Next j
Next i


' ####Print to message box loop####
Dim k As Integer
Dim kk As Integer
Dim msg As String

For k = 1 To row1
    For kk = 1 To com1
    msg = msg & vect3(k, kk) & vbTab
    Next kk
    msg = msg & vbCrLf
Next k
    MsgBox msg
    
End Sub

Sub Multiplication()
'#### incompatible matrix size####
If com1 > row2 Then
    MsgBox "Incompatible matrix sizes."
    Matrix.Show
    Unload Matrix
ElseIf row2 > com1 Then
    MsgBox "Incompatible matrix sizes."
    Matrix.Show
    Unload Matrix
End If

Dim vect1() As Variant, vect2() As Variant, vect3() As Variant


ReDim vect1(row1, com1), vect2(row2, com2), vect3(row1, com2)


Dim i As Integer, j As Integer, b As Integer, w() As Variant
ReDim w(row1, com2)

For i = 1 To row1
    For j = 1 To com2
        For b = 1 To row2
            w(i, j) = vect1(i, b) * vect2(b, j)
            vect3(i, j) = vect3(i, j) + w(i, j)
        Next b
    Next j
Next i


' ####Print to message box loop####
Dim k As Integer
Dim kk As Integer
Dim msg As String

For k = 1 To row1
    For kk = 1 To com2
        msg = msg & vect3(k, kk) & vbTab
    Next kk
    msg = msg & vbCrLf
Next k
    MsgBox msg
    
End Sub

Sub Division()

'#### incompatible matrix size####
If com1 > row2 Then
    MsgBox "Incompatible matrix sizes."
    Matrix.Show
    Unload Matrix
ElseIf row2 > com1 Then
    MsgBox "Incompatible matrix sizes."
    Matrix.Show
    Unload Matrix
End If

Dim vect1() As Variant, vect2() As Variant, vect3() As Variant, vect2inverse() As Variant


ReDim vect1(row1, com1), vect2(row2, com2), vect3(row1, com2), vect2inverse(row2, com2)


vect2inverse = WorksheetFunction.MInverse(vect2)

Dim i As Integer, j As Integer, b As Integer, w() As Variant
ReDim w(row1, com2)

For i = 1 To row1
    For j = 1 To com2
        For b = 1 To row2
            w(i, j) = vect1(i, b) * vect2inverse(b, j)
            vect3(i, j) = vect3(i, j) + w(i, j)
        Next b
    Next j
Next i


' ####Print to message box loop####
Dim k As Integer
Dim kk As Integer
Dim msg As String

For k = 1 To row1
    For kk = 1 To com2
        msg = msg & vect3(k, kk) & vbTab
    Next kk
    msg = msg & vbCrLf
Next k
    MsgBox msg
End Sub

