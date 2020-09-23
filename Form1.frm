VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JawBreaker"
   ClientHeight    =   4095
   ClientLeft      =   5910
   ClientTop       =   3555
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   3015
   Begin VB.CommandButton Command1 
      Caption         =   "End Game"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3105
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim matrix(12, 13) As Long
Dim nb As Integer
Dim mb(20, 2) As Integer
Dim c As Long
Dim points
Dim xi, yi
'#########################################################################
Private Sub Init()
Dim X As Integer, Y As Integer
Picture1.Cls
For Y = 1 To 12
    For X = 1 To 11
       c = Int(Rnd * 5) + 1
       Select Case c
        Case 1
            matrix(X, Y) = vbBlue
        Case 2
            matrix(X, Y) = vbGreen
        Case 3
            matrix(X, Y) = vbMagenta
        Case 4
            matrix(X, Y) = vbYellow
        Case 5
            matrix(X, Y) = vbRed
       End Select
    Next
Next
For X = 1 To 11
    matrix(X, 13) = vbWhite
Next
For Y = 1 To 12
    matrix(12, Y) = vbWhite
Next
DrawCircles
End Sub
'#########################################################################
Private Sub Form_Load()
Randomize Time
Picture1.Scale (0, 120)-(110, 0)
Picture1.AutoRedraw = True
Init
End Sub
'#########################################################################
Private Sub Picture1_DblClick()
Dim i As Integer
If matrix(xi, yi) <> vbWhite Then
    If nb > 1 Then
        points = points + nb * (nb - 1)
        Label1.Caption = points
    End If
    For i = 1 To nb
        matrix(mb(i, 1), mb(i, 2)) = vbWhite
    Next
    PackCols
    PackFiles
    DrawCircles
End If
If Final Then
    MsgBox "END!" & vbCrLf & "POINTS: " & points
    Init
End If
End Sub
'#########################################################################
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DrawCircles
xi = X \ 10 + 1
yi = Y \ 10 + 1
If matrix(xi, yi) <> vbWhite Then
    nb = 0
    search xi, yi, matrix(xi, yi)
    PaintSearch
End If
End Sub
'#########################################################################
Private Sub search(XX, YY, c)
If XX > 0 And XX < 12 And YY > 0 And YY < 13 Then
    If Not exist(XX + 1, YY) Then
        If matrix(XX + 1, YY) = c Then
            nb = nb + 1
            mb(nb, 1) = XX + 1
            mb(nb, 2) = YY
    '        Debug.Print XX + 1, yy
            search XX + 1, YY, c
        End If
    End If
    If Not exist(XX - 1, YY) Then
        If matrix(XX - 1, YY) = c Then
            nb = nb + 1
            mb(nb, 1) = XX - 1
            mb(nb, 2) = YY
    '        Debug.Print XX - 1, yy
            search XX - 1, YY, c
        End If
    End If
    If Not exist(XX, YY + 1) Then
        If matrix(XX, YY + 1) = c Then
            nb = nb + 1
            mb(nb, 1) = XX
            mb(nb, 2) = YY + 1
    '        Debug.Print XX, yy + 1
            search XX, YY + 1, c
        End If
    End If
    If Not exist(XX, YY - 1) Then
        If matrix(XX, YY - 1) = c Then
            nb = nb + 1
            mb(nb, 1) = XX
            mb(nb, 2) = YY - 1
    '        Debug.Print XX, yy - 1
            search XX, YY - 1, c
        End If
    End If
End If
End Sub
'#########################################################################
Private Function exist(XX, YY) As Boolean
Dim i As Integer
    exist = False
    For i = 1 To nb
        If mb(i, 1) = XX And mb(i, 2) = YY Then
            exist = True
            Exit Function
        End If
    Next
End Function
'#########################################################################
Private Sub PaintSearch()
Dim i As Integer
For i = 1 To nb
    With Picture1
        .CurrentX = mb(i, 1) * 10 - 6
        .CurrentY = mb(i, 2) * 10 - 1
    End With
    Picture1.Print "X"
Next
End Sub
'#########################################################################
Private Sub DrawCircles()
Dim X As Integer, Y As Integer
Picture1.Cls
Picture1.FillStyle = vbFSSolid
For Y = 1 To 12
    For X = 1 To 11
        If matrix(X, Y) <> vbWhite Then
           Picture1.FillColor = matrix(X, Y)
           Picture1.Circle (X * 10 - 5, Y * 10 - 5), 5, vbBlack
        End If
    Next
Next
End Sub
'#########################################################################
Private Sub PackCols()
Dim i As Integer, j As Integer, X As Integer
For X = 1 To 11
    For i = 12 To 1 Step -1
        If matrix(X, i - 1) = vbWhite Then
            For j = i - 1 To 12
                matrix(X, j) = matrix(X, j + 1)
            Next
        End If
    Next
Next
End Sub
'#########################################################################
Private Sub PackFiles()
Dim i As Integer, j As Integer, X As Integer
For X = 11 To 1 Step -1
   If matrix(X - 1, 1) = vbWhite Then
        For i = X - 1 To 11
            For j = 1 To 12
                matrix(i, j) = matrix(i + 1, j)
            Next
        Next
   End If
Next
End Sub
'#########################################################################
Private Function Final() As Boolean
Dim i As Integer, j As Integer
Final = True
For i = 1 To 11
    For j = 1 To 12
        If matrix(i, j) <> vbWhite Then
            nb = 0
            search i, j, matrix(i, j)
            If nb > 1 Then
               Final = False
               Exit Function
            End If
        End If
    Next
Next
End Function
