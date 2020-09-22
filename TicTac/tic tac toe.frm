VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "X and Zero"
   ClientHeight    =   2565
   ClientLeft      =   4845
   ClientTop       =   4470
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3585
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1965
      TabIndex        =   9
      Top             =   1065
      Width           =   1575
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   8
      Left            =   1260
      TabIndex        =   8
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   7
      Left            =   735
      TabIndex        =   7
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   210
      TabIndex        =   6
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   5
      Left            =   1260
      TabIndex        =   5
      Top             =   840
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   4
      Left            =   735
      TabIndex        =   4
      Top             =   840
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   210
      TabIndex        =   3
      Top             =   840
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   2
      Left            =   1260
      TabIndex        =   2
      Top             =   315
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   735
      TabIndex        =   1
      Top             =   315
      Width           =   540
   End
   Begin VB.CommandButton buton 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   540
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1965
      TabIndex        =   11
      Top             =   450
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   2205
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private i As Integer, j As Integer
Private mat2(1 To 8, 1 To 3) As Integer, mat3() As Integer
Private t As Integer
Private terminat As Boolean
Private Sub buton_Click(Index As Integer)
If terminat = True Then Exit Sub
Dim k As Integer, flag1 As Byte, flag2 As Byte
Dim a As Boolean
Dim c1(8) As Integer, c(8, 3) As Integer
Dim mat(1 To 3, 1 To 3) As Integer
a = False
If buton(Index).Caption <> "0" And buton(Index).Caption <> "X" Then
    buton(Index).Caption = "X"
    a = True
End If
t = 0
For i = 1 To 3
For j = 1 To 3
    mat(i, j) = t
    t = t + 1
Next j
Next i
'------------------------------------------------------------
If a = True Then
    ReDim mat3(1 To 8, 1 To 3) As Integer
            For i = 1 To 3
            For j = 1 To 3
                mat2(i, j) = buton(mat(i, j)).Index
                If buton(mat2(i, j)).Caption = "X" Then mat3(i, j) = 11
                If buton(mat2(i, j)).Caption = "0" Then mat3(i, j) = 5
                If buton(mat2(i, j)).Caption = "" Then mat3(i, j) = 7
            Next j
            Next i
            For i = 1 To 3
            For j = 1 To 3
                mat2(i + 3, j) = buton(mat(j, i)).Index
                If buton(mat2(i + 3, j)).Caption = "X" Then mat3(i + 3, j) = 11
                If buton(mat2(i + 3, j)).Caption = "0" Then mat3(i + 3, j) = 5
                If buton(mat2(i + 3, j)).Caption = "" Then mat3(i + 3, j) = 7
            Next j
            Next i
            For i = 1 To 3
                mat2(7, i) = buton(mat(i, i)).Index
                If buton(mat2(7, i)).Caption = "X" Then mat3(7, i) = 11
                If buton(mat2(7, i)).Caption = "0" Then mat3(7, i) = 5
                If buton(mat2(7, i)).Caption = "" Then mat3(7, i) = 7
            Next i
            For i = 1 To 3
                mat2(8, i) = buton(mat(i, 4 - i)).Index
                If buton(mat2(8, i)).Caption = "X" Then mat3(8, i) = 11
                If buton(mat2(8, i)).Caption = "0" Then mat3(8, i) = 5
                If buton(mat2(8, i)).Caption = "" Then mat3(8, i) = 7
            Next i
            For i = 1 To 8
                For j = 1 To 3
                If j = 1 Then
                c(mat2(i, 1), c1(mat2(i, j))) = mat3(i, 2) + mat3(i, 3)
                End If
                If j = 2 Then
                c(mat2(i, 2), c1(mat2(i, 2))) = mat3(i, 1) + mat3(i, 3)
                End If
                If j = 3 Then
                c(mat2(i, 3), c1(mat2(i, 3))) = mat3(i, 2) + mat3(i, 1)
                End If
                c1(mat2(i, j)) = c1(mat2(i, j)) + 1
                Next j
            Next i
           '-----------------------------------------------------------------------------------------
            For k = 0 To 8
            If buton(k).Caption = "" Then
           '------------------------------------------------------------------------------------------
                If flag1 <= 1 Then
                    flag2 = k
                    flag1 = 1
                End If
            '------------------------------------------------------------------------------------------
                If flag1 < 10 Then
                    If c(k, 0) = 18 Or c(k, 1) = 18 Or c(k, 2) = 18 Or c(k, 3) = 18 Then
                    flag2 = k
                    flag1 = 10
                    End If
                End If
             '------------------------------------------------------------------------------------------
                If flag1 < 15 Then
                    If (c(k, 0) = 18 Or c(k, 1) = 18 Or c(k, 2) = 18 Or c(k, 3) = 18) And k = 4 Then
                    flag2 = k
                    flag1 = 15
                    End If
                End If
            '------------------------------------------------------------------------------------------
                If flag1 < 25 Then
                    If c(k, 0) = 12 Or c(k, 1) = 12 Or c(k, 2) = 12 Or c(k, 3) = 12 Then
                    flag2 = k
                    flag1 = 25
                    End If
                End If
                '------------------------------------------------------------------------------------------
                If flag1 < 26 Then
                    If (c(k, 0) = 12 Or c(k, 1) = 12 Or c(k, 2) = 12 Or c(k, 3) = 12) And _
                    (c(k, 0) = 18 Or c(k, 1) = 18 Or c(k, 2) = 18 Or c(k, 3) = 18) Then
                    flag2 = k
                    flag1 = 26
                    End If
                End If
                '------------------------------------------------------------------------------------------
                If flag1 < 33 Then
                    If (c(k, 0) = 18 And c(k, 1) = 18) Or (c(k, 0) = 18 And c(k, 2) = 18) _
                    Or (c(k, 0) = 18 And c(k, 3) = 18) Or (c(k, 1) = 18 And c(k, 2) = 18) _
                    Or (c(k, 1) = 18 And c(k, 3) = 18) Or (c(k, 2) = 18 And c(k, 3) = 18) Then
                    flag2 = k
                    flag1 = 33
                    End If
                End If
                '------------------------------------------------------------------------------------------
                If flag1 < 34 Then
                    If (c(k, 0) = 12 Or c(k, 1) = 12 Or c(k, 2) = 12 Or c(k, 3) = 12) And _
                        (k = 4 Or ((k = 1 Or k = 5 Or k = 3 Or k = 7) And buton(4).Caption = "0")) Then
                    flag2 = k
                    flag1 = 34
                    End If
                End If
                '------------------------------------------------------------------------------------------
                If flag1 < 40 Then
                    If ((c(k, 0) = 18 And c(k, 1) = 18) Or (c(k, 0) = 18 And c(k, 2) = 18) _
                    Or (c(k, 0) = 18 And c(k, 3) = 18) Or (c(k, 1) = 18 And c(k, 2) = 18) _
                    Or (c(k, 1) = 18 And c(k, 3) = 18) Or (c(k, 2) = 18 And c(k, 3) = 18)) _
                    And (c(k, 0) = 12 Or c(k, 1) = 12 Or c(k, 2) = 12 Or c(k, 3) = 12) And buton(4).Caption = "" Then
                    flag2 = k
                    flag1 = 40
                    End If
                End If
            '------------------------------------------------------------------------------------------
            If flag1 < 50 Then
                    If (c(k, 0) = 12 And c(k, 1) = 12) Or (c(k, 0) = 12 And c(k, 2) = 12) _
                    Or (c(k, 0) = 12 And c(k, 3) = 12) Or (c(k, 1) = 12 And c(k, 2) = 12) _
                    Or (c(k, 1) = 12 And c(k, 3) = 12) Or (c(k, 2) = 12 And c(k, 3) = 12) Then
                    flag2 = k
                    flag1 = 50
                    End If
                End If
            '------------------------------------------------------------------------------------------
                If flag1 < 220 Then
                    If c(k, 0) = 22 Or c(k, 1) = 22 Or c(k, 2) = 22 Or c(k, 3) = 22 Then
                    flag2 = k
                    flag1 = 220
                    End If
                End If
           '------------------------------------------------------------------------------------------
                If flag1 < 250 Then
                    If c(k, 0) = 10 Or c(k, 1) = 10 Or c(k, 2) = 10 Or c(k, 3) = 10 Then
                    flag2 = k
                    flag1 = 250
                    Label1.Caption = "You are a loser !"
                    terminat = True
                    End If
                End If
            End If

            Next k
If flag1 > 0 Then buton(flag2).Caption = "0"
If flag1 = 1 Or flag1 = 0 Then Label1.Caption = "Draw !"
End If
End Sub

Private Sub Command1_Click()
terminat = False
Randomize
For i = 0 To 8
buton(i).Caption = ""
Next i
i = CInt(Rnd * 1)
If i = 0 Then
buton(Rnd * 8).Caption = "0"
Label2.Caption = "Program move first !"
Else
Label2.Caption = "You move first !"
End If
ReDim mat(1 To 3, 1 To 3) As Integer
t = 0
For i = 1 To 3
For j = 1 To 3
    mat(i, j) = t
    t = t + 1
Next j
Next i
Label1.Caption = ""
End Sub

Private Sub Form_Activate()
Command1.SetFocus
ReDim mat(1 To 3, 1 To 3)
Randomize
i = CInt(Rnd * 1)
If i = 0 Then
buton(Rnd * 8).Caption = "0"
Label2.Caption = "Program move first !"
Else
Label2.Caption = "You move first !"
End If
End Sub

