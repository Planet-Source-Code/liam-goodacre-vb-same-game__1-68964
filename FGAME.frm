VERSION 5.00
Begin VB.Form FGAME 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBSame Game"
   ClientHeight    =   5295
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   5535
      TabIndex        =   3
      Top             =   4950
      Width           =   5535
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2565
         TabIndex        =   6
         Top             =   75
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   75
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4950
         TabIndex        =   4
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.PictureBox PT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   5700
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   1
      Top             =   75
      Width           =   4815
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   75
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   75
      Width           =   4815
      Begin VB.Image IMG 
         Height          =   480
         Index           =   0
         Left            =   75
         Picture         =   "FGAME.frx":0000
         Top             =   75
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image IMG 
         Height          =   480
         Index           =   3
         Left            =   75
         Picture         =   "FGAME.frx":0C42
         Top             =   1650
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image IMG 
         Height          =   480
         Index           =   4
         Left            =   75
         Picture         =   "FGAME.frx":1884
         Top             =   2175
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image IMG 
         Height          =   480
         Index           =   5
         Left            =   75
         Picture         =   "FGAME.frx":24C6
         Top             =   2700
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image IMG 
         Height          =   480
         Index           =   2
         Left            =   75
         Picture         =   "FGAME.frx":3108
         Top             =   1125
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image IMG 
         Height          =   480
         Index           =   1
         Left            =   75
         Picture         =   "FGAME.frx":3D4A
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image IMG 
         Height          =   480
         Index           =   6
         Left            =   75
         Picture         =   "FGAME.frx":498C
         Top             =   3225
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info:"
      Height          =   195
      Left            =   4950
      TabIndex        =   2
      Top             =   75
      Width           =   315
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mgNEW 
         Caption         =   "New Game"
      End
      Begin VB.Menu mgRES 
         Caption         =   "Restart Game"
      End
      Begin VB.Menu mgLET 
         Caption         =   "Letters"
      End
   End
End
Attribute VB_Name = "FGAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExtFloodFill Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal colorCode As Long, ByVal fillType As Long) As Long

Public LEV As String
Public FLev As String
Public fx As Single, fy As Single
Public SelCount
Public TotalP As Double
Public NOL

Private Sub Form_Load()
NOL = 3
NewLev
End Sub

Function GetColor(ByVal n As Integer) As Long
Select Case n
Case 1: GetColor = vbRed
Case 2: GetColor = vbYellow
Case 3: GetColor = vbGreen
Case 4: GetColor = vbCyan
Case 5: GetColor = vbBlue
Case 6: GetColor = vbMagenta
End Select
End Function

Private Sub ResLev()
On Error Resume Next
TotalP = 0
SelCount = 0
fx = -1
fy = -1
LEV = FLev
P.Enabled = True
PN
End Sub

Private Sub NewLev()
On Error Resume Next
TotalP = 0
SelCount = 0
fx = -1
fy = -1
Randomize
FLev = ""
For X = 0 To 9
For Y = 9 To 0 Step -1
Randomize
FLev = FLev & "|" & Int(1 + (Rnd * NOL)) & "," & X & "," & Y
Next Y
Next X
LEV = FLev
P.Enabled = True
PN
End Sub

Private Sub GetInfo()
Dim gi() As String
Label1.Caption = "Info:"
For i = 1 To NOL
gi = Split("|" & LEV, "|" & i & ",")
Label1.Caption = Label1.Caption & vbCrLf & UCase(Chr(96 + i)) & " = " & UBound(gi)
Next i
End Sub

Private Sub PN(Optional ByVal dx As Double = -1, Optional ByVal dy As Double = -1)
On Error Resume Next
Dim f() As String
Dim g() As String
Dim Col As Long
Dim ADX As Single
Dim ADY As Single
Dim SMX As Integer
Dim cx As Single
f = Split(LEV, "|")
P.Cls
PT.Cls
ADX = 0
cx = 0
SMX = -1
For i = 1 To UBound(f)
g = Split(f(i), ",")
If cx = 10 Then
    If Val(g(1)) <> SMX Then
        ADX = ADX + 1
    End If
End If
If Val(g(1)) <> SMX Then ADY = 0: SMX = Val(g(1)): cx = 0
If Val(g(0)) = 0 Then
ADY = ADY + 1
cx = cx + 1
Else
Col = GetColor(Val(g(0)))
g(1) = (Val(g(1)) - ADX)
g(2) = (Val(g(2)) + ADY)
PT.Line (Val(g(1)) * 32, 16 + Val(g(2)) * 32)-(32 + Val(g(1)) * 32, 16 + Val(g(2)) * 32), Col
PT.Line (16 + Val(g(1)) * 32, Val(g(2)) * 32)-(16 + Val(g(1)) * 32, 32 + Val(g(2)) * 32), Col
End If
Next i
PT.FillStyle = 1
PT.Line (0, 0)-(PT.ScaleWidth - 1, PT.ScaleHeight - 1), vbBlack, B
PT.FillStyle = 0
If dx <> -1 And (dy) <> -1 Then ExtFloodFill PT.hDC, (dx * 32) + 16, ((dy) * 32) + 16, PT.Point((dx * 32) + 16, ((dy) * 32) + 16), 1

SelCount = 0
numlets = 0
ADX = 0
cx = 0
SMX = -1
For i = 1 To UBound(f)
g = Split(f(i), ",")
If cx = 10 Then
    If Val(g(1)) <> SMX Then
        ADX = ADX + 1
    End If
End If
If Val(g(1)) <> SMX Then ADY = 0: SMX = Val(g(1)): cx = 0
If Val(g(0)) = 0 Then
ADY = ADY + 1
cx = cx + 1
Else
numlets = 1
g(1) = (Val(g(1)) - ADX)
g(2) = (Val(g(2)) + ADY)
Col = IIf(Int(PT.Point((Val(g(1)) * 32) + 16, (Val(g(2)) * 32) + 16)) = Int(vbWhite), 6, Val(g(0)) - 1)
SelCount = SelCount + IIf(PT.Point((Val(g(1)) * 32) + 16, (Val(g(2)) * 32) + 16) = vbWhite, 1, 0)
P.PaintPicture IMG(Col), Val(g(1)) * 32, Val(g(2)) * 32
P.Line (Val(g(1)) * 32, Val(g(2)) * 32)-(32 + Val(g(1)) * 32, 32 + Val(g(2)) * 32), vbBlack, B
P.CurrentX = ((Val(g(1)) * 32 + 16) - P.TextWidth(Val(g(1))) / 2) - 2
P.CurrentY = ((Val(g(2)) * 32 + 16) - P.TextHeight(Val(g(2))) / 2) - 1
P.Print UCase(Chr(96 + Val(g(0))))
End If
Next i
P.Line (0, 0)-(P.ScaleWidth - 1, P.ScaleHeight - 1), vbBlack, B
If numlets = 0 Then MsgBox "Good Job! You get an extra 500 Points for clearing the screen!": TotalP = TotalP + 500: P.Enabled = False
If SelCount = 1 Then
If fx <> -1 Then
If fx <> -1 Then
fx = -1
fy = -1
PN
End If
End If
End If
Label2.Caption = "Selected: " & SelCount
Label3.Caption = "Points: " & (SelCount ^ 2)
Label4.Caption = "Score: " & TotalP
GetInfo
End Sub


Private Sub mgLet_Click()
f = Int(Val(InputBox("How many letters?", "Letters?")))
If f <> "" Then NOL = IIf(IIf(Int(Val(f)) > 6, 6, Int(Val(f))) < 2, 2, IIf(Int(Val(f)) > 6, 6, Int(Val(f)))): NewLev
End Sub

Private Sub mgNEW_Click()
NewLev
End Sub

Private Sub mgRes_Click()
ResLev
End Sub

Private Sub P_DblClick()
On Error Resume Next
If fx = -1 Then Exit Sub
Dim g() As String
Dim f() As String
Dim h As String
Dim j As String
Dim SMX As Single
Dim ADY As Single
Dim ADX As Single
Dim cx As Single
f = Split(LEV, "|")
h = ""
ADX = 0
cx = 0
For i = 1 To UBound(f)
g = Split(f(i), ",")
If cx = 10 Then
    If Val(g(1)) <> Val(SMX) Then
        ADX = ADX + 1
        cx = 0
    End If
End If
If Val(g(1)) <> Val(SMX) Then ADY = 0: SMX = Val(g(1)): cx = 0
If Val(g(0)) = 0 Then ADY = ADY + 1: cx = cx + 1
j = IIf(Int(PT.Point(((Val(g(1)) - ADX) * 32) + 16, ((Val(g(2)) + ADY) * 32) + 16)) = Int(vbWhite), "|0," & g(1) & "," & g(2), "|" & g(0) & "," & g(1) & "," & g(2))
h = h & j
Next i
LEV = h
TotalP = TotalP + (SelCount ^ 2)
GetInfo
fx = -1
fy = -1
PN
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Int(PT.Point(16 + Int((IIf(X < 1, 1, X) - 1) / 32) * 32, 16 + Int((IIf(Y < 1, 1, Y) - 1) / 32) * 32)) = Int(vbWhite) Then P_DblClick: Exit Sub
If Int(PT.Point(16 + Int((IIf(X < 1, 1, X) - 1) / 32) * 32, 16 + Int((IIf(Y < 1, 1, Y) - 1) / 32) * 32)) = Int(vbBlack) Then fx = -1: fy = -1: PN: Exit Sub
fx = Int((IIf(X < 1, 1, X) - 1) / 32)
fy = Int((IIf(Y < 1, 1, Y) - 1) / 32)
PN fx, fy
End Sub
