VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WavVortex  -  Dmkware.2000"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   3240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   -120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Variables"
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   7335
      Begin VB.OptionButton Option1 
         Caption         =   "Distort6"
         Height          =   255
         Index           =   12
         Left            =   3840
         TabIndex        =   41
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Distort4"
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   38
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   37
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   36
         Top             =   1800
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   35
         Top             =   1800
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Split"
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LP Dist"
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   31
         Top             =   1440
         Width           =   845
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Distort2"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   30
         Top             =   720
         Width           =   855
      End
      Begin MSComctlLib.Slider sA1 
         Height          =   255
         Left            =   495
         TabIndex        =   28
         Top             =   600
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider sA2 
         Height          =   255
         Left            =   495
         TabIndex        =   29
         Top             =   840
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Text            =   "0"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Text            =   "20"
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "FM"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1600
         Value           =   1  'Checked
         Width           =   555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Noise"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         Height          =   855
         Left            =   5790
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   20
         Top             =   1200
         Width           =   1400
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   44
            X2              =   0
            Y1              =   52
            Y2              =   32
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   600
            Shape           =   3  'Circle
            Top             =   720
            Width           =   135
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   1170
            Shape           =   3  'Circle
            Top             =   30
            Width           =   135
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clone"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LP"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Text            =   "2"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   405
         Left            =   5760
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Play"
         Height          =   435
         Left            =   5760
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Text            =   "11250"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   5655
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Text            =   "1"
         Top             =   1200
         Width           =   495
      End
      Begin MSComctlLib.Slider sF1 
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Min             =   -6000
         Max             =   6000
         SelStart        =   32
         TickFrequency   =   500
         Value           =   32
      End
      Begin MSComctlLib.Slider sF2 
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Min             =   -6000
         Max             =   6000
         SelStart        =   7
         TickFrequency   =   500
         Value           =   1102
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Distort3"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "FreakD"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Half"
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   39
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Distort1"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   16
         Top             =   480
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Distort5"
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   40
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "FM Speed"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F2"
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Decay"
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Attack"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "FM Range"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Osc2 Options:"
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A2"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wav Length    (11250=1s)"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F1"
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sine(35) As Single, CoSn(35) As Single, ChangeTable As Integer, Val As Integer, LastVal As Integer, CurVU As Integer, SFr As Single
Dim leng As Long, N As Long, I As Long, s3 As Byte, gen As Boolean
Dim A1 As Single, F1 As Single
Dim A2 As Single, F2 As Single
Dim Atk1 As Single, Dk1 As Single, s1 As Byte
Dim Atk2 As Single, Dk2 As Single, s2 As Byte
Dim Sample1 As Byte, Sample2 As Byte

Private Sub Check2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then
Check2(0).Value = 0
Check2(1).Value = 1
Else
Check2(0).Value = 1
Check2(1).Value = 0
End If
End Sub

Private Sub Command1_Click(): leng = Text3: ReDim Buffer(leng)
If gen = True Then Exit Sub

Picture1.ScaleMode = 0: Picture1.ScaleHeight = 255: Picture1.ScaleWidth = leng: Picture1.Cls

s1 = 0: s2 = 0: s3 = 1: Atk1 = 0: Atk2 = 0: Dk1 = A1: Dk2 = A2: SFr = F1: gen = True

Command2.Enabled = False: Command4.Enabled = True

For I = 0 To leng: N = I * 0.01745329251994: DoEvents
On Error Resume Next
If Check1.Value = 1 Then
If Check2(1).Value = 1 And -(I And 1) Then
If s3 = 0 Then
sF1.Value = sF1.Value + 1
sF2.Value = sF2.Value + 1
If sF1.Value > Text8 Then s3 = 1
Else
sF1.Value = sF1.Value - 1
sF2.Value = sF2.Value - 1
If sF1.Value < Text9 Then s3 = 0
End If
ElseIf Check2(0).Value = 1 Then
If s3 = 0 Then
sF1.Value = sF1.Value + 1
sF2.Value = sF2.Value + 1
If sF1.Value > Text8 Then s3 = 1
Else
sF1.Value = sF1.Value - 1
sF2.Value = sF2.Value - 1
If sF1.Value < Text9 Then s3 = 0
End If
End If
End If
A1 = sA1.Value: F1 = 0.0001 * sF1.Value
A2 = sA2.Value: F2 = 0.0001 * sF2.Value
If s1 = 0 Then
If Atk1 < A1 Then Atk1 = Atk1 + (Text6 / 100)
If Atk1 >= A1 And Not ((leng - I) * (Text7 / 100)) > A1 Then s1 = 1: Dk1 = A1
Sample1 = (Cos(F1 * I) * Atk1) + &H7F
Else
If Dk1 > 0 Then Dk1 = Dk1 - (Text7 / 100)
If Dk1 < 0 Then s1 = 0: Atk1 = 0
Sample1 = (Cos(F1 * I) * Dk1) + &H7F
End If
If s2 = 0 Then
If Atk2 < A2 Then Atk2 = Atk2 + (Text6 / 100)
If Atk2 >= A2 And Not ((leng - I) * (Text7 / 100)) > A2 Then s2 = 1: Dk2 = A2
Sample2 = (Cos(F2 * I) * Atk2) + &H7F
If Option1(0).Value Then
Sample2 = (Sample1 / (10 * Log(1 + (F2 / (0.1 * Atk2))))) + 10
ElseIf Option1(1).Value Then
Sample2 = Sample1 / Int(Cos(2 * 3.14 * I * F1) * Atn(I * 4)) + &H7F
ElseIf Option1(3).Value Then
Sample2 = Sample2 / (Atk2 * Tan(N * Atk2))
ElseIf Option1(4).Value Then
Sample2 = (Rnd * Atk2) + &H7F
ElseIf Option1(5).Value Then
Sample2 = Sample1 / Log(Sample1 - Sample2)
ElseIf Option1(6).Value Then
Sample2 = (Sample1 / (10 * Log(1 + (F2 / (0.1 * Sample1))))) / 2
ElseIf Option1(7).Value Then
If Sample1 < 127 Then
Sample2 = Sample1 - 10
Else
Sample2 = Sample1 + 10
End If
ElseIf Option1(8).Value Then
Sample2 = Cos(Sample1 * (3.14 / 180))
ElseIf Option1(9).Value Then
Sample2 = ((Sample1) + (Sample2)) / 2
ElseIf Option1(10).Value Then
Sample2 = Abs(Sample1 - &H7F) + &H7F
ElseIf Option1(11).Value Then
Sample2 = (Sample1 + Sample2) / Atk2
ElseIf Option1(12).Value Then
Sample2 = Sample1 / Int(Cos(2 * 3.14 * I * F1) * Atn(I * 4)) + &H7F
Sample2 = (0.2 * Sample2) + &H7F
End If
Else
If Dk2 > 0 Then Dk2 = Dk2 - (Text7 / 100)
If Dk2 < 0 Then s2 = 0: Atk2 = 0
Sample2 = (Cos(F2 * I) * Dk2) + &H7F
If Option1(0).Value Then
Sample2 = (Sample1 / (10 * Log(1 + (F2 / (0.1 * Dk2))))) + 10
ElseIf Option1(1).Value Then
Sample2 = Sample1 / Int(Cos(2 * 3.14 * I * F1) * Atn(I * 4)) + &H7F
ElseIf Option1(3).Value Then
Sample2 = Sample2 / (Dk2 * Tan(N * Dk2))
ElseIf Option1(4).Value Then
Sample2 = (Rnd * Dk2) + &H7F
ElseIf Option1(5).Value Then
Sample2 = Sample1 / Log(Sample1 - Sample2)
ElseIf Option1(6).Value Then
Sample2 = (Sample1 / (10 * Log(1 + (F2 / (0.1 * Sample1))))) / 2
ElseIf Option1(7).Value Then
If Sample1 < 127 Then
Sample2 = Sample1 - 10
Else
Sample2 = Sample1 + 10
End If
ElseIf Option1(8).Value Then
Sample2 = Cos(Sample1 * (3.14 / 180))
ElseIf Option1(9).Value Then
Sample2 = ((Sample1) + (Sample2)) / 2
ElseIf Option1(10).Value Then
Sample2 = Abs(Sample1 - &H7F) + &H7F
ElseIf Option1(11).Value Then
Sample2 = (Sample1 + Sample2) / Dk2
ElseIf Option1(12).Value Then
Sample2 = Sample1 / Int(Cos(2 * 3.14 * I * F1) * Atn(I * 4)) + &H7F
Sample2 = (0.2 * Sample2) + &H7F
End If
End If
    If -(I And 1) Then
        Picture1.PSet (I, Sample1), vbGreen
        If A1 > 0 Then
        Buffer(I) = Sample1
        Else
        Buffer(I) = Sample2
        End If
    Else
        Picture1.PSet (I, Sample2), vbYellow
        If A2 > 0 Then
        Buffer(I) = Sample2
        Else
        Buffer(I) = Sample1
        End If
    End If
If gen = False Then GoTo ExitLoop
Next
ExitLoop:
gen = False
WH.RiffID = "RIFF"
WH.RiffLength = leng - 8
WH.WavID = "WAVE"
WH.FmtID = "fmt "
WH.FmtLength = 16
WH.wavformattag = 1
WH.Channels = 1
WH.SamplesPerSec = 11250
WH.BytesPerSec = 0
WH.BlockAlign = 11250
WH.FmtSpecific = 0
WH.Padding = 524289
WH.DataID = "data"
WH.DataLength = leng - 44

Open App.Path & "\temp.wav" For Binary As #1
Put #1, , WH
Put #1, , Buffer()
Close #1

Command2.Enabled = True: Command4.Enabled = False
End Sub

Private Sub Command2_Click()
LoadFile App.Path & "\temp.wav", 1
Play 1, True, 0
End Sub

Private Sub Command3_Click()
On Error GoTo canceled
CD.CancelError = True
CD.Filter = "Wav Files (*.wav)|*.wav"
CD.ShowSave

Open CD.FileName For Binary As #1
Put #1, , WH
Put #1, , Buffer()
Close #1
canceled:
End Sub

Private Sub Command4_Click()
gen = False
End Sub

Private Sub Form_Load()
Initialize_DSEngine Form1.hWnd, 44100
For I = 0 To 35
Sine(I) = Sin(I * (3.14 / 18))
CoSn(I) = Cos(I * (3.14 / 18))
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Terminate_DSEngine: End
End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3) Then Text3 = 36000
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6) Then Text6 = 0.8
End Sub

Private Sub Text7_Change()
If Not IsNumeric(Text7) Then Text7 = 0.4
End Sub

Private Sub Timer1_Timer()
CurVU = GetVuStatus
If CurVU > LastVal Then
Val = Val + 1
ElseIf CurVU > LastVal + 2 Then
Val = Val + 2
ElseIf CurVU < LastVal Then
Val = Val - 1
ElseIf CurVU < LastVal - 2 Then
Val = Val - 2
End If
Update Val
LastVal = Val
End Sub

Sub Update(Value As Integer)
Select Case Value
Case 0: ChangeTable = 25
Shape2.FillColor = &H80&
Case 1: ChangeTable = 24
Shape2.FillColor = &H80&
Case 2: ChangeTable = 23
Shape2.FillColor = &H80&
Case 3: ChangeTable = 21
Shape2.FillColor = &H80&
Case 4: ChangeTable = 20
Shape2.FillColor = &H80&
Case 5: ChangeTable = 19
Shape2.FillColor = &H80&
Case 6: ChangeTable = 17
Shape2.FillColor = &H80&
Case 7: ChangeTable = 16
Shape2.FillColor = &H80&
Case 8: ChangeTable = 14
Shape2.FillColor = &H80&
Case 9: ChangeTable = 12
Shape2.FillColor = &HFF
Case 10: ChangeTable = 11
Shape2.FillColor = &HFF
Case Else: ChangeTable = 25
Shape2.FillColor = &H80&
End Select
Line1.X2 = Line1.X1 + (40 * Sine(ChangeTable))
Line1.Y2 = Line1.Y1 + (40 * CoSn(ChangeTable))
End Sub

