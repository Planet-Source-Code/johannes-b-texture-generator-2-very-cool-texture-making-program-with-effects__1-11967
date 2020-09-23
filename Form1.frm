VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TEXTURE GENERATOR 2"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Effects"
      Height          =   1815
      Left            =   5040
      TabIndex        =   30
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   37
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   0
         TabIndex        =   33
         Top             =   600
         Width           =   2175
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   35
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Glass effect"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   31
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Dizzy (faster)"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Random size"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   3720
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "DRAW!!!"
         Height          =   375
         Left            =   6000
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export bmp..."
         Height          =   375
         Left            =   6000
         TabIndex        =   38
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Graphic"
         Height          =   735
         Left            =   0
         TabIndex        =   23
         Top             =   1560
         Width           =   6015
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   29
            ToolTipText     =   "Click to change"
            Top             =   480
            Width           =   495
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Select"
            Height          =   195
            Left            =   1680
            TabIndex        =   28
            ToolTipText     =   "Select a color (not recommended)"
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Get color (cool)"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "Very good interlace method! (recommended)"
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cool interlace (cool when drawing)"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Do not use this if you want a nice picture to save! This effect is only cool when drawing!"
            Top             =   240
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.Line Line4 
            X1              =   3000
            X2              =   3000
            Y1              =   120
            Y2              =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Pixel size (1 = slow, 5 = fast/ugly)"
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   4695
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   22
            Text            =   "1"
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   21
            Text            =   "2"
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   960
            MaxLength       =   1
            TabIndex        =   20
            Text            =   "0"
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   17
            Text            =   "20"
            ToolTipText     =   "Double this value if you decrase your pixel size from 2 to 1"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "3"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "2"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   960
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Auto incrase"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Delay (size)"
            Height          =   255
            Left            =   2280
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "RGB (step)"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Color"
         Height          =   615
         Left            =   1680
         TabIndex        =   5
         Top             =   120
         Width           =   2055
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   720
            MaxLength       =   3
            TabIndex        =   11
            Text            =   "255"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   360
            MaxLength       =   3
            TabIndex        =   10
            Text            =   "200"
            Top             =   360
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "RGB (1-255)"
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   180
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   0
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "200"
            Top             =   360
            Width           =   375
         End
         Begin VB.OptionButton Option3 
            Caption         =   "VB"
            Height          =   195
            Left            =   1320
            TabIndex        =   7
            Top             =   180
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   6
            Text            =   "999999"
            Top             =   360
            Width           =   735
         End
         Begin VB.Line Line2 
            X1              =   1200
            X2              =   1200
            Y1              =   120
            Y2              =   600
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Random:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Color cycling"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   1320
         X2              =   1680
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   1680
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   0
      Top             =   2280
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B As Integer
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim D, R, Q As Integer


Dim RR, GG, BB, RRR, GGG, BBB As Integer

Private Sub Check1_Click()
If Check1.Value = False Then
Option5.Enabled = False
Option6.Enabled = False
Picture2.Enabled = False
Else
Text14.Text = 0
Check1.Value = 1
Option5.Enabled = True
Option6.Enabled = True
Picture2.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Picture1.Cls
A = 0
B = 0
RR = 0
GG = 0
BB = 0
RRR = 0
GGG = 0
BBB = 0
RRRR = 0
GGGG = 0
BBBB = 0

D = Combo1.Text
R = Text1.Text
On Error Resume Next
Form1.Caption = "Drawing..."
Do

If Option2.Value = True And Option3.Value = True Then
Picture1.ForeColor = Rnd * R
End If
If Option2.Value = True And Option4.Value = True Then
Picture1.ForeColor = RGB(Rnd * Text2.Text, Rnd * Text2.Text, Rnd * Text2.Text)
End If

If Option1.Value = True And T >= Text8.Text Then
T = 0

If RR > 255 Then RRR = 1: Text5.Text = Text5.Text + Val(Text9.Text)
If RR <= 0 Then RRR = 0

If GG > 255 Then GGG = 1: Text6.Text = Text6.Text + Val(Text10.Text)
If GG <= 0 Then GGG = 0

If BB > 255 Then GGG = 1: Text7.Text = Text7.Text + Val(Text11.Text)
If BB <= 0 Then GGG = 0


If RRR = 0 Then
RR = RR + Text5.Text
Else
RR = RR - Text5.Text
End If

If GGG = 0 Then
GG = GG + Text6.Text
Else
GG = GG - Text6.Text
End If

If GGG = 0 Then
BB = BB + Text7.Text
Else
BB = BB - Text7.Text
End If



Picture1.ForeColor = RGB(RR + Rnd * Val(Text13.Text), GG + Rnd * Val(Text13.Text), BB + Rnd * Val(Text13.Text))
End If

T = T + 1

If Check1.Value = 0 Then
'DRAW!!!
Picture1.Circle (A, B), D + Rnd * Val(Text12.Text)
Else
Picture1.Circle (A, B), D + 1 + Rnd * Val(Text12.Text)
'Interlace
If Option6.Value = True Then Picture1.ForeColor = Picture2.BackColor
Picture1.Line (0, A)-(Picture1.ScaleWidth, A)
A = A + 1
End If

A = A + D + Rnd * Val(Text14.Text)

If A >= Picture1.ScaleWidth Then
A = 0
B = B + D + Rnd * Val(Text14.Text)
Picture1.Refresh
End If


Loop Until B > Picture1.ScaleHeight
Picture1.Refresh
Form1.Caption = "TEXTURE GENERATOR 2"
End Sub

Private Sub Command2_Click()
CM.CancelError = True
On Error GoTo ui
CM.Filter = "Bitmap (*.bmp)|*.BMP"
CM.ShowSave
SavePicture Picture1.Image, CM.FileName
Exit Sub
ui:
Exit Sub
End Sub


Private Sub Form_Load()
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"

Combo1 = 2
End Sub


Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please vote if you like the program!"
End Sub


Private Sub Picture2_Click()
CM.CancelError = True
On Error GoTo kalle

CM.ShowColor

Picture2.BackColor = CM.Color

Exit Sub
kalle:
Exit Sub
End Sub


Private Sub Text14_Change()
Check1.Value = 0
End Sub

Private Sub Text15_Change()
End Sub

Private Sub Text8_Change()
If Text8.Text = 0 Then Text8.Text = 1
End Sub


