VERSION 5.00
Begin VB.Form loginn 
   Caption         =   "Form1"
   ClientHeight    =   11835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22800
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   11835
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   36
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   855
      Left            =   16080
      TabIndex        =   8
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   14280
      PasswordChar    =   "#"
      TabIndex        =   6
      Top             =   7920
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14280
      TabIndex        =   5
      Top             =   6840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14280
      TabIndex        =   4
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   18
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13080
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11640
      TabIndex        =   3
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "EMAIL ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11640
      TabIndex        =   2
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11640
      TabIndex        =   1
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LOGIN "
      BeginProperty Font 
         Name            =   "Nunito ExtraBold"
         Size            =   23.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   13920
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      FillStyle       =   0  'Solid
      Height          =   8175
      Left            =   10800
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   7455
   End
End
Attribute VB_Name = "loginn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 

Dim a, b, c, d
a = Text1.Text
b = Text2.Text
c = Text3.Text

If (a = "admin" And b = "mail@admin" And c = "#access") Then
MsgBox ("welcome")
CreateObject("sapi.SPvoice").speak ("welcome")
MDIForm1.Show
Else
MsgBox ("wrong credientals")
End If
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = RGB(230, 73, 25)
End Sub

Private Sub Form_Load()
Command1.BackColor = RGB(96, 3, 252)

Label1.BackColor = RGB(6, 12, 33)
Label1.ForeColor = RGB(135, 254, 4)
Label2.BackColor = RGB(6, 12, 33)
Label3.BackColor = RGB(6, 12, 33)
Label4.BackColor = RGB(6, 12, 33)
Text4.BackColor = RGB(6, 12, 33)
Text1.BackColor = RGB(6, 12, 33)
Text2.BackColor = RGB(6, 12, 33)
Text3.BackColor = RGB(6, 12, 33)
Text1.ForeColor = RGB(46, 204, 113)
Text2.ForeColor = RGB(46, 204, 113)
Text3.ForeColor = RGB(46, 204, 113)
Shape1.FillColor = RGB(6, 12, 33)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
  Case 32 To 64, 91 To 96, 123 To 126
     MsgBox ("must be a letter ! please try again !")
     KeyAscii = 0
   Exit Sub
 End Select
End Sub



Private Sub Text4_Change()
If (Text4.Text = "a" Or Text4.Text = "A") Then
Text1.Text = "admin"
Text2.Text = "mail@admin"
Text3.Text = "#access"
End If


End Sub
