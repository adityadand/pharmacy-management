VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form medd 
   Caption         =   "MED"
   ClientHeight    =   11550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   26550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "medd.frx":0000
   ScaleHeight     =   13740
   ScaleWidth      =   28680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Data Report"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   25
      Top             =   11040
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   24360
      TabIndex        =   23
      Top             =   7680
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   24360
      TabIndex        =   21
      Top             =   6480
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   735
      Left            =   24360
      TabIndex        =   20
      Top             =   5400
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   82182145
      CurrentDate     =   43575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   24360
      TabIndex        =   17
      Top             =   4320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   82182145
      CurrentDate     =   43575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17160
      TabIndex        =   15
      Top             =   6960
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   24120
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   480
      Picture         =   "medd.frx":CD56E
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   14
      Top             =   360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   1560
      TabIndex        =   13
      Top             =   4680
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   31
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      TabIndex        =   10
      Top             =   11040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   9
      Top             =   11040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12840
      TabIndex        =   8
      Top             =   11040
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15600
      TabIndex        =   7
      Top             =   11040
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   735
      Left            =   21240
      TabIndex        =   6
      Top             =   11040
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      Height          =   735
      Left            =   22800
      TabIndex        =   5
      Top             =   11040
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18360
      TabIndex        =   4
      Top             =   11040
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17160
      TabIndex        =   3
      Top             =   5760
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17160
      TabIndex        =   2
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "S_PRICE"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   21360
      TabIndex        =   24
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "B_PRICE"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   21360
      TabIndex        =   22
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MANUF DT"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   21360
      TabIndex        =   19
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "EXPIRY DT"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   21360
      TabIndex        =   18
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BRAND"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   13680
      TabIndex        =   16
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   -360
      X2              =   32760
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   -360
      X2              =   32760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   25080
      TabIndex        =   12
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICINE"
      BeginProperty Font 
         Name            =   "Nunito ExtraLight"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   1575
      Left            =   12360
      TabIndex        =   11
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICINE NAME"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   13680
      TabIndex        =   1
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICINE ID"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   13680
      TabIndex        =   0
      Top             =   4800
      Width           =   2535
   End
End
Attribute VB_Name = "medd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Command1_Click()

'adding a new record
'con.Open
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
rs.Open "med", con
'rs.AddNew
s1 = "insert into med values('" + Text1.Text + "', '" + Text2.Text + "' , '" + Text3.Text + "'  , '" + Format(DTPicker1.Value, "DD/MM/YYYY") + "' , '" + Format(DTPicker2.Value, "DD/MM/YYYY") + "' , '" + Text4.Text + "' , '" + Text5.Text + "'  )"
con.Execute s1
MsgBox "record added"
MsgBox "added successfully"

DataGrid1.Refresh
Call Form_Load

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Text4.Text = ""
Text5.Text = ""


End Sub

Private Sub Command3_Click()
'update
'con.Open
s1 = "update med set mname = '" & Text2.Text & "' , brand = '" & Text3.Text & "' , MANUF_DT = '" & Format(DTPicker1.Value, "DD/MM/YYYY") & "' , EXPIRY_DT = '" & Format(DTPicker2.Value, "DD/MM/YYYY") & "' , B_PRICE = " & Text4.Text & " , S_PRICE = " & Text5.Text & "   where mid = " & Text1.Text & "  "
MsgBox s1
con.Execute s1

con.Close

DataGrid1.Refresh
Call Form_Load



End Sub

Private Sub Command4_Click()
'delete
Dim str ' =inputbox("enter ")
str = InputBox("enter")
'str = Text1.Text
'con.Open
s1 = "delete from med where mid=" & str & ""
con.Execute s1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Text4.Text = ""
Text5.Text = ""

rs.MoveFirst
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
DTPicker1.Value = rs.Fields(3)
DTPicker2.Value = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)

con.Close
DataGrid1.Refresh
Call Form_Load

End Sub

Private Sub Command5_Click()
rs.MovePrevious
If (rs.BOF) Then
MsgBox ("ÿou are at first record")
Else
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
DTPicker1.Value = rs.Fields(3)
DTPicker2.Value = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)

End If


End Sub
Private Sub Command6_Click()
rs.MoveNext
If (rs.EOF) Then
MsgBox ("no further record")
Else
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
DTPicker1.Value = rs.Fields(3)
DTPicker2.Value = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)

End If

End Sub

Private Sub Command7_Click()
Dim str
str = Val(InputBox("enter the id to search"))
Set rs = New ADODB.Recordset
rs.Open "select * from med where mid = " & str & "", con
If rs.EOF Then
MsgBox ("not found")
Else
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
DTPicker1.Value = rs.Fields(3)
DTPicker2.Value = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)


End If


End Sub

Private Sub Command8_Click()
DataReport4.Show
End Sub

Private Sub Form_Load()

Timer1_Timer

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\pharm.mdb"
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\med.mdb"
con.Open
Set rs = New ADODB.Recordset

rs.CursorLocation = adUseClient
rs.Open "select * from med", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
DTPicker1.Value = rs.Fields(3)
DTPicker2.Value = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)

'Set Text1.DataSource = rs
'Text1.DataField = "mid"
'Set Text1.DataSource = rs
'Text1.DataField = "mname"

'Set Text1.DataSource = rs
'Text1.DataField = "mid"
'Set Text2.DataSource = rs
'Text2.DataField = "dname"
'Set Text3.DataSource = rs
'Text3.DataField = "designation"
'Set Text4.DataSource = rs
'Text4.DataField = "address"
'Set Text5.DataSource = rs
'Text5.DataField = "phno"
'Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 32 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or KeyAscii >= 123 Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER ALPHABETS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Timer1_Timer()
Label9.Caption = Time
End Sub
