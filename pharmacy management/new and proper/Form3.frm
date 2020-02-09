VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form customerr 
   Caption         =   "CUSTOMER"
   ClientHeight    =   12570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   28680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   12570
   ScaleWidth      =   28680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Show Report"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   11160
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   23760
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   480
      Picture         =   "Form3.frx":9403B
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   25
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   735
      Left            =   24720
      TabIndex        =   23
      Top             =   8640
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   735
      Left            =   24720
      TabIndex        =   22
      Top             =   7800
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3975
      Left            =   960
      TabIndex        =   21
      Top             =   4320
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   25
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
         Size            =   12
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
      TabIndex        =   18
      Top             =   11160
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      Height          =   855
      Left            =   22920
      TabIndex        =   17
      Top             =   11040
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   855
      Left            =   21360
      TabIndex        =   16
      Top             =   11040
      Width           =   1455
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
      TabIndex        =   15
      Top             =   11160
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
      TabIndex        =   14
      Top             =   11160
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
      TabIndex        =   13
      Top             =   11160
      Width           =   2055
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
      TabIndex        =   12
      Top             =   11160
      Width           =   2055
   End
   Begin VB.TextBox Text6 
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
      Left            =   24720
      TabIndex        =   11
      Top             =   6480
      Width           =   3135
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
      Left            =   24720
      MaxLength       =   10
      TabIndex        =   9
      Top             =   5160
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
      Left            =   24720
      TabIndex        =   8
      Top             =   3840
      Width           =   3135
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
      Left            =   18360
      TabIndex        =   7
      Top             =   7200
      Width           =   3135
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
      Left            =   18360
      TabIndex        =   6
      Top             =   5880
      Width           =   3135
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
      Left            =   18360
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
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
      Height          =   615
      Left            =   22080
      TabIndex        =   24
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   33240
      Y1              =   10200
      Y2              =   10200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   -240
      X2              =   32880
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER"
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
      Left            =   12120
      TabIndex        =   20
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
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
      Left            =   24600
      TabIndex        =   19
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
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
      Height          =   855
      Left            =   21960
      TabIndex        =   10
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO"
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
      Height          =   855
      Left            =   21960
      TabIndex        =   4
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTOR NAME"
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
      Left            =   21960
      TabIndex        =   3
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTOR ID"
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
      Height          =   855
      Left            =   15480
      TabIndex        =   2
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
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
      Height          =   855
      Left            =   15480
      TabIndex        =   1
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ID"
      BeginProperty Font 
         Name            =   "Nunito Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   855
      Left            =   15480
      TabIndex        =   0
      Top             =   4560
      Width           =   2655
   End
End
Attribute VB_Name = "customerr"
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
rs.Open "cust", con
'rs.AddNew
If Option1.Value = True Then
gen = Option1.Caption
Else
gen = Option2.Caption
End If


s1 = "insert into cust values('" + Text1.Text + "', '" + Text2.Text + "', '" + Text3.Text + "' ,'" + Text4.Text + "','" + Text5.Text + "', '" + Text6.Text + "' ,'" + gen + "' )"
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
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub

Private Sub Command3_Click()
'update
'con.Open
If Option1.Value = True Then
gen = Option1.Caption
Else
gen = Option2.Caption
End If

s1 = "update cust set cname = '" & Text2.Text & "' , did = " & Text3.Text & " , dname = '" & Text4.Text & "' , MOBILE_NO = " & Text5.Text & " , caddress = '" & Text6.Text & "' , gender = '" & gen & "'  where cid = " & Text1.Text & " "
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
s1 = "delete from cust where cid =" & str & ""
con.Execute s1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Option1.Value = False
Option2.Value = False
rs.MoveFirst
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Option1.Caption = rs.Fields(6)
If (Option1.Caption = "MALE") Then
Option1.Value = True
Else
Option2.Value = True
End If

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
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Option1.Caption = rs.Fields(6)
If (Option1.Caption = "MALE") Then
Option1.Value = True
Else
Option2.Value = True

End If


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
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Option1.Caption = rs.Fields(6)
If (Option1.Caption = "MALE") Then
Option1.Value = True
Else
Option2.Value = True
End If

End If

End Sub

Private Sub Command7_Click()
Dim str
str = Val(InputBox("enter the id to search"))
Set rs = New ADODB.Recordset
rs.Open "select * from cust where cid = " & str & "", con
If rs.EOF Then
MsgBox ("not found")
Else
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)

End If


End Sub

Private Sub Command8_Click()
DataReport2.Show
End Sub

Private Sub Form_Load()

Timer1_Timer
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\pharm.mdb"
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\cust.mdb"
con.Open
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select * from cust", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)



'Set Text1.DataSource = rs
'Text1.DataField = "did"
'Set Text2.DataSource = rs
'Text2.DataField = "dname"
'Set Text3.DataSource = rs
'Text3.DataField = "designation"
'Set Text4.DataSource = rs
'Text4.DataField = "address"
'Set Text5.DataSource = rs
'Text5.DataField = "phno"
'Adodc1.Refresh
'DataGrid1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 32 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or KeyAscii >= 123 Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER ALPHABETS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text3_LostFocus()
Dim str As String
str = Text3.Text
Set rs = New ADODB.Recordset

rs.Open "select dname from  doc where did = " & str & " ", con
Text4.Text = rs!dname

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 32 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or KeyAscii >= 123 Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER ALPHABETS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Timer1_Timer()
Label7.Caption = Time
End Sub
