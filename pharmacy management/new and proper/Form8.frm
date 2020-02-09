VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form sinvoicee 
   Caption         =   "SINVOICE"
   ClientHeight    =   12810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   26595
   BeginProperty Font 
      Name            =   "Nunito Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   Picture         =   "Form8.frx":0000
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
      Left            =   5280
      TabIndex        =   31
      Top             =   9480
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   24360
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   240
      Picture         =   "Form8.frx":A1DE5
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   30
      Top             =   240
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   1200
      TabIndex        =   29
      Top             =   10920
      Width           =   26295
      _ExtentX        =   46381
      _ExtentY        =   5106
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
      Left            =   18600
      TabIndex        =   27
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   22680
      TabIndex        =   26
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   21120
      TabIndex        =   25
      Top             =   9480
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
      Left            =   15840
      TabIndex        =   24
      Top             =   9480
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
      Left            =   13320
      TabIndex        =   23
      Top             =   9480
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
      Left            =   8160
      TabIndex        =   22
      Top             =   9480
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
      Left            =   10800
      TabIndex        =   21
      Top             =   9480
      Width           =   2055
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
      Left            =   11520
      TabIndex        =   19
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox Text10 
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
      Left            =   19440
      TabIndex        =   18
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox Text9 
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
      Left            =   19440
      MaxLength       =   10
      TabIndex        =   16
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox Text8 
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
      Left            =   19440
      TabIndex        =   15
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   19440
      TabIndex        =   12
      Top             =   4080
      Width           =   2415
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
      Left            =   19440
      TabIndex        =   11
      Top             =   3000
      Width           =   2415
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
      Left            =   11520
      TabIndex        =   8
      Top             =   7200
      Width           =   2415
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
      Left            =   11520
      TabIndex        =   7
      Top             =   5040
      Width           =   2415
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
      Height          =   735
      Left            =   11520
      TabIndex        =   6
      Top             =   4080
      Width           =   2415
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
      Left            =   11520
      TabIndex        =   5
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   -1320
      X2              =   28680
      Y1              =   10920
      Y2              =   10920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   -240
      X2              =   28680
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   -4440
      X2              =   28680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
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
      Height          =   615
      Left            =   25200
      TabIndex        =   28
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER INVOICE"
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
      Left            =   9960
      TabIndex        =   20
      Top             =   360
      Width           =   9975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER ADDRESS"
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
      Left            =   15720
      TabIndex        =   17
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER PHNO"
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
      Left            =   15720
      TabIndex        =   14
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BUYING PRICE"
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
      Left            =   15720
      TabIndex        =   13
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
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
      Left            =   15720
      TabIndex        =   10
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label6 
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
      Left            =   15720
      TabIndex        =   9
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label5 
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
      Height          =   855
      Left            =   7560
      TabIndex        =   4
      Top             =   7200
      Width           =   3255
   End
   Begin VB.Label Label4 
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
      Left            =   7560
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER NAME"
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
      Left            =   7560
      TabIndex        =   2
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER ID"
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
      Left            =   7560
      TabIndex        =   1
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVOICE ID"
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
      Left            =   7560
      TabIndex        =   0
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "sinvoicee"
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
rs.Open "sinvoice", con
'rs.AddNew
s1 = "insert into sinvoice values('" + Text1.Text + "', '" + Text2.Text + "', '" + Text3.Text + "' ,'" + Text4.Text + "','" + Text5.Text + "' , '" + Text6.Text + "', '" + Text7.Text + "', '" + Text8.Text + "' ,'" + Text9.Text + "','" + Text10.Text + "' )"
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
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""

End Sub

Private Sub Command3_Click()
'update
'con.Open
s1 = "update sinvoice set sid = " & Text2.Text & " , sname = '" & Text3.Text & "' , mid = " & Text4.Text & " , mname = '" & Text5.Text & "' , brand = '" & Text6.Text & "' , quantity = " & Text7.Text & " , bprice = " & Text8.Text & " , sphno = " & Text9.Text & " , sadd = '" & Text10.Text & "'  where invoiceid = " & Text1.Text & " "
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
s1 = "delete from sinvoice where sid=" & str & ""
con.Execute s1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
rs.MoveFirst
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(7)
Text9.Text = rs.Fields(8)
Text10.Text = rs.Fields(9)

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
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(7)
Text9.Text = rs.Fields(8)
Text10.Text = rs.Fields(9)

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
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(7)
Text9.Text = rs.Fields(8)
Text10.Text = rs.Fields(9)

End If

End Sub

Private Sub Command7_Click()
Dim str
str = Val(InputBox("enter the id to search"))
Set rs = New ADODB.Recordset
rs.Open "select * from sinvoice where sid = " & str & "", con
If rs.EOF Then
MsgBox ("not found")
Else
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(7)
Text9.Text = rs.Fields(8)
Text10.Text = rs.Fields(9)

End If


End Sub

Private Sub Command8_Click()
DataReport6.Show
End Sub

Private Sub Form_Load()


Timer1_Timer

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\pharm.mdb"
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\sinvoice.mdb"
con.Open
Set rs = New ADODB.Recordset

rs.CursorLocation = adUseClient
rs.Open "select * from sinvoice", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs

Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(7)
Text9.Text = rs.Fields(8)
Text10.Text = rs.Fields(9)

'Set Text1.DataSource = rs
'Text1.DataField = "sid"
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
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text2_LostFocus()
Dim str As String
str = Text2.Text
Set rs = New ADODB.Recordset

rs.Open "select sname,sphno,sadd from supplier where sid = " & str & " ", con
Text3.Text = rs!sname
Text9.Text = rs!sphno
Text10.Text = rs!sadd

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

Private Sub Text4_LostFocus()
Dim str As String
str = Text4.Text
Set rs = New ADODB.Recordset

rs.Open "select mname,brand,b_price from med where mid = " & str & " ", con
Text5.Text = rs!mname
Text6.Text = rs!brand
Text8.Text = rs!b_price


End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 32 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or KeyAscii >= 123 Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER ALPHABETS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Timer1_Timer()
Label12.Caption = Time
End Sub
