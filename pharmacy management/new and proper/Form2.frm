VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form billl 
   Caption         =   "bill"
   ClientHeight    =   12315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   25500
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   12315
   ScaleWidth      =   25500
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   23280
      TabIndex        =   37
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   23280
      TabIndex        =   36
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   23280
      TabIndex        =   35
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      TabIndex        =   34
      Top             =   6720
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      TabIndex        =   33
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   21240
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   1455
      Left            =   120
      Picture         =   "Form2.frx":79EE4
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   32
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Show Data Report"
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
      Left            =   4920
      TabIndex        =   30
      Top             =   8760
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      TabIndex        =   29
      Top             =   4560
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   3495
      Left            =   0
      TabIndex        =   28
      Top             =   9960
      Width           =   28695
      _ExtentX        =   50615
      _ExtentY        =   6165
      _Version        =   393216
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   29
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   855
      Left            =   23280
      TabIndex        =   27
      Top             =   5640
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1508
      _Version        =   393216
      Format          =   82313217
      CurrentDate     =   43563
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
      Left            =   11400
      TabIndex        =   25
      Top             =   8760
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
      Left            =   8640
      TabIndex        =   24
      Top             =   8760
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
      Left            =   14160
      TabIndex        =   23
      Top             =   8760
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
      Left            =   16800
      TabIndex        =   22
      Top             =   8760
      Width           =   2055
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
      Left            =   22200
      TabIndex        =   21
      Top             =   8760
      Width           =   1455
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
      Left            =   23760
      TabIndex        =   20
      Top             =   8760
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
      Left            =   19560
      TabIndex        =   19
      Top             =   8760
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Form2.frx":7A845
      Left            =   23280
      List            =   "Form2.frx":7A855
      TabIndex        =   17
      Text            =   "cash"
      Top             =   6720
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      TabIndex        =   16
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15120
      TabIndex        =   15
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   14
      Top             =   6720
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   13
      Top             =   5640
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   12
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   11
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Nunito"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   10
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT METHOD"
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
      Left            =   19680
      TabIndex        =   42
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL_DT"
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
      Left            =   19680
      TabIndex        =   41
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Left            =   19680
      TabIndex        =   40
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "GST"
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
      Left            =   19680
      TabIndex        =   39
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label13 
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
      Left            =   19560
      TabIndex        =   38
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   -240
      X2              =   33240
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   -1800
      X2              =   31320
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   -600
      X2              =   32520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL"
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
      Left            =   14040
      TabIndex        =   26
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   " 00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   21720
      TabIndex        =   18
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SPRICE"
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
      Left            =   11640
      TabIndex        =   9
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MNAME"
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
      Left            =   11640
      TabIndex        =   8
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "DNAME"
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
      Left            =   11640
      TabIndex        =   7
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CADDRESS"
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
      Left            =   11640
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CNAME"
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
      Left            =   11640
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MID"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK ID"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DID"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CID"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL ID"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   2520
      Width           =   3015
   End
End
Attribute VB_Name = "billl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset



Private Sub Command8_Click()
DataReport1.Show
End Sub


 
Private Sub Command1_Click()

'adding a new record
'con.Open
Set rs = New ADODB.Recordset
rs.LockType = adLockOptimistic
rs.Open "bill", con
'rs.AddNew


gst = Val(Text10.Text) * (5 / 100)
Text12.Text = gst

sprice = Val(Text10.Text)
quan = Val(Text11.Text)
total = (sprice + gst) * quan

Text13.Text = total


s1 = "insert into bill values('" + Text1.Text + "', '" + Text2.Text + "', '" + Text3.Text + "' ,'" + Text4.Text + "','" + Text5.Text + "' , '" + Text6.Text + "', '" + Text7.Text + "' , '" + Text8.Text + "' , '" + Text9.Text + "', '" + Text10.Text + "' ,'" + Text11.Text + "','" + Text12.Text + "' , '" + Text13.Text + "', '" + Format(DTPicker1.Value, "DD/MM/YYYY") + "'  ,'" + Combo1.Text + "' )"
's1 = "insert into bill values(11,12,13,'Puja','pune',66,'gg','4/4/2019',600,'cash')"
con.Execute s1
s2 = "update stock set quantity = quantity - " & Text11.Text & " where stockid = " & Text4.Text & ""
con.Execute s2


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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
DTPicker1.Value = Date
Combo1.Text = "cash"

End Sub

Private Sub Command3_Click()
'update
'con.Open

gst = Val(Text10.Text) * (5 / 100)
Text12.Text = gst

sprice = Val(Text10.Text)
quan = Val(Text11.Text)
total = (sprice + gst) * quan

Text13.Text = total

s1 = "update bill set CID = " & Text2.Text & " , DID = " & Text3.Text & " , STOCKID = " & Text4.Text & " , MID = " & Text5.Text & "  , CNAME = '" & Text6.Text & "'  , CADDRESS = '" & Text7.Text & "' , DNAME = '" & Text8.Text & "' , MNAME = '" & Text9.Text & "' , SPRICE = " & Text10.Text & " , QUANTITY = " & Text11.Text & " , GST = " & Text12.Text & " , TOTAL = " & Text13.Text & " , bill_dt = '" & Format(DTPicker1.Value, "DD/MM/YYYY") & "'  , payment_m = '" & Combo1.Text & "'   where bill_id = " & Text1.Text & " "
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
s1 = "delete from bill where bill_id = " & str & " "
con.Execute s1
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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
DTPicker1.Value = Date
Combo1.Text = "cash"

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
Text11.Text = rs.Fields(10)
Text12.Text = rs.Fields(11)
Text13.Text = rs.Fields(12)
DTPicker1.Value = rs.Fields(13)
Combo1.Text = rs.Fields(14)

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
Text11.Text = rs.Fields(10)
Text12.Text = rs.Fields(11)
Text13.Text = rs.Fields(12)
DTPicker1.Value = rs.Fields(13)
Combo1.Text = rs.Fields(14)

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
Text11.Text = rs.Fields(10)
Text12.Text = rs.Fields(11)
Text13.Text = rs.Fields(12)
DTPicker1.Value = rs.Fields(13)
Combo1.Text = rs.Fields(14)


End If

End Sub

Private Sub Command7_Click()
Dim str
str = Val(InputBox("enter the id to search"))
Set rs = New ADODB.Recordset
rs.Open "select * from bill where did = " & str & "", con
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
Text11.Text = rs.Fields(10)
Text12.Text = rs.Fields(11)
Text13.Text = rs.Fields(12)
DTPicker1.Value = rs.Fields(13)
Combo1.Text = rs.Fields(14)


End If


End Sub

Private Sub Form_Load()

Timer1_Timer

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\pharm.mdb"
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Admin\My Documents\bill.mdb"
con.Open
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

rs.Open "select * from bill", con, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs
'Set Text1.DataSource = rs
'Text1.DataField = "bill_id"
'Set Text2.DataSource = rs
'Text2.DataField = "gst_no"
'Set Text3.DataSource = rs
'Text3.DataField = "cid"
'Set Text4.DataSource = rs
'Text4.DataField = "cname"
'Set Text5.DataSource = rs
'Text5.DataField = "caddress"
'Set Text6.DataSource = rs
'Text6.DataField = "did"
'Set Text7.DataSource = rs
'Text7.DataField = "dname"
'Set DTPicker1.DataSource = rs
'DTPicker1.DataField = "bill_dt"
'Set Text8.DataSource = rs
'Text8.DataField = "amount"
'Set Combo1.DataSource = rs
'Combo1.DataField = "payment_m"


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
Text11.Text = rs.Fields(10)
Text12.Text = rs.Fields(11)
Text13.Text = rs.Fields(12)
DTPicker1.Value = rs.Fields(13)
Combo1.Text = rs.Fields(14)


DataGrid1.Refresh



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



Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65) Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER NUMBERS", vbExclamation, "Ërror"
End If
End Sub


Private Sub Text13_KeyPress(KeyAscii As Integer)
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

rs.Open "select cname,caddress from cust where cid = " & str & " ", con
Text6.Text = rs!cname
Text7.Text = rs!caddress

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

rs.Open "select dname from doc where did = " & str & " ", con
Text8.Text = rs!dname

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

Private Sub Text5_LostFocus()
Dim str As String
str = Text5.Text
Set rs = New ADODB.Recordset

rs.Open "select mname,s_price  from med where mid = " & str & " ", con
Text9.Text = rs!mname
Text10.Text = rs!s_price

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 32 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or KeyAscii >= 123 Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER ALPHABETS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 32 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or KeyAscii >= 123 Then
KeyAscii = 0
MsgBox "YOU CAN ONLY ENTER ALPHABETS", vbExclamation, "Ërror"
End If
End Sub

Private Sub Timer1_Timer()
Label11.Caption = Time
End Sub
