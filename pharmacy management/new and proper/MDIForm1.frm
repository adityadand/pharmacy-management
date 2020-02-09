VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   18960
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu home 
      Caption         =   "HOME"
   End
   Begin VB.Menu customer 
      Caption         =   "CUSTOMER"
   End
   Begin VB.Menu doctor 
      Caption         =   "DOCTOR"
   End
   Begin VB.Menu med 
      Caption         =   "MEDICINE"
   End
   Begin VB.Menu supplier 
      Caption         =   "SUPPLIER"
   End
   Begin VB.Menu stock 
      Caption         =   "STOCK"
   End
   Begin VB.Menu ORDER 
      Caption         =   "ORDER"
   End
   Begin VB.Menu bill 
      Caption         =   "BILL"
   End
   Begin VB.Menu sinvoice 
      Caption         =   "SINVOICE"
   End
   Begin VB.Menu ABOUTUS 
      Caption         =   "ABOUT US"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ABOUTUS_Click()
about.Show
End Sub

Private Sub bill_Click()
billl.Show

End Sub

Private Sub customer_Click()
customerr.Show
End Sub

Private Sub doctor_Click()
doctorr.Show
End Sub

Private Sub home_Click()

billl.Hide
customerr.Hide
doctorr.Hide
medd.Hide
orderr.Hide
sinvoicee.Hide
stockk.Hide
supplierr.Hide
MDIForm1.Show


End Sub

Private Sub master_Click()
masterr.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
ans = MsgBox("Do you want to quit", vbYesNo, "exit")
If ans = vbYes Then
CreateObject("sapi.SPvoice").speak ("thank you")
Else
Cancel = 1
End If
End Sub

Private Sub med_Click()
medd.Show
End Sub
 
Private Sub sinvoice_Click()
sinvoicee.Show
End Sub

Private Sub stock_Click()
stockk.Show
End Sub

Private Sub supplier_Click()
supplierr.Show
End Sub

Private Sub ORDER_Click()
orderr.Show

End Sub
