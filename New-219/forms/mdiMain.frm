VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00404040&
   Caption         =   "Sales, Order and Inventory System of Noah's Marketing"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11205
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   8265
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "10/31/2008"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:17 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuUser 
         Caption         =   "User Account"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "LogOff"
         Shortcut        =   ^L
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "Customers"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "Suppliers"
         Shortcut        =   ^S
      End
      Begin VB.Menu s35 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItems 
         Caption         =   "Items"
         Shortcut        =   ^I
      End
      Begin VB.Menu s7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuNewSales 
         Caption         =   "New Sales"
         Shortcut        =   ^N
      End
      Begin VB.Menu s9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateSales 
         Caption         =   "Update Sales"
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
      End
      Begin VB.Menu mnuOrder 
         Caption         =   "Orders"
      End
      Begin VB.Menu mnuOrders 
         Caption         =   "Items with Less Remaining Qty"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVCustomers 
         Caption         =   "Customers"
      End
      Begin VB.Menu mnuVSuppliers 
         Caption         =   "Suppliers"
      End
      Begin VB.Menu s11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVOrders 
         Caption         =   "Items with Less Remaining Qty"
      End
      Begin VB.Menu mnuOrder1 
         Caption         =   "Orders"
      End
      Begin VB.Menu s19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVItems 
         Caption         =   "Items"
      End
      Begin VB.Menu s13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVDelivery 
         Caption         =   "Delivery"
         Begin VB.Menu mnuVUndelivered 
            Caption         =   "Undelivered"
         End
         Begin VB.Menu s14 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVdelivered 
            Caption         =   "Delivered"
         End
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuRSales 
         Caption         =   "Sales"
      End
      Begin VB.Menu mnuROrders 
         Caption         =   "Items with Less Remaining Qty"
      End
      Begin VB.Menu mnuOrder2 
         Caption         =   "Orders"
      End
      Begin VB.Menu s32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCustomer 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuRSupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu sep31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRItems 
         Caption         =   "Items"
      End
      Begin VB.Menu s30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRDelivery 
         Caption         =   "Delivery"
         Begin VB.Menu mnuRUndelivered 
            Caption         =   "Undelivered"
         End
         Begin VB.Menu mnuRDelivered 
            Caption         =   "Delivered"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuUserManual 
         Caption         =   "User Manual"
      End
      Begin VB.Menu s20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResearcher 
         Caption         =   "Researcher"
      End
      Begin VB.Menu s21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManagement 
         Caption         =   "Management"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
'With stbMain 'ensures all panels are visible upon loading
 DE.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB\DbaseHardware.mdb;Persist Security Info=False;Jet OLEDB:Database Password= noah"
mdiMain.Width = 15360
       mdiMain.stbMain.Panels(1).Width = (mdiMain.stbMain.Width) - (mdiMain.stbMain.Panels(2).Width + mdiMain.stbMain.Panels(3).Width + mdiMain.stbMain.Panels(4).Width + mdiMain.stbMain.Panels(5).Width + mdiMain.stbMain.Panels(6).Width)
'Timer1_Timer
'Timer2_Timer
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim reply
reply = MsgBox("Do you want to Quit this program?", vbYesNo + vbQuestion, "Confirm Quit")
If reply = vbYes Then
End
End If
error:
Cancel = -1
End Sub

Private Sub mnuCustomers_Click()
frmCustomers.Show vbModal
End Sub

Private Sub mnuItems_Click()
frmFItems.Show vbModal
End Sub

Private Sub mnuLogOff_Click()
Dim reply
reply = MsgBox("Do you really want to LogOff?", vbYesNo + vbQuestion, "Confirm Log-Off")
If reply = vbYes Then
frmLogin.cmdClose.Caption = "Cancel"
frmLogin.Show vbModal
End If
End Sub

Private Sub mnuManagement_Click()
frmManagement.Show vbModal
End Sub

Private Sub mnuNewSales_Click()
'On Error Resume Next
Call SQLDB1(frmNewSales.AdoSales, "Select * from Sales")
Call SQLDB2(frmNewSales.adoCustomer, "Select * from Customer order by CustomerNumber")
    If frmNewSales.AdoSales.Recordset.RecordCount >= 0 And frmNewSales.AdoSales.Recordset.RecordCount < 9 Then
        frmNewSales.txtSalesNumber.Text = "00000" & frmNewSales.AdoSales.Recordset.RecordCount + 1
    End If
    If frmNewSales.AdoSales.Recordset.RecordCount >= 9 And frmNewSales.AdoSales.Recordset.RecordCount < 99 Then
        frmNewSales.txtSalesNumber.Text = "0000" & frmNewSales.AdoSales.Recordset.RecordCount + 1
    End If
    If frmNewSales.AdoSales.Recordset.RecordCount >= 99 And frmNewSales.AdoSales.Recordset.RecordCount < 999 Then
    frmNewSales.txtSalesNumber.Text = "000" & frmNewSales.AdoSales.Recordset.RecordCount + 1
    End If
    If frmNewSales.AdoSales.Recordset.RecordCount >= 999 And frmNewSales.AdoSales.Recordset.RecordCount < 9999 Then
    frmNewSales.txtSalesNumber.Text = "00" & frmNewSales.AdoSales.Recordset.RecordCount + 1
    End If
    If frmNewSales.AdoSales.Recordset.RecordCount >= 9999 And frmNewSales.AdoSales.Recordset.RecordCount < 99999 Then
    frmNewSales.txtSalesNumber.Text = "0" & frmNewSales.AdoSales.Recordset.RecordCount + 1
    End If
    If frmNewSales.AdoSales.Recordset.RecordCount >= 99999 And frmNewSales.AdoSales.Recordset.RecordCount > 999999 Then
    frmNewSales.txtSalesNumber.Text = frmNewSales.AdoSales.Recordset.RecordCount + 1
    End If

frmNewSales.txtAmountPaid.Text = "0.00"
frmNewSales.txtBalance.Text = "0.00"
frmNewSales.txtTotalCost.Text = "0.00"
frmNewSales.txtDeliverySatus.Text = "Undelivered"
frmNewSales.dtDate.Value = Date
frmNewSales.DTDelivery = Date
frmNewSales.Show vbModal
End Sub

Private Sub mnuOrder_Click()
frmOrders.Show vbModal
End Sub

Private Sub mnuOrder1_Click()
frmViewOrders.Show vbModal
End Sub

Private Sub mnuOrder2_Click()
Call SQLDB(frmOrders.AdoOrders, "Select * from OrderInfo Order by SupplierNum")
Set rptOrders.DataSource = frmOrders.AdoOrders
rptOrders.Refresh
rptOrders.Show vbModal
End Sub

Private Sub mnuOrders_Click()
frmItemsReorder.Show vbModal
End Sub

Private Sub mnuQuit_Click()
Dim reply
reply = MsgBox("Do you want to Quit this program?", vbYesNo + vbQuestion, "Confirm Quit")
If reply = vbYes Then
End
End If
End Sub
Private Sub mnuRCustomer_Click()
Call SQLDB(frmCustomers.AdoCustomers, "Select * from Customer order by CustomerNumber")
Set rptCustomers.DataSource = frmCustomers.AdoCustomers
rptCustomers.Refresh
rptCustomers.Show vbModals
End Sub
Private Sub mnuRDelivered_Click()
Call SQLDB(frmSales.AdoSalesInfoAll, "Select * from qrySales where DeliveryStatus='" & "Delivered" & "'")
Set rptSales.DataSource = frmSales.AdoSalesInfoAll
rptSales.Sections("Section2").Controls("lblHeader").Caption = "List Of Delivered Items"
rptSales.Refresh
rptSales.Caption = "Undelivered Report"
rptSales.Show vbModal
End Sub

Private Sub mnuResearcher_Click()
frmResearcher.Show vbModal
End Sub

Private Sub mnuRItems_Click()
Call SQLDB(frmFItems.AdoItems, "Select * from Items order by ItemNumber")
Set rptItems.DataSource = frmFItems.AdoItems
rptItems.Refresh
rptItems.Show vbModal
End Sub

Private Sub mnuROrders_Click()
Call SQLDB(frmItemsReorder.AdoItems, "Select * from Items where RemainingQty <= 15 Order by ItemNumber")
Set rptItems.DataSource = frmItemsReorder.AdoItems
rptItems.Sections("Section2").Controls("lbl").Caption = "Items with Less Remaining Quantity"
'rptItems.Refresh
rptItems.Show vbModal
End Sub

Private Sub mnuRSales_Click()
frmSalesReport.Show vbModal
End Sub

Private Sub mnuRSupplier_Click()
Call SQLDB(frmSupplier.Adosupplier, "Select * from Supplier order by SupplierNumber")
Set rptSuppliers.DataSource = frmSupplier.Adosupplier
rptSuppliers.Refresh
rptSuppliers.Show vbModal
End Sub

Private Sub mnuRUndelivered_Click()
Call SQLDB(frmSales.AdoSalesInfoAll, "Select * from qrySales where DeliveryStatus='" & "Undelivered" & "'")
Set rptSales.DataSource = frmSales.AdoSalesInfoAll
rptSales.Sections("Section2").Controls("lblHeader").Caption = "List Of Undelivered Items"

rptSales.Refresh
rptSales.Caption = "Undelivered Report"
rptSales.Show vbModal
End Sub

Private Sub mnuSales_Click()
frmSales.Show vbModal
End Sub

Private Sub mnuSuppliers_Click()
frmSupplier.Show vbModal
End Sub

Private Sub mnuUpdateSales_Click()
frmUpdateBalance.Show vbModal
End Sub

Private Sub mnuUser_Click()
frmUser.Show vbModal ' the frmUser will show
End Sub

Private Sub s22_Click()

End Sub

Private Sub mnuUserManual_Click()
frmManual.Show vbModal ' the frmManual show
End Sub

Private Sub mnuVCustomers_Click()
frmViewCustomers.Show vbModal ' the frmViewCustomers show
End Sub

Private Sub mnuVdelivered_Click()
frmDelivered.Show 'frmdelivered show
End Sub

Private Sub mnuVItems_Click()
frmViewItems.Show vbModal ' frmviewitems show
End Sub

Private Sub mnuVOrders_Click()
frmItemsReorder.Show vbModal
End Sub

Private Sub mnuVSuppliers_Click()
frmViewSuppliers.Show vbModal ' frmviewsuppliers show
End Sub

Private Sub mnuVUndelivered_Click()
frmUndelivered.Show vbModal ' frmundelivered show
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Picture1_Click()

End Sub
