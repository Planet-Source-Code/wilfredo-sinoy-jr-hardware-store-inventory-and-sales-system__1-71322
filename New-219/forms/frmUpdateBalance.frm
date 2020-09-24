VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUpdateBalance 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Sales"
   ClientHeight    =   6255
   ClientLeft      =   3720
   ClientTop       =   2175
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6600
   Begin VB.TextBox txttotalCost 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4080
      Width           =   4215
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   1680
      TabIndex        =   19
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtDeliveryStatus 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox txtdeliveryDate 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox txtdatePurchased 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox txtCustomerNumber 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtsalesnumber 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtCustomerName 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtAmmountPaid 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5040
      Width           =   4215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   5760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483647
      LockHover       =   1
      cGradient       =   8421504
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdUpdatePayment 
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Payment"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483647
      LockHover       =   1
      cGradient       =   8421504
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Search"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483647
      LockHover       =   1
      cGradient       =   8421504
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSAdodcLib.Adodc AdoSales 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoItems"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdUpdateDelivery 
      Height          =   375
      Left            =   2160
      TabIndex        =   23
      Top             =   5760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Update Delivery"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483647
      LockHover       =   1
      cGradient       =   8421504
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Number:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   360
      Width           =   6375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   5640
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Status:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Purchased:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Number:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Number:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   4455
      Left            =   120
      Top             =   1080
      Width           =   6375
   End
End
Attribute VB_Name = "frmUpdateBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error GoTo NotFound
Dim temp As String
    Me.AdoSales.Refresh
    Me.AdoSales.Recordset.Find ("SalesNumber = '" & UCase(txtSearch.Text) & "'")
    temp = Me.AdoSales.Recordset.Fields(1)
    Me.cmdUpdatePayment.Enabled = True
    Me.cmdUpdateDelivery.Enabled = True
    settext
Exit Sub
NotFound:
     MsgBox "The record you requested could not be found.", vbOKOnly + vbExclamation, "NotFound"
    txtSearch.SetFocus
   SendKeys HiLyt
    Me.cmdUpdatePayment.Enabled = False
    Me.txtAmmountPaid.Text = "0.00"
    Me.txtBalance.Text = "0.00"
    Me.txtCustomerName.Text = ""
    Me.txtCustomerNumber.Text = ""
    Me.txtDatePurchased.Text = ""
    Me.txtDeliveryDate.Text = ""
    Me.txtDeliveryStatus.Text = ""
    Me.txtSalesNumber.Text = ""
    Me.txtTotalCost.Text = "0.00"
End Sub

Private Sub cmdUpdateDelivery_Click()
With frmUpdateDelivery
    .txtDeliverySatus.Text = Me.txtDeliveryStatus.Text
    .DTDelivery.Value = Me.txtDeliveryDate.Text
    .txtSalesNumber.Text = Me.txtSalesNumber.Text
    .Show vbModal
End With
End Sub

Private Sub cmdUpdatePayment_Click()
With frmAddPayment
    .txtSalesNumber.Text = Me.txtSalesNumber.Text
    .txtAmmountPaid.Text = Format(Me.txtAmmountPaid.Text, "##0.00")
    .txtBalance.Text = Format(Me.txtBalance.Text, "##0.00")
    .txtTotalCost.Text = Format(Me.txtTotalCost.Text, "##0.00")
    .txtPayment.Text = "0.00"
    .Show vbModal
End With
End Sub

Private Sub Form_Load()
Call SQLDB(AdoSales, "Select * from Sales")
AdoSales.Refresh
End Sub
 Private Sub settext()
 On Error Resume Next
 Me.txtAmmountPaid.Text = Format(Val(Me.AdoSales.Recordset.Fields("AmountPaid")), "##0.00")
 Me.txtBalance.Text = Format(Val(Me.AdoSales.Recordset.Fields("Balance")), "##0.00")
 Me.txtCustomerName.Text = Me.AdoSales.Recordset.Fields("CustomerName")
 Me.txtCustomerNumber.Text = Me.AdoSales.Recordset.Fields("CustomerNumber")
 Me.txtDatePurchased.Text = Me.AdoSales.Recordset.Fields("DatePurchased")
 Me.txtDeliveryDate.Text = Me.AdoSales.Recordset.Fields("DeliveryDate")
 Me.txtDeliveryStatus.Text = Me.AdoSales.Recordset.Fields("DeliveryStatus")
 Me.txtSalesNumber.Text = Me.AdoSales.Recordset.Fields("SalesNumber")
 Me.txtTotalCost.Text = Format(Val(Me.AdoSales.Recordset.Fields("TotalCost")), "##0.00")
 End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch_Click
End If
If KeyAscii = 8 Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then
        Exit Sub
    End If
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
