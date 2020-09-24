VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDelivered 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivered Sales Information"
   ClientHeight    =   7725
   ClientLeft      =   1815
   ClientTop       =   1410
   ClientWidth     =   11010
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
   ScaleHeight     =   7725
   ScaleWidth      =   11010
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Sales Personal Information:"
      ForeColor       =   &H8000000E&
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10575
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtAmountPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtCustomerName 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtSalesNumber 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtCustomerNumber 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtDatePurchased 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtDeliveryDate 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtTotalCost 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo txtSearch 
         Bindings        =   "frmDelivered.frx":0000
         Height          =   360
         Left            =   8040
         TabIndex        =   9
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "SalesNumber"
         Text            =   ""
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Sales No.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   6000
         TabIndex        =   18
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   6120
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Number:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Number:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purchased:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   6120
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cost:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   615
         Left            =   5880
         Top             =   2040
         Width           =   4575
      End
   End
   Begin lvButton.lvButtons_H cmdPrevious 
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "Previous"
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
   Begin lvButton.lvButtons_H cmdNext 
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      Caption         =   "Next"
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
   Begin lvButton.lvButtons_H cmdReport 
      Height          =   375
      Left            =   6360
      TabIndex        =   21
      Top             =   3360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "Report"
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
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2775
      Left            =   360
      TabIndex        =   22
      Top             =   4080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   4210752
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sales Item Information"
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
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   7440
      TabIndex        =   23
      Top             =   7200
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc AdoSales 
      Height          =   330
      Left            =   360
      Top             =   7200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc AdoSalesInfo 
      Height          =   330
      Left            =   1560
      Top             =   7200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc1"
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
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   7080
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   11055
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   3240
      Width           =   10575
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   3015
      Left            =   240
      Top             =   3960
      Width           =   10575
   End
End
Attribute VB_Name = "frmDelivered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
Me.AdoSales.Recordset.MoveNext
If Me.AdoSales.Recordset.EOF Then
Me.AdoSales.Recordset.MovePrevious
End If
settext
Call SQLDB5(Me.AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
setgrid
End Sub


Private Sub cmdPrevious_Click()
Me.AdoSales.Recordset.MovePrevious
If Me.AdoSales.Recordset.BOF Then
Me.AdoSales.Recordset.MoveNext
End If
settext
Call SQLDB5(Me.AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
setgrid
End Sub

Private Sub cmdReport_Click()
Call SQLDB(frmSales.AdoSalesInfoAll, "Select * from qrySales where DeliveryStatus='" & "Delivered" & "'")
Set rptSales.DataSource = frmSales.AdoSalesInfoAll
rptSales.Sections("Section2").Controls("lblHeader").Caption = "List Of Delivered Items"
rptSales.Refresh
rptSales.Caption = "Undelivered Report"
rptSales.Show vbModal

End Sub

Private Sub Form_Load()
On Error Resume Next
Call SQLDB2(Me.AdoSales, "Select * from Sales where DeliveryStatus ='" & "Delivered" & "' order by SalesNumber")
Me.AdoSales.Refresh
settext
Call SQLDB5(Me.AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
If Me.AdoSales.Recordset.RecordCount = 0 Then
Me.cmdNext.Enabled = False
Me.cmdPrevious.Enabled = False
Me.cmdReport.Enabled = False
End If
setgrid
End Sub
Private Sub settext()
On Error Resume Next
Me.txtAmountPaid.Text = Format(Me.AdoSales.Recordset.Fields("AmountPaid"), "##0.00")
Me.txtBalance.Text = Format(Me.AdoSales.Recordset.Fields("balance"), "##0.00")
Me.txtCustomerName.Text = Me.AdoSales.Recordset.Fields("CustomerName")
Me.txtCustomerNumber.Text = Me.AdoSales.Recordset.Fields("CustomerNumber")
Me.txtDatePurchased.Text = Me.AdoSales.Recordset.Fields("DatePurchased")
Me.txtDeliveryDate.Text = Me.AdoSales.Recordset.Fields("DeliveryDate")
Me.txtSalesNumber.Text = Me.AdoSales.Recordset.Fields("SalesNumber")
Me.txtTotalCost.Text = Format(Me.AdoSales.Recordset.Fields("TotalCost"), "##0.00")
End Sub
Private Sub setgrid()
Set Grid.DataSource = Me.AdoSalesInfo
    With Grid
        .Columns(0).DataField = "SalesNumber"
        .Columns(0).Caption = "Sales Number"
        .Columns(0).Width = 1500
                
        .Columns(1).DataField = "Date"
        .Columns(1).Caption = "Date"
        .Columns(1).Width = 1000
        .Columns(1).Visible = False
        
        .Columns(2).DataField = "ItemNumber"
        .Columns(2).Caption = "Item Number"
        .Columns(2).Width = 1500
        .Columns(2).Visible = False
        
        .Columns(3).DataField = "ItemName"
        .Columns(3).Caption = "Item Name"
        .Columns(3).Width = 2000
        
        .Columns(4).DataField = "Qty"
        .Columns(4).Caption = "Qty"
        .Columns(4).Width = 1000
        
        .Columns(5).DataField = "UnitPrice"
        .Columns(5).Caption = "UnitPrice"
        .Columns(5).Width = 1500
        .Columns(5).NumberFormat = "##0.00"
        
        .Columns(6).DataField = "TotalPrice"
        .Columns(6).Caption = "Total Price"
        .Columns(6).Width = 2000
        .Columns(6).NumberFormat = "##0.00"
    End With
End Sub

Private Sub Grid_Click()
settext
End Sub

Private Sub txtSearch_Change()
On Error Resume Next
Me.AdoSales.Refresh
Me.AdoSales.Recordset.Find ("SalesNumber = '" & Me.txtSearch.Text & "'")
Call settext
Call SQLDB5(Me.AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
Call setgrid
End Sub

Private Sub txtSearch_Click(Area As Integer)
On Error Resume Next
Me.AdoSales.Refresh
Me.AdoSales.Recordset.Find ("SalesNumber = '" & Me.txtSearch.Text & "'")
Call settext
Call SQLDB5(Me.AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
Call setgrid
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
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
