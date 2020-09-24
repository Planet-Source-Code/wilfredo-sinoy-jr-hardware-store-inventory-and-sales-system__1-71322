VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmViewOrders 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orders Information"
   ClientHeight    =   7860
   ClientLeft      =   1800
   ClientTop       =   1995
   ClientWidth     =   10800
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
   ScaleHeight     =   7860
   ScaleWidth      =   10800
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox txtInfo1 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox txtInfo2 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   6
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   5
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc AdoOrders 
      Height          =   375
      Left            =   5760
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2655
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4683
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
      Caption         =   "Grid View"
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
   Begin lvButton.lvButtons_H cmdReport 
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   2760
      Width           =   3375
      _ExtentX        =   5953
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
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   7320
      Width           =   3375
      _ExtentX        =   5953
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
   Begin lvButton.lvButtons_H cmdPrev 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   2880
      TabIndex        =   8
      Top             =   3600
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoItems 
      Height          =   375
      Left            =   360
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc AdoSuppliers 
      Height          =   375
      Left            =   1560
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Supplier Name:"
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
      Left            =   6360
      TabIndex        =   16
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "lblTotal"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   3600
      Width           =   4575
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   6600
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   3480
      Width           =   10335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2895
      Left            =   240
      Top             =   4200
      Width           =   10335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Number:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier No.:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   855
      Left            =   6240
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   3015
      Left            =   240
      Top             =   360
      Width           =   10335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   7200
      Width           =   10335
   End
End
Attribute VB_Name = "frmViewOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdNext_Click()
Me.AdoOrders.Recordset.MoveNext
If Me.AdoOrders.Recordset.EOF Then
Me.AdoOrders.Recordset.MovePrevious
End If
Call settext
End Sub
Private Sub cmdPrev_Click()
Me.AdoOrders.Recordset.MovePrevious
If Me.AdoOrders.Recordset.BOF Then
Me.AdoOrders.Recordset.MoveNext
End If
Call settext
End Sub
Private Sub setgrid()
Set Grid.DataSource = AdoOrders

    With Grid
        .Columns(0).DataField = "SupplierNum"
        .Columns(0).Caption = "Supplier Number"
        .Columns(0).Width = 1500
                
        .Columns(1).DataField = "SupplierName"
        .Columns(1).Caption = "Supplier Name"
        .Columns(1).Width = 2000
        
        .Columns(2).DataField = "ItemNumber"
        .Columns(2).Caption = "Item Number"
        .Columns(2).Width = 1500
        
        .Columns(3).DataField = "ItemName"
        .Columns(3).Caption = "Item Name"
        .Columns(3).Width = 2500
                
        .Columns(4).DataField = "Unit"
        .Columns(4).Caption = "Unit"
        .Columns(4).Width = 1500
        
        .Columns(5).DataField = "Qty"
        .Columns(5).Caption = "Qty"
        .Columns(5).Width = 1500
    End With
End Sub
Private Sub settext()
On Error Resume Next
Me.txtInfo1.Text = Me.AdoOrders.Recordset.Fields("SupplierNum")
Me.txtinfo(1).Text = Me.AdoOrders.Recordset.Fields("SupplierName")
Me.txtInfo2.Text = Me.AdoOrders.Recordset.Fields("ItemNumber")
Me.txtinfo(4).Text = Me.AdoOrders.Recordset.Fields("ItemName")
Me.txtinfo(5).Text = Me.AdoOrders.Recordset.Fields("Unit")
Me.txtinfo(6).Text = Me.AdoOrders.Recordset.Fields("Qty")
End Sub

Private Sub cmdReport_Click()
Set rptOrders.DataSource = Me.AdoOrders
rptOrders.Refresh
rptOrders.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SQLDB(AdoOrders, "Select * from OrderInfo Order by SupplierNum")
Call SQLDB1(AdoItems, "Select * from Items")
Call SQLDB2(AdoSuppliers, "Select * from Supplier")
AdoItems.Refresh
AdoOrders.Refresh
AdoSuppliers.Refresh
Call setgrid
Call settext
If Me.AdoOrders.Recordset.RecordCount = 0 Then
Me.cmdNext.Enabled = False
Me.cmdPrev.Enabled = False
Me.cmdReport.Enabled = False
Me.txtSearch.Locked = True
Call setgrid
End If
lblTotal.Caption = "Total Number of Records: " & Me.AdoOrders.Recordset.RecordCount
End Sub

Private Sub Grid_Click()
Call settext
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call settext
End Sub

Private Sub txtSearch_Change()
On Error Resume Next
'Me.AdoOrders.Refresh
'Me.AdoOrders.Recordset.Find ("SupplierNumber = '" & Me.txtSearch.Text & "'")
'Call settext
'Call setgrid
Me.AdoOrders.Refresh
Me.AdoOrders.Recordset.Filter = "SupplierName LIKE '" & LCase(Me.txtSearch.Text) & "*'"
Call settext
Call setgrid
If Me.txtSearch.Text = "" Then
Me.AdoOrders.Refresh
settext
setgrid
End If

End Sub

'Private Sub txtSearch_Click(Area As Integer)
'On Error Resume Next
'Me.AdoOrders.Refresh
'Me.AdoOrders.Recordset.Find ("SupplierNumber = '" & Me.txtSearch.Text & "'")
'Call settext
'Call setgrid
'End Sub

