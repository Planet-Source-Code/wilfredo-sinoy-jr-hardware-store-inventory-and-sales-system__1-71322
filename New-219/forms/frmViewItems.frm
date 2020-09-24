VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmViewItems 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items Information"
   ClientHeight    =   7485
   ClientLeft      =   2070
   ClientTop       =   2175
   ClientWidth     =   10635
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
   ScaleHeight     =   7485
   ScaleWidth      =   10635
   Begin VB.TextBox txtUnit 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtinfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   5
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   6
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin MSDataListLib.DataCombo txtsearch 
      Bindings        =   "frmViewItems.frx":0000
      Height          =   360
      Left            =   7920
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "ItemNumber"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoItems 
      Height          =   375
      Left            =   5760
      Top             =   6960
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
      TabIndex        =   1
      Top             =   3960
      Width           =   9975
      _ExtentX        =   17595
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
      Left            =   6000
      TabIndex        =   8
      Top             =   2400
      Width           =   3975
      _ExtentX        =   7011
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
      TabIndex        =   9
      Top             =   6960
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
   Begin lvButton.lvButtons_H cmdPrev 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3240
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
      TabIndex        =   11
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   10695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Number:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Quantity:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2895
      Left            =   240
      Top             =   3840
      Width           =   10215
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   5760
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "lblTotal"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Search ItemNumber:"
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
      Left            =   5880
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   6840
      Width           =   10215
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   5760
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2655
      Left            =   240
      Top             =   360
      Width           =   10215
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   3120
      Width           =   10215
   End
End
Attribute VB_Name = "frmViewItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
Me.AdoItems.Recordset.MoveNext
If Me.AdoItems.Recordset.EOF Then
Me.AdoItems.Recordset.MovePrevious
End If
Call settext
End Sub

Private Sub cmdPrev_Click()
Me.AdoItems.Recordset.MovePrevious
If Me.AdoItems.Recordset.BOF Then
Me.AdoItems.Recordset.MoveNext
End If
Call settext
End Sub

Private Sub cmdReport_Click()
Set rptItems.DataSource = Me.AdoItems
rptItems.Refresh
rptItems.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SQLDB(AdoItems, "Select * from Items")
AdoItems.Refresh
Call setgrid
Call settext
If Me.AdoItems.Recordset.RecordCount = 0 Then
Me.cmdNext.Enabled = False
Me.cmdPrev.Enabled = False
Me.txtSearch.Locked = True
Call setgrid
End If
lblTotal.Caption = "Total Number of Records: " & Me.AdoItems.Recordset.RecordCount
End Sub
Private Sub setgrid()
Set Grid.DataSource = AdoItems

    With Grid
        .Columns(0).DataField = "ItemNumber"
        .Columns(0).Caption = "ItemNumber"
        .Columns(0).Width = 1500
                
        .Columns(1).DataField = "ItemName"
        .Columns(1).Caption = "Item Name"
        .Columns(1).Width = 2000
        
        .Columns(2).DataField = "Unit"
        .Columns(2).Caption = "Size"
        .Columns(2).Width = 1500
        
        .Columns(3).DataField = "Description"
        .Columns(3).Caption = "Description"
        .Columns(3).Width = 2500
        
        .Columns(4).DataField = "UnitPrice"
        .Columns(4).Caption = "Price"
        .Columns(4).Width = 1500
        .Columns(4).NumberFormat = "##0.00"
        .Columns(4).Alignment = dbgRight
        
        .Columns(5).DataField = "TotalQty"
        .Columns(5).Caption = "Total Qty"
        .Columns(5).Width = 1500
        
        .Columns(6).DataField = "RemainingQty"
        .Columns(6).Caption = "Remaining Qty"
        .Columns(6).Width = 1500
    End With
End Sub
Private Sub settext()
On Error Resume Next
Me.txtinfo(0).Text = Me.AdoItems.Recordset.Fields("ItemNumber")
Me.txtinfo(1).Text = Me.AdoItems.Recordset.Fields("ItemName")
Me.txtUnit.Text = Me.AdoItems.Recordset.Fields("Unit")
Me.txtinfo(3).Text = Me.AdoItems.Recordset.Fields("Description")
Me.txtinfo(4).Text = Format(AdoItems.Recordset.Fields("UnitPrice"), "##0.00")
Me.txtinfo(5).Text = Me.AdoItems.Recordset.Fields("TotalQty")
Me.txtinfo(6).Text = Me.AdoItems.Recordset.Fields("RemainingQty")
End Sub

Private Sub Grid_Click()
Call settext
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call settext
End Sub

Private Sub txtSearch_Change()
On Error Resume Next
Me.AdoItems.Refresh
Me.AdoItems.Recordset.Find ("ItemNumber = '" & Me.txtSearch.Text & "'")
Call settext
End Sub

Private Sub txtSearch_Click(Area As Integer)
On Error Resume Next
Me.AdoItems.Refresh
Me.AdoItems.Recordset.Find ("ItemNumber = '" & Me.txtSearch.Text & "'")
Call settext
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
