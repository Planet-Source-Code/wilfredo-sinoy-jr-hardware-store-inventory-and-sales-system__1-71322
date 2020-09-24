VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmViewCustomers 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers Information"
   ClientHeight    =   7125
   ClientLeft      =   2070
   ClientTop       =   2265
   ClientWidth     =   10650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10650
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin lvButton.lvButtons_H cmdPrev 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2880
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
      cFHover         =   0
      cBhover         =   -2147483647
      LockHover       =   3
      cGradient       =   8421504
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdNext 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2880
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
      cFHover         =   0
      cBhover         =   -2147483647
      LockHover       =   3
      cGradient       =   8421504
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSDataListLib.DataCombo txtsearch 
      Bindings        =   "frmViewCustomers.frx":0000
      Height          =   360
      Left            =   8400
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "CustomerNumber"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2655
      Left            =   360
      TabIndex        =   7
      Top             =   3600
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
   Begin MSAdodcLib.Adodc AdoCustomers 
      Height          =   375
      Left            =   5400
      Top             =   6600
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
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   6600
      Width           =   3615
      _ExtentX        =   6376
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
      cFHover         =   0
      cBhover         =   -2147483647
      LockHover       =   3
      cGradient       =   8421504
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdReport 
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   1920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Caption         =   "Report"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   -2147483647
      LockHover       =   3
      cGradient       =   12632256
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
      Width           =   10575
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   6120
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Customer No."
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
      TabIndex        =   14
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2895
      Left            =   240
      Top             =   3480
      Width           =   10215
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "lblTotal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   6120
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2295
      Left            =   240
      Top             =   360
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   6480
      Width           =   10215
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   240
      Top             =   2760
      Width           =   10095
   End
End
Attribute VB_Name = "frmViewCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
Me.AdoCustomers.Recordset.MoveNext
If Me.AdoCustomers.Recordset.EOF Then
Me.AdoCustomers.Recordset.MovePrevious
End If
Call settext
End Sub

Private Sub cmdPrev_Click()
Me.AdoCustomers.Recordset.MovePrevious
If Me.AdoCustomers.Recordset.BOF Then
Me.AdoCustomers.Recordset.MoveNext
End If
Call settext
End Sub

Private Sub cmdReport_Click()
Set rptCustomers.DataSource = Me.AdoCustomers
rptCustomers.Refresh
rptCustomers.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SQLDB(AdoCustomers, "Select * from Customer order by CustomerNumber")
AdoCustomers.Refresh
Call setgrid
Call settext
If Me.AdoCustomers.Recordset.RecordCount = 0 Then
Me.cmdNext.Enabled = False
Me.cmdPrev.Enabled = False
Me.cmdReport.Enabled = False
Call setgrid
End If
lblTotal.Caption = "Total Number of Records: " & Me.AdoCustomers.Recordset.RecordCount
End Sub
Private Sub setgrid()
Set Grid.DataSource = AdoCustomers
    With Grid
        .Columns(0).DataField = "CustomerNumber"
        .Columns(0).Caption = "Customer Number"
        .Columns(0).Width = 2000
                
        .Columns(1).DataField = "CustomerName"
        .Columns(1).Caption = "Customer Name"
        .Columns(1).Width = 2500
        
        .Columns(2).DataField = "Address"
        .Columns(2).Caption = "Address"
        .Columns(2).Width = 3000
        
        .Columns(3).DataField = "ContactNumber"
        .Columns(3).Caption = "Contact Number"
        .Columns(3).Width = 2000
    End With
End Sub
Private Sub settext()
On Error Resume Next
Me.txtinfo(0).Text = Me.AdoCustomers.Recordset.Fields("CustomerNumber")
Me.txtinfo(1).Text = Me.AdoCustomers.Recordset.Fields("CustomerName")
Me.txtinfo(2).Text = Me.AdoCustomers.Recordset.Fields("Address")
Me.txtinfo(3).Text = Me.AdoCustomers.Recordset.Fields("ContactNumber")
End Sub

Private Sub Grid_Click()
Call settext

End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call settext
End Sub

Private Sub txtSearch_Change()
On Error Resume Next
Me.AdoCustomers.Refresh
Me.AdoCustomers.Recordset.Find ("CustomerNumber = '" & Me.txtsearch.Text & "'")
Call settext
Call setgrid
End Sub

Private Sub txtSearch_Click(Area As Integer)
On Error Resume Next
Me.AdoCustomers.Refresh
Me.AdoCustomers.Recordset.Find ("CustomerNumber = '" & Me.txtsearch.Text & "'")
Call settext
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
