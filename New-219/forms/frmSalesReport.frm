VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesReport 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report"
   ClientHeight    =   2085
   ClientLeft      =   5190
   ClientTop       =   3405
   ClientWidth     =   4830
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
   ScaleHeight     =   2085
   ScaleWidth      =   4830
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Select Date Range:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Begin MSComCtl2.DTPicker dtfrom 
         Height          =   345
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   4210752
         CalendarForeColor=   16777215
         CalendarTitleBackColor=   8208173
         CalendarTitleForeColor=   781309
         CalendarTrailingForeColor=   8421504
         Format          =   51183617
         CurrentDate     =   38065
      End
      Begin MSComCtl2.DTPicker DTto 
         Height          =   345
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   4210752
         CalendarForeColor=   16777215
         CalendarTitleBackColor=   8208173
         CalendarTitleForeColor=   781309
         CalendarTrailingForeColor=   8421504
         Format          =   51183617
         CurrentDate     =   38065
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc AdoSales 
      Height          =   375
      Left            =   3840
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
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
   Begin lvButton.lvButtons_H cmdPreview 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Preview"
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
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   -960
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dtfrom_Change()
If DTto.Value < Me.dtfrom.Value Then
            MsgBox "Ending date value cannot be less than Beginning date value. Adjust accordingly.", vbOKOnly + vbExclamation, "Natasha"
            DTto.Value = dtpFrom.Value
            DTto.SetFocus
            Exit Sub
    End If
End Sub

Private Sub Form_Load()
Call SQLDB(AdoSales, "Select * from qrySales")
dtfrom.Value = Date
DTto.Value = Date
End Sub

Private Sub dtTo_Change()
    If DTto.Value < Me.dtfrom.Value Then
            MsgBox "Ending date value cannot be less than Beginning date value. Adjust accordingly.", vbOKOnly + vbExclamation, "Natasha"
            DTto.Value = dtpFrom.Value
            DTto.SetFocus
            Exit Sub
    End If
End Sub


Private Sub cmdPreview_Click()
'On Error Resume Next
        AdoSales.Recordset.Filter = ""
        AdoSales.Recordset.Filter = "Date >= '" & dtfrom.Value & "' and Date <= '" & DTto.Value & "'"
        'Set rptSales.DataSource = AdoSales
        DE.rscmdSales1.Filter = "DatePurchased >= '" & dtfrom.Value & "' and DatePurchased <= '" & DTto.Value & "'"
        rptSales1.Refresh
        rptSales1.Show vbModal
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub




