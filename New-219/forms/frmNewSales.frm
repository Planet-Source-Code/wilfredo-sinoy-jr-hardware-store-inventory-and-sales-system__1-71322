VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewSales 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Transaction"
   ClientHeight    =   8175
   ClientLeft      =   645
   ClientTop       =   1905
   ClientWidth     =   13740
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
   ScaleHeight     =   8175
   ScaleWidth      =   13740
   Begin MSAdodcLib.Adodc Adoitems 
      Height          =   330
      Left            =   2280
      Top             =   6960
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
   Begin VB.ComboBox txtDeliverySatus 
      Height          =   360
      ItemData        =   "frmNewSales.frx":0000
      Left            =   5880
      List            =   "frmNewSales.frx":000A
      TabIndex        =   20
      Top             =   6240
      Width           =   4215
   End
   Begin MSDataListLib.DataCombo txtCustomerName 
      Bindings        =   "frmNewSales.frx":0026
      Height          =   360
      Left            =   5880
      TabIndex        =   17
      Top             =   4800
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "CustomerName"
      BoundColumn     =   "CustomerName"
      Text            =   ""
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7200
      Width           =   4215
   End
   Begin VB.TextBox txtTotalCost 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtAmountPaid 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6720
      Width           =   4215
   End
   Begin VB.TextBox txtCustomerNumber 
      Height          =   375
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   4215
   End
   Begin VB.TextBox txtSalesNumber 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3840
      Width           =   4215
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   12303
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
      Caption         =   "Availabled Items"
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
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   7560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "Add to Sales"
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
   Begin MSDataGridLib.DataGrid grid2 
      Height          =   2535
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4471
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
      Caption         =   "Saled Items"
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
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   345
      Left            =   5880
      TabIndex        =   18
      Top             =   5280
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   8421504
      CalendarForeColor=   16777215
      CalendarTitleBackColor=   4210752
      CalendarTitleForeColor=   65535
      CalendarTrailingForeColor=   8421504
      Format          =   16580609
      CurrentDate     =   38065
   End
   Begin MSComCtl2.DTPicker DTDelivery 
      Height          =   345
      Left            =   5880
      TabIndex        =   19
      Top             =   5760
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   8421504
      CalendarForeColor=   16777215
      CalendarTitleBackColor=   4210752
      CalendarTitleForeColor=   65535
      CalendarTrailingForeColor=   8421504
      Format          =   16580609
      CurrentDate     =   38065
   End
   Begin lvButton.lvButtons_H cmdReciept 
      Height          =   375
      Left            =   10680
      TabIndex        =   22
      Top             =   6480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "Receipt"
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
   Begin MSAdodcLib.Adodc AdoSalesInfo 
      Height          =   330
      Left            =   4080
      Top             =   3720
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
   Begin MSAdodcLib.Adodc AdoSales 
      Height          =   330
      Left            =   4080
      Top             =   7080
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
   Begin MSAdodcLib.Adodc adoCustomer 
      Height          =   330
      Left            =   4200
      Top             =   4560
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
   Begin VB.TextBox txtPayment 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   6720
      Width           =   4215
   End
   Begin lvButton.lvButtons_H cmdPayment 
      Height          =   375
      Left            =   10680
      TabIndex        =   28
      Top             =   6840
      Width           =   2775
      _ExtentX        =   4895
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
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   375
      Left            =   10680
      TabIndex        =   30
      Top             =   7200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "New"
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
   Begin lvButton.lvButtons_H cmdNewCustomer 
      Height          =   375
      Left            =   10680
      TabIndex        =   21
      Top             =   6120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "New Customers"
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
   Begin lvButton.lvButtons_H cmdCalculator 
      Height          =   375
      Left            =   10680
      TabIndex        =   25
      Top             =   5760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "Calculator"
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
      Left            =   10680
      TabIndex        =   23
      Top             =   7560
      Width           =   2775
      _ExtentX        =   4895
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
   Begin lvButton.lvButtons_H cmdEditItems 
      Height          =   375
      Left            =   4080
      TabIndex        =   31
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Edit Items"
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
   Begin lvButton.lvButtons_H cmdRemove 
      Height          =   375
      Left            =   6480
      TabIndex        =   32
      Top             =   3120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "Remove"
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
   Begin VB.Label lblChange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   6960
      TabIndex        =   29
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
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
      Left            =   6120
      TabIndex        =   27
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewSales.frx":0040
      ForeColor       =   &H80000014&
      Height          =   1455
      Left            =   10800
      TabIndex        =   24
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   1695
      Left            =   10680
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   4455
      Left            =   10560
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   375
      Left            =   3960
      Top             =   7800
      Width           =   6375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost:"
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
      Left            =   9480
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Number:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Number:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Purchased:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Status:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   3975
      Left            =   3960
      Top             =   3720
      Width           =   6375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   3255
      Left            =   3960
      Top             =   360
      Width           =   9615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   7695
      Left            =   240
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmNewSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
Grid_DblClick
End Sub

Private Sub cmdCalculator_Click()
Shell "calc.exe", vbMaximizedFocus
End Sub

Private Sub cmdClose_Click()
'On Error Resume Next
Call SQLDB1(AdoSales, "Select * from Sales")
Me.AdoSales.Refresh
If Val(Me.txtTotalCost.Text) = 0 Then
Unload Me
Else
    If Me.txtCustomerName.Text = "" Then
    MsgBox "Please complete data before closing.", vbInformation, "Information"
    Else
    With Me.AdoSales
    .Refresh
    .Recordset.AddNew
    .Recordset.Fields("SalesNumber") = Me.txtSalesNumber.Text
    .Recordset.Fields("CustomerNumber") = Me.txtCustomerNumber.Text
    .Recordset.Fields("CustomerName") = Me.txtCustomerName.Text
    .Recordset.Fields("Datepurchased") = Me.dtDate.Value
    .Recordset.Fields("DeliveryDate") = Me.DTDelivery.Value
    .Recordset.Fields("DeliveryStatus") = Me.txtDeliverySatus.Text
    .Recordset.Fields("TotalCost") = Me.txtTotalCost.Text
    .Recordset.Fields("AmountPaid") = Me.txtAmountPaid.Text
    .Recordset.Fields("Balance") = Me.txtBalance.Text
    .Recordset.Update
    .Refresh
    Unload Me
    End With
    End If
End If
End Sub

Private Sub cmdEditItems_Click()
'.Recordset.Fields("SalesNumber") = Me.txtinfo(0).Text
'.Recordset.Fields("Date") = Me.txtinfo(1).Text
'.Recordset.Fields("ItemNumber") = Me.txtinfo(2).Text
'.Recordset.Fields("ItemName") = Me.txtinfo(3).Text
'.Recordset.Fields("Qty") = Me.txtinfo(4).Text
'.Recordset.Fields("UnitPrice") = Me.txtinfo(5).Text
'.Recordset.Fields("TotalPrice") = Me.txtinfo(6).Text
'''''''''''''ja lng anay'''''''''''''
Me.AdoItems.Refresh
Me.AdoItems.Recordset.Find ("ItemNumber = '" & Me.AdoSalesInfo.Recordset.Fields("ItemNumber") & "'")
frmEditItems.txtinfo(0).Text = Me.AdoSalesInfo.Recordset.Fields("SalesNumber")
frmEditItems.txtinfo(1).Text = Me.AdoSalesInfo.Recordset.Fields("Date")
frmEditItems.txtinfo(2).Text = Me.AdoSalesInfo.Recordset.Fields("ItemNumber")
frmEditItems.txtinfo(3).Text = Me.AdoSalesInfo.Recordset.Fields("ItemName")
frmEditItems.txtinfo(4).Text = Me.AdoSalesInfo.Recordset.Fields("Qty")
frmEditItems.txtinfo(5).Text = Me.AdoSalesInfo.Recordset.Fields("UnitPrice")
frmEditItems.txtinfo(6).Text = Me.AdoSalesInfo.Recordset.Fields("TotalPrice")
frmEditItems.Show vbModal
End Sub

Private Sub cmdNew_Click()
'On Error Resume Next
Call SQLDB1(AdoSales, "Select * from Sales")
Me.AdoSales.Refresh
If Val(Me.txtTotalCost.Text) = 0 Then
Else
    If Me.txtCustomerName.Text = "" Then
    MsgBox "Please complete data before closing.", vbInformation, "Information"
    Else
    With Me.AdoSales
    .Refresh
    .Recordset.AddNew
    .Recordset.Fields("SalesNumber") = Me.txtSalesNumber.Text
    .Recordset.Fields("CustomerNumber") = Me.txtCustomerNumber.Text
    .Recordset.Fields("CustomerName") = Me.txtCustomerName.Text
    .Recordset.Fields("Datepurchased") = Me.dtDate.Value
    .Recordset.Fields("DeliveryDate") = Me.DTDelivery.Value
    .Recordset.Fields("DeliveryStatus") = Me.txtDeliverySatus.Text
    .Recordset.Fields("TotalCost") = Me.txtTotalCost.Text
    .Recordset.Fields("AmountPaid") = Me.txtAmountPaid.Text
    .Recordset.Fields("Balance") = Me.txtBalance.Text
    .Recordset.Update
    .Refresh
    Me.txtCustomerName.Locked = False
    Me.txtCustomerNumber.Locked = False
    Me.txtCustomerName.Text = ""
    Me.txtCustomerNumber.Text = ""
    Me.txtAmountPaid.Text = "0.00"
    Call SQLDB1(frmNewSales.AdoSales, "Select * from Sales")
    Call SQLDB2(frmNewSales.adoCustomer, "Select * from Customer order by CustomerName")
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
    Me.cmdNew.Enabled = False
    Me.cmdNewCustomer.Enabled = True
    Me.cmdPayment.Enabled = False
    Me.cmdReciept.Enabled = False
    Me.cmdEditItems.Enabled = False
    Me.cmdRemove.Enabled = False
    Call SQLDB3(AdoItems, "Select * from Items where RemainingQty > 0")
    Call SQLDB(AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
    AdoSalesInfo.Refresh
    AdoItems.Refresh
    'Me.cmdNew.Enabled = True
    Set Grid2.DataSource = AdoSalesInfo
    Grid2.Columns(1).Visible = False
    Grid2.Columns(2).Visible = False
    Grid2.Columns(5).NumberFormat = "##0.00"
    Grid2.Columns(6).NumberFormat = "##0.00"
    Set Grid.DataSource = AdoItems
    Grid.Columns(0).Visible = False
    Grid.Columns(2).Visible = False
    Grid.Columns(5).Visible = False
    Grid.Columns(6).Visible = True
    Me.txtCustomerName.Locked = False
    Me.txtCustomerName.Locked = False
    lblChange.Caption = "0.00"
    Me.txtCustomerName.Text = ""
    Me.txtCustomerNumber.Text = ""
    Me.txtAmountPaid.Text = "0.00"
    Me.txtCustomerName.Locked = False
    Me.txtCustomerNumber.Locked = False
    Me.txtCustomerName.Text = ""
    Me.txtCustomerNumber.Text = ""
    Me.txtAmountPaid.Visible = True
    'Unload Me
    End With
    End If
End If
End Sub

Private Sub cmdNewCustomer_Click()
With frmNewCustomer
    If Me.adoCustomer.Recordset.RecordCount >= 0 And Me.adoCustomer.Recordset.RecordCount < 9 Then
        .txtinfo(0).Text = "00000" & Me.adoCustomer.Recordset.RecordCount + 1
    End If
    If Me.adoCustomer.Recordset.RecordCount >= 9 And Me.adoCustomer.Recordset.RecordCount < 99 Then
        .txtinfo(0).Text = "0000" & Me.adoCustomer.Recordset.RecordCount + 1
    End If
    If Me.adoCustomer.Recordset.RecordCount >= 99 And Me.adoCustomer.Recordset.RecordCount < 999 Then
        .txtinfo(0).Text = "000" & Me.adoCustomer.Recordset.RecordCount + 1
    End If
    If Me.adoCustomer.Recordset.RecordCount >= 999 And Me.adoCustomer.Recordset.RecordCount < 9999 Then
        .txtinfo(0).Text = "00" & Me.adoCustomer.Recordset.RecordCount + 1
    End If
    If Me.adoCustomer.Recordset.RecordCount >= 9999 And Me.adoCustomer.Recordset.RecordCount < 99999 Then
        .txtinfo(0).Text = "0" & Me.adoCustomer.Recordset.RecordCount + 1
    End If
    If Me.adoCustomer.Recordset.RecordCount >= 99999 And Me.adoCustomer.Recordset.RecordCount > 999999 Then
        .txtinfo(0).Text = Me.adoCustomer.Recordset.RecordCount + 1
    End If
    .Show vbModal
End With
End Sub

Private Sub cmdPayment_Click()
If Me.AdoSalesInfo.Recordset.RecordCount = 0 Then
MsgBox "You cannot proceed to payment because there is no item in your sales item list!!!", vbExclamation, "Error"
Else
frmPayment.txtBalance.Text = Me.txtBalance.Text
frmPayment.txtChange.Text = "0.00"
frmPayment.txtPayment.Text = "0"
frmPayment.Show vbModal
End If
End Sub

Private Sub cmdReciept_Click()
If Me.AdoSalesInfo.Recordset.RecordCount = 0 Then
MsgBox "There is no item!!!", vbExclamation, "Hardware System"
Else
    If Me.txtCustomerName.Text = "" Or Me.txtDeliverySatus.Text = "" Then
    MsgBox "Please supply personal data before you view the receipt.", vbInformation, "Information"
    Else
    Set RptReceipt.DataSource = Me.AdoSalesInfo
    RptReceipt.Refresh
    RptReceipt.Sections("Section2").Controls("lblSalesNumber").Caption = frmNewSales.txtSalesNumber.Text
    RptReceipt.Sections("Section2").Controls("lblTotalCost").Caption = Format(frmNewSales.txtTotalCost.Text, "#,##0.00")
    RptReceipt.Sections("Section2").Controls("lblCustomerNumber").Caption = frmNewSales.txtCustomerNumber.Text
    RptReceipt.Sections("Section2").Controls("lblCustomerName").Caption = frmNewSales.txtCustomerName.Text
    RptReceipt.Sections("Section2").Controls("lblDatePurchased").Caption = frmNewSales.dtDate.Value
    RptReceipt.Sections("Section2").Controls("lblDeliveryDate").Caption = frmNewSales.DTDelivery.Value
    RptReceipt.Sections("Section2").Controls("lblAmountPaid").Caption = Format(frmNewSales.txtPayment.Text, "#,##0.00")
    RptReceipt.Sections("Section2").Controls("lblBalance").Caption = Format(frmNewSales.txtBalance.Text, "#,##0.00")
     RptReceipt.Sections("Section2").Controls("lblChange").Caption = Format(frmNewSales.lblChange.Caption, "#,##0.00")
    cmdNew.Enabled = True
    RptReceipt.Show vbModal
    End If
End If
End Sub

Private Sub cmdRemove_Click()
Dim reply
reply = MsgBox("Are you sure you want to remove this item to sales list?", vbYesNo + vbQuestion, "Confirmation to Remove")
If reply = vbYes Then
 frmNewSales.AdoItems.Refresh
    Call SQLDB3(AdoItems, "Select * from Items")
    frmNewSales.AdoItems.Recordset.Find ("ItemNumber= '" & Me.AdoSalesInfo.Recordset.Fields("ItemNumber") & "'")
    frmNewSales.AdoItems.Recordset.Fields("RemainingQty") = Val(frmNewSales.AdoItems.Recordset.Fields("RemainingQty")) + Me.AdoSalesInfo.Recordset.Fields("Qty")
    frmNewSales.AdoItems.Recordset.Update
    frmNewSales.AdoItems.Refresh
    frmNewSales.txtTotalCost.Text = Format(Val(frmNewSales.txtTotalCost.Text) - Me.AdoSalesInfo.Recordset.Fields("TotalPrice"), "##0.00")
    frmNewSales.txtBalance.Text = Format(Val(frmNewSales.txtBalance.Text) - Me.AdoSalesInfo.Recordset.Fields("TotalPrice"), "##0.00")
    frmNewSales.AdoSalesInfo.Recordset.Delete
    frmNewSales.AdoSalesInfo.Refresh
    Call SQLDB3(AdoItems, "Select * from Items where RemainingQty > 0")
    Call SQLDB(AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
    AdoSalesInfo.Refresh
    AdoItems.Refresh
    'Me.cmdNew.Enabled = True
    Set Grid2.DataSource = AdoSalesInfo
        Grid2.Columns(1).Visible = False
        Grid2.Columns(2).Visible = False
        Grid2.Columns(5).NumberFormat = "##0.00"
        Grid2.Columns(6).NumberFormat = "##0.00"
    Set Grid.DataSource = AdoItems
    Grid.Columns(0).Visible = False
    Grid.Columns(2).Visible = False
    Grid.Columns(5).Visible = False
    If Me.AdoSalesInfo.Recordset.RecordCount = 0 Then
    Me.cmdEditItems.Enabled = False
    Me.cmdRemove.Enabled = False
    Me.cmdPayment.Enabled = False
    End If
Else
End If
End Sub
Private Sub Form_Load()
'On Error Resume Next
Call SQLDB3(AdoItems, "Select * from Items where RemainingQty > 0")
Call SQLDB(AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber.Text & "'")
Call SQLDB2(frmNewSales.adoCustomer, "Select * from Customer Order by CustomerName")
adoCustomer.Refresh
AdoSalesInfo.Refresh
AdoItems.Refresh
'Me.cmdNew.Enabled = True
Set Grid2.DataSource = AdoSalesInfo
    Grid2.Columns(1).Visible = False
    Grid2.Columns(2).Visible = False
    Grid2.Columns(5).NumberFormat = "##0.00"
    Grid2.Columns(6).NumberFormat = "##0.00"
Set Grid.DataSource = AdoItems
Grid.Columns(0).Visible = False
Grid.Columns(2).Visible = False
Grid.Columns(5).Visible = False
Grid.Columns(6).Visible = True
'Call setgrid
'Call settext
'Call setgrid
'lblTotal.Caption = "Total Number of Records: " & Me.AdoSalesInfo.Recordset.RecordCount
End Sub

Private Sub setgrid()
Set Me.Grid2.DataSource = Me.AdoSalesInfo
Grid.Columns(0).Visible = False
Grid.Columns(2).Visible = False
Grid.Columns(3).Visible = False
Grid.Columns(4).Visible = False
Grid.Columns(5).Visible = False
Grid.Columns(6).Visible = True
End Sub

Private Sub Grid_DblClick()
With frmSalesConfirm
    .txtinfo(0).Text = Me.txtSalesNumber.Text
    .txtinfo(1).Text = Me.dtDate.Value
    .txtinfo(2).Text = Me.AdoItems.Recordset.Fields("ItemNumber")
    .txtinfo(3).Text = Me.AdoItems.Recordset.Fields("ItemName")
    .txtinfo(4).Text = "0"
    .txtinfo(5).Text = Me.AdoItems.Recordset.Fields("UnitPrice")
    .txtinfo(6).Text = "0.00"
    If Me.AdoItems.Recordset.Fields("Unit") = "Kilo" Then
    .Label5.Caption = "Qty per Kilo:"
    End If
    .Show vbModal
End With
End Sub

Private Sub lvButtons_H3_Click()
Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub lvButtons_H1_Click()

End Sub

Private Sub txtAmountPaid_Change()
'If Val(Me.txtAmountPaid.Text) > Val(Me.txtTotalCost.Text) Then
'MsgBox "Amount paid is higher than its total cost.", vbExclamation, "Hardware System"
'Me.txtAmountPaid.Text = "0.00"
'Else
'Me.txtBalance.Text = Format(Val(Me.txtTotalCost.Text) - Val(Me.txtAmountPaid.Text), "##0.00")
'End If
End Sub

Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)
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

Private Sub txtCustomerName_Change()
On Error Resume Next
Me.adoCustomer.Refresh
Me.adoCustomer.Recordset.Find ("CustomerName = '" & Me.txtCustomerName.Text & "'")
Me.txtCustomerNumber.Text = Me.adoCustomer.Recordset.Fields("CustomerNumber")
End Sub

Private Sub txtCustomerNumber_KeyPress(KeyAscii As Integer)
KeyAscii = False
End Sub

Private Sub txtDeliverySatus_KeyPress(KeyAscii As Integer)
KeyAscii = False
End Sub

