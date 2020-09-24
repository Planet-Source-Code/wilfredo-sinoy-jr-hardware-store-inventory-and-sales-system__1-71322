VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSales 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Information"
   ClientHeight    =   8565
   ClientLeft      =   2580
   ClientTop       =   1590
   ClientWidth     =   11070
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
   ScaleHeight     =   8565
   ScaleWidth      =   11070
   Begin MSAdodcLib.Adodc AdoAllSales 
      Height          =   330
      Left            =   600
      Top             =   8160
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13150
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "General Sales"
      TabPicture(0)   =   "frmSales.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Grid1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdReport1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNext1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrevious1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Collectibles"
      TabPicture(1)   =   "frmSales.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Shape11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Shape12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Grid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdReport2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdNext2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdPrevious2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Paid Purchases"
      TabPicture(2)   =   "frmSales.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Shape4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Shape13"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Shape14"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Grid3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdReport3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdNext3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdPrevious3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         Caption         =   "Customer Information:"
         ForeColor       =   &H8000000E&
         Height          =   3015
         Left            =   120
         TabIndex        =   52
         Top             =   480
         Width           =   10575
         Begin VB.TextBox txtTotalCost3 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtDeliveryStatus3 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtDeliveryDate3 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   2280
            Width           =   3735
         End
         Begin VB.TextBox txtDatePurchased3 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   1800
            Width           =   3735
         End
         Begin VB.TextBox txtCustomerNumber3 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   840
            Width           =   3735
         End
         Begin VB.TextBox txtSalesNumber3 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtCustomerName3 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtAmountPaid3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtBalance3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1800
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo txtSearch3 
            Bindings        =   "frmSales.frx":0054
            Height          =   360
            Left            =   8040
            TabIndex        =   62
            Top             =   2400
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "SalesNumber"
            Text            =   ""
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cost:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   72
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Paid:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   71
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Status:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   70
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Date:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   69
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Purchased:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Number:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Number:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   65
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   64
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label21 
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
            TabIndex        =   63
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000004&
            Height          =   615
            Left            =   5880
            Top             =   2280
            Width           =   4575
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         Caption         =   "Customer Information:"
         ForeColor       =   &H8000000E&
         Height          =   3015
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   10575
         Begin VB.TextBox txtTotalCost2 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtDeliveryStatus2 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtDeliveryDate2 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   2280
            Width           =   3735
         End
         Begin VB.TextBox txtDatePurchased2 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1800
            Width           =   3735
         End
         Begin VB.TextBox txtCustomerNumber2 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   840
            Width           =   3735
         End
         Begin VB.TextBox txtSalesNumber2 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtCustomerName2 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtAmountPaid2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtBalance2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1800
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo txtSearch2 
            Bindings        =   "frmSales.frx":006E
            Height          =   360
            Left            =   8040
            TabIndex        =   37
            Top             =   2400
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "SalesNumber"
            Text            =   ""
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cost:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   47
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Paid:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   46
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Status:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   45
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Date:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Purchased:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Number:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Number:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   39
            Top             =   1920
            Width           =   1815
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
            TabIndex        =   38
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000004&
            Height          =   615
            Left            =   5880
            Top             =   2280
            Width           =   4575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Caption         =   "Customer Information:"
         ForeColor       =   &H8000000E&
         Height          =   3015
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   10575
         Begin VB.TextBox txtBalance1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox txtAmmountPaid1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtCustomerName1 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtsalesnumber1 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtCustomerNumber1 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   840
            Width           =   3735
         End
         Begin VB.TextBox txtdatePurchased1 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1800
            Width           =   3735
         End
         Begin VB.TextBox txtdeliveryDate1 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2280
            Width           =   3735
         End
         Begin VB.TextBox txtDeliveryStatus1 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txttotalCost1 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   840
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo txtsearch1 
            Bindings        =   "frmSales.frx":0088
            Height          =   360
            Left            =   7800
            TabIndex        =   24
            Top             =   2400
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "SalesNumber"
            Text            =   ""
         End
         Begin VB.Label Label10 
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
            TabIndex        =   25
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   20
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Number:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Number:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Purchased:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Date:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Status:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Paid:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   13
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cost:"
            ForeColor       =   &H80000014&
            Height          =   375
            Left            =   6120
            TabIndex        =   12
            Top             =   960
            Width           =   1455
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000004&
            Height          =   615
            Left            =   5880
            Top             =   2280
            Width           =   4575
         End
      End
      Begin lvButton.lvButtons_H cmdPrevious1 
         Height          =   375
         Left            =   -74760
         TabIndex        =   21
         Top             =   3720
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
      Begin lvButton.lvButtons_H cmdNext1 
         Height          =   375
         Left            =   -72000
         TabIndex        =   22
         Top             =   3720
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
      Begin lvButton.lvButtons_H cmdReport1 
         Height          =   375
         Left            =   -68760
         TabIndex        =   23
         Top             =   3720
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
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   26
         Top             =   4440
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
      Begin lvButton.lvButtons_H cmdPrevious2 
         Height          =   375
         Left            =   -74760
         TabIndex        =   48
         Top             =   3720
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
      Begin lvButton.lvButtons_H cmdNext2 
         Height          =   375
         Left            =   -72000
         TabIndex        =   49
         Top             =   3720
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
      Begin lvButton.lvButtons_H cmdReport2 
         Height          =   375
         Left            =   -68760
         TabIndex        =   50
         Top             =   3720
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
      Begin MSDataGridLib.DataGrid Grid2 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   51
         Top             =   4440
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
      Begin lvButton.lvButtons_H cmdPrevious3 
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   3720
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
      Begin lvButton.lvButtons_H cmdNext3 
         Height          =   375
         Left            =   3000
         TabIndex        =   74
         Top             =   3720
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
      Begin lvButton.lvButtons_H cmdReport3 
         Height          =   375
         Left            =   6240
         TabIndex        =   75
         Top             =   3720
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
      Begin MSDataGridLib.DataGrid Grid3 
         Height          =   2775
         Left            =   240
         TabIndex        =   76
         Top             =   4440
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
      Begin VB.Shape Shape14 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   3015
         Left            =   120
         Top             =   4320
         Width           =   10575
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   615
         Left            =   120
         Top             =   3600
         Width           =   10575
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   3015
         Left            =   -74880
         Top             =   4320
         Width           =   10575
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   615
         Left            =   -74880
         Top             =   3600
         Width           =   10575
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   3015
         Left            =   -74880
         Top             =   4320
         Width           =   10575
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   615
         Left            =   -74880
         Top             =   3600
         Width           =   10575
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   7095
         Left            =   -75000
         Top             =   360
         Width           =   10815
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   7095
         Left            =   0
         Top             =   360
         Width           =   10815
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         Height          =   7095
         Left            =   -75000
         Top             =   360
         Width           =   10815
      End
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   8040
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
   Begin MSAdodcLib.Adodc AdoFullSales 
      Height          =   330
      Left            =   2760
      Top             =   7920
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
   Begin MSAdodcLib.Adodc AdoBalances 
      Height          =   330
      Left            =   1560
      Top             =   8040
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
   Begin MSAdodcLib.Adodc AdoSalesInfoAll 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc AdoFullyPaidSalesInfo 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc AdoSalesInfoBalance 
      Height          =   330
      Left            =   0
      Top             =   0
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
      Left            =   120
      Top             =   7920
      Width           =   10815
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
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNext1_Click()
Me.AdoAllSales.Recordset.MoveNext
If Me.AdoAllSales.Recordset.EOF Then
Me.AdoAllSales.Recordset.MovePrevious
End If
settextAll
Call SQLDB3(Me.AdoSalesInfoAll, "Select * from SalesInfo where SalesNumber='" & Me.txtsalesnumber1.Text & "'")
setgrid1
End Sub

Private Sub cmdNext2_Click()
Me.AdoBalances.Recordset.MoveNext
If Me.AdoBalances.Recordset.EOF Then
Me.AdoBalances.Recordset.MovePrevious
End If
settextBalance
Call SQLDB4(Me.AdoSalesInfoBalance, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber2.Text & "'")
setgrid2
End Sub

Private Sub cmdNext3_Click()
Me.AdoFullSales.Recordset.MoveNext
If Me.AdoFullSales.Recordset.EOF Then
Me.AdoFullSales.Recordset.MovePrevious
End If
settextFull
Call SQLDB5(Me.AdoFullyPaidSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber3.Text & "'")
setgrid3
End Sub

Private Sub cmdPrevious1_Click()
Me.AdoAllSales.Recordset.MovePrevious
If Me.AdoAllSales.Recordset.BOF Then
Me.AdoAllSales.Recordset.MoveNext
End If
settextAll
Call SQLDB3(Me.AdoSalesInfoAll, "Select * from SalesInfo where SalesNumber='" & Me.txtsalesnumber1.Text & "'")
setgrid1
End Sub

Private Sub cmdPrevious2_Click()
Me.AdoBalances.Recordset.MovePrevious
If Me.AdoBalances.Recordset.BOF Then
Me.AdoBalances.Recordset.MoveNext
End If
settextBalance
Call SQLDB4(Me.AdoSalesInfoBalance, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber2.Text & "'")
setgrid2
End Sub

Private Sub cmdPrevious3_Click()
Me.AdoFullSales.Recordset.MovePrevious
If Me.AdoFullSales.Recordset.BOF Then
Me.AdoFullSales.Recordset.MoveNext
End If
settextFull
Call SQLDB5(Me.AdoFullyPaidSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber3.Text & "'")
setgrid3
End Sub

Private Sub cmdReport1_Click()
'Set rptSalesPersonal.DataSource = Me.AdoAllSales
rptSales1.Caption = "General Sales"
rptSales1.Sections("Section2").Controls("lblHeader").Caption = "General Sales"
rptSales1.Refresh
rptSales1.Show vbModal
End Sub

Private Sub cmdReport2_Click()
'Set rptSalesPersonal.DataSource = Me.AdoBalances
rptBalance.Caption = "Collectables"
rptBalance.Sections("Section2").Controls("lblHeader").Caption = "Collectibles"
rptBalance.Refresh
rptBalance.Show vbModal
End Sub

Private Sub cmdReport3_Click()
'Set rptSalesPersonal.DataSource = Me.AdoFullSales
rptPaid.Caption = "Paid Purchases"
rptPaid.Sections("Section2").Controls("lblHeader").Caption = "Paid Purchases"
rptPaid.Refresh
rptPaid.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SQLDB(Me.AdoAllSales, "Select * from Sales Order by SalesNumber")
Call SQLDB1(Me.AdoBalances, "Select * from Sales where Balance >0 order by SalesNumber")
Call SQLDB2(Me.AdoFullSales, "Select * from Sales where Balance <=0 order by SalesNumber")
Me.AdoAllSales.Refresh
Me.AdoBalances.Refresh
Me.AdoFullSales.Refresh
settextAll
settextBalance
settextFull
Call SQLDB3(Me.AdoSalesInfoAll, "Select * from SalesInfo where SalesNumber='" & Me.txtsalesnumber1.Text & "'")
Call SQLDB4(Me.AdoSalesInfoBalance, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber2.Text & "'")
Call SQLDB5(Me.AdoFullyPaidSalesInfo, "Select * from SalesInfo where SalesNumber='" & Me.txtSalesNumber3.Text & "'")
If Me.AdoAllSales.Recordset.RecordCount = 0 Then
Me.cmdNext1.Enabled = False
Me.cmdPrevious1.Enabled = False
Me.cmdReport1.Enabled = False
End If
If Me.AdoBalances.Recordset.RecordCount = 0 Then
Me.cmdNext2.Enabled = False
Me.cmdPrevious2.Enabled = False
Me.cmdReport2.Enabled = False
End If
If Me.AdoFullSales.Recordset.RecordCount = 0 Then
Me.cmdNext3.Enabled = False
Me.cmdPrevious3.Enabled = False
Me.cmdReport3.Enabled = False
End If
setgrid1
setgrid2
setgrid3
End Sub
Private Sub settextAll()
On Error Resume Next
Me.txtAmmountPaid1.Text = Format(Me.AdoAllSales.Recordset.Fields("AmountPaid"), "##0.00")
Me.txtBalance1.Text = Format(Me.AdoAllSales.Recordset.Fields("balance"), "##0.00")
Me.txtCustomerName1.Text = Me.AdoAllSales.Recordset.Fields("CustomerName")
Me.txtCustomerNumber1.Text = Me.AdoAllSales.Recordset.Fields("CustomerNumber")
Me.txtdatePurchased1.Text = Me.AdoAllSales.Recordset.Fields("DatePurchased")
Me.txtdeliveryDate1.Text = Me.AdoAllSales.Recordset.Fields("DeliveryDate")
Me.txtDeliveryStatus1.Text = Me.AdoAllSales.Recordset.Fields("DeliveryStatus")
Me.txtsalesnumber1.Text = Me.AdoAllSales.Recordset.Fields("SalesNumber")
Me.txttotalCost1.Text = Format(Me.AdoAllSales.Recordset.Fields("TotalCost"), "##0.00")
End Sub
Private Sub settextBalance()
On Error Resume Next
Me.txtAmountPaid2.Text = Format(Me.AdoBalances.Recordset.Fields("AmountPaid"), "##0.00")
Me.txtBalance2.Text = Format(Me.AdoBalances.Recordset.Fields("balance"), "##0.00")
Me.txtCustomerName2.Text = Me.AdoBalances.Recordset.Fields("CustomerName")
Me.txtCustomerNumber2.Text = Me.AdoBalances.Recordset.Fields("CustomerNumber")
Me.txtDatePurchased2.Text = Me.AdoBalances.Recordset.Fields("DatePurchased")
Me.txtDeliveryDate2.Text = Me.AdoBalances.Recordset.Fields("DeliveryDate")
Me.txtDeliveryStatus2.Text = Me.AdoBalances.Recordset.Fields("DeliveryStatus")
Me.txtSalesNumber2.Text = Me.AdoBalances.Recordset.Fields("SalesNumber")
Me.txtTotalCost2.Text = Format(Me.AdoBalances.Recordset.Fields("TotalCost"), "##0.00")
End Sub
Private Sub settextFull()
On Error Resume Next
Me.txtAmountPaid3.Text = Format(Me.AdoFullSales.Recordset.Fields("AmountPaid"), "##0.00")
Me.txtBalance3.Text = Format(Me.AdoFullSales.Recordset.Fields("balance"), "##0.00")
Me.txtCustomerName3.Text = Me.AdoFullSales.Recordset.Fields("CustomerName")
Me.txtCustomerNumber3.Text = Me.AdoFullSales.Recordset.Fields("CustomerNumber")
Me.txtDatePurchased3.Text = Me.AdoFullSales.Recordset.Fields("DatePurchased")
Me.txtDeliveryDate3.Text = Me.AdoFullSales.Recordset.Fields("DeliveryDate")
Me.txtDeliveryStatus3.Text = Me.AdoFullSales.Recordset.Fields("DeliveryStatus")
Me.txtSalesNumber3.Text = Me.AdoFullSales.Recordset.Fields("SalesNumber")
Me.txtTotalCost3.Text = Format(Me.AdoFullSales.Recordset.Fields("TotalCost"), "##0.00")
End Sub
Private Sub setgrid1()
Set Grid1.DataSource = Me.AdoSalesInfoAll
    With Grid1
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
Private Sub setgrid2()
Set Grid2.DataSource = Me.AdoSalesInfoBalance
    With Grid2
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

Private Sub setgrid3()
Set Grid3.DataSource = Me.AdoFullyPaidSalesInfo
    With Grid3
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

Private Sub txtsearch1_Change()
On Error Resume Next
Me.AdoAllSales.Refresh
Me.AdoAllSales.Recordset.Find ("SalesNumber = '" & Me.txtsearch1.Text & "'")
Call settextAll
Call setgrid1
End Sub

Private Sub txtsearch1_Click(Area As Integer)
On Error Resume Next
Me.AdoAllSales.Refresh
Me.AdoAllSales.Recordset.Find ("SalesNumber = '" & Me.txtsearch1.Text & "'")
Call settextAll
Call setgrid1
End Sub

Private Sub txtsearch1_KeyPress(KeyAscii As Integer)
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

Private Sub txtSearch2_Change()
On Error Resume Next
Me.AdoBalances.Refresh
Me.AdoBalances.Recordset.Find ("SalesNumber = '" & Me.txtSearch2.Text & "'")
Call settextBalance
Call setgrid2
End Sub

Private Sub txtSearch2_Click(Area As Integer)
On Error Resume Next
Me.AdoBalances.Refresh
Me.AdoBalances.Recordset.Find ("SalesNumber = '" & Me.txtSearch2.Text & "'")
Call settextBalance
Call setgrid2
End Sub

Private Sub txtSearch2_KeyPress(KeyAscii As Integer)

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

Private Sub txtSearch3_Change()
On Error Resume Next
Me.AdoFullSales.Refresh
Me.AdoFullSales.Recordset.Find ("SalesNumber = '" & Me.txtSearch3.Text & "'")
Call settextFull
Call setgrid3
End Sub

Private Sub txtSearch3_Click(Area As Integer)
On Error Resume Next
Me.AdoFullSales.Refresh
Me.AdoFullSales.Recordset.Find ("SalesNumber = '" & Me.txtSearch3.Text & "'")
Call settextFull
Call setgrid3
End Sub

Private Sub txtSearch3_KeyPress(KeyAscii As Integer)
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
