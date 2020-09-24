VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSupplier 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier"
   ClientHeight    =   7785
   ClientLeft      =   2160
   ClientTop       =   1995
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10425
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtinfo 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   4095
   End
   Begin lvButton.lvButtons_H cmdPrev 
      Height          =   375
      Left            =   360
      TabIndex        =   0
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
      TabIndex        =   1
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
   Begin MSDataListLib.DataCombo txtsearch 
      Bindings        =   "frmSupplier.frx":0000
      Height          =   360
      Left            =   8040
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "SupplierNumber"
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
   Begin lvButton.lvButtons_H cmdReport 
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   2880
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
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "Add New"
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
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "Edit"
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
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "Save"
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
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   1440
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "Refresh"
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
      Height          =   2655
      Left            =   240
      TabIndex        =   10
      Top             =   4320
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
   Begin MSAdodcLib.Adodc Adosupplier 
      Height          =   375
      Left            =   5640
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
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   7320
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier No.:"
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
      Left            =   360
      TabIndex        =   20
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name:"
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
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   1575
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
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
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
      Left            =   360
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
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
      Left            =   240
      TabIndex        =   12
      Top             =   7320
      Width           =   6615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2895
      Left            =   120
      Top             =   4200
      Width           =   10215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Supplier No."
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
      TabIndex        =   3
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   5760
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   3015
      Left            =   6480
      Top             =   360
      Width           =   3855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   3015
      Left            =   120
      Top             =   360
      Width           =   6135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   735
      Left            =   120
      Top             =   7200
      Width           =   10215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   -120
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim i As Integer
Dim reply
reply = MsgBox("Do you want to add new customers?", vbYesNo + vbQuestion, "Confirm Add")
If reply = vbYes Then
txtinfo(0).Text = ""
txtinfo(1) = ""
txtinfo(2).Text = ""
txtinfo(3) = ""
Call textunlocked
Call cmdDisabled
    If Me.Adosupplier.Recordset.RecordCount >= 0 And Me.Adosupplier.Recordset.RecordCount < 9 Then
        Me.txtinfo(0).Text = "00000" & Me.Adosupplier.Recordset.RecordCount + 1
    End If
    If Me.Adosupplier.Recordset.RecordCount >= 9 And Me.Adosupplier.Recordset.RecordCount < 99 Then
        Me.txtinfo(0).Text = "0000" & Me.Adosupplier.Recordset.RecordCount + 1
    End If
    If Me.Adosupplier.Recordset.RecordCount >= 99 And Me.Adosupplier.Recordset.RecordCount < 999 Then
    Me.txtinfo(0).Text = "000" & Me.Adosupplier.Recordset.RecordCount + 1
    End If
    If Me.Adosupplier.Recordset.RecordCount >= 999 And Me.Adosupplier.Recordset.RecordCount < 9999 Then
    Me.txtinfo(0).Text = "00" & Me.Adosupplier.Recordset.RecordCount + 1
    End If
    If Me.Adosupplier.Recordset.RecordCount >= 9999 And Me.Adosupplier.Recordset.RecordCount < 99999 Then
    Me.txtinfo(0).Text = "0" & Me.Adosupplier.Recordset.RecordCount + 1
    End If
    If Me.Adosupplier.Recordset.RecordCount >= 99999 And Me.Adosupplier.Recordset.RecordCount > 999999 Then
    Me.txtinfo(0).Text = Me.Adosupplier.Recordset.RecordCount + 1
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Call textlocked
Call cmdEnabled
Me.cmdSave.Caption = "Save"
Me.Adosupplier.Refresh
Call setgrid
Call settext
If Me.Adosupplier.Recordset.RecordCount = 0 Then
Me.cmdSave.Caption = "Update"
'Me.cmdAdd.Enabled = False
'Me.cmdAddQty.Enabled = False
Me.cmdEdit.Enabled = False
Me.cmdNext.Enabled = False
Me.cmdPrev.Enabled = False
Me.cmdRefresh.Enabled = False
Me.cmdReport.Enabled = False
Me.txtsearch.Locked = True
End If
End Sub

Private Sub cmdClose_Click()
If Me.cmdAdd.Enabled = False Then
MsgBox "Please finish your transaction before you clase.", vbQuestion, "Confirm"
Else
Unload Me
End If
End Sub

Private Sub cmdEdit_Click()
Call textunlocked
Call cmdDisabled
Me.cmdSave.Caption = "Update"
End Sub

Private Sub cmdNext_Click()
Me.Adosupplier.Recordset.MoveNext
If Me.Adosupplier.Recordset.EOF Then
Me.Adosupplier.Recordset.MovePrevious
End If
Call settext
End Sub

Private Sub cmdPrev_Click()
On Error Resume Next
Me.Adosupplier.Recordset.MovePrevious
If Me.Adosupplier.Recordset.BOF Then
Me.Adosupplier.Recordset.MoveNext
End If
Call settext
End Sub
Private Sub textlocked()
Dim i As Integer
txtinfo(i).Locked = True
txtsearch.Locked = False
End Sub
Private Sub textunlocked()
Me.txtinfo(1).Locked = False
Me.txtinfo(2).Locked = False
Me.txtinfo(3).Locked = False
Me.txtsearch.Locked = True
End Sub
Private Sub cmdDisabled()
Me.cmdAdd.Enabled = False
Me.cmdCancel.Enabled = True
Me.cmdEdit.Enabled = False
Me.cmdNext.Enabled = False
Me.cmdPrev.Enabled = False
Me.cmdRefresh.Enabled = False
Me.cmdReport.Enabled = False
Me.cmdSave.Enabled = True
End Sub
Private Sub cmdEnabled()
Me.cmdAdd.Enabled = True
Me.cmdCancel.Enabled = False
Me.cmdEdit.Enabled = True
Me.cmdNext.Enabled = True
Me.cmdPrev.Enabled = True
Me.cmdRefresh.Enabled = True
Me.cmdReport.Enabled = True
Me.cmdSave.Enabled = False
End Sub

Private Sub cmdRefresh_Click()
Me.Adosupplier.Refresh
Call settext
setgrid
End Sub

Private Sub cmdReport_Click()
Set rptSuppliers.DataSource = Me.Adosupplier
rptSuppliers.Refresh
rptSuppliers.Show vbModal
End Sub

Private Sub cmdSave_Click()
If Me.txtinfo(0).Text = "" Or Me.txtinfo(1).Text = "" Or Me.txtinfo(2).Text = "" Then
MsgBox "Please complete empty box/boxes.", vbExclamation, "Hardware System"
Else
Dim reply
reply = MsgBox("Do you really want to Save/Update this data?", vbYesNoCancel + vbQuestion, "Confirm Save")
If reply = vbYes Then
    If Me.cmdSave.Caption = "Save" Then
    Me.Adosupplier.Recordset.AddNew
    Me.Adosupplier.Recordset.Fields("SupplierNumber") = Me.txtinfo(0).Text
    Me.Adosupplier.Recordset.Fields("SupplierName") = txtinfo(1).Text
    Me.Adosupplier.Recordset.Fields("Address") = txtinfo(2).Text
    Me.Adosupplier.Recordset.Fields("ContactNumber") = txtinfo(3).Text
    Me.Adosupplier.Recordset.Update
    Me.Adosupplier.Refresh
    End If
    If Me.cmdSave.Caption = "Update" Then
    'Me.AdoItems.Recordset.AddNew
    'Me.Adosupplier.Recordset.AddNew
    Me.Adosupplier.Recordset.Fields("SupplierNumber") = Me.txtinfo(0).Text
    Me.Adosupplier.Recordset.Fields("SupplierName") = txtinfo(1).Text
    Me.Adosupplier.Recordset.Fields("Address") = txtinfo(2).Text
    Me.Adosupplier.Recordset.Fields("ContactNumber") = txtinfo(3).Text
    Me.Adosupplier.Recordset.Update
    Me.Adosupplier.Refresh
    End If
    Call textlocked
    Call cmdEnabled
    Call settext
    setgrid
    cmdRefresh_Click
End If
If reply = vbNo Then
cmdCancel_Click
Call textlocked
Call cmdEnabled
Call settext
setgrid
cmdRefresh_Click
End If
End If
lblTotal.Caption = "Total Number of Records: " & Me.Adosupplier.Recordset.RecordCount
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SQLDB(Adosupplier, "Select * from Supplier")
Adosupplier.Refresh
Call setgrid
Call settext
If Me.Adosupplier.Recordset.RecordCount = 0 Then
'Me.cmdAdd.Enabled = False
Me.cmdEdit.Enabled = False
Me.cmdNext.Enabled = False
Me.cmdPrev.Enabled = False
Me.cmdRefresh.Enabled = False
Me.cmdReport.Enabled = False
Me.txtsearch.Locked = True
Call setgrid
End If
lblTotal.Caption = "Total Number of Records: " & Me.Adosupplier.Recordset.RecordCount
End Sub
Private Sub setgrid()
Set Grid.DataSource = Adosupplier
    With Grid
        .Columns(0).DataField = "SupplierNumber"
        .Columns(0).Caption = "Supplier Number"
        .Columns(0).Width = 2000
                
        .Columns(1).DataField = "SupplierName"
        .Columns(1).Caption = "Supplier Name"
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
Me.txtinfo(0).Text = Me.Adosupplier.Recordset.Fields("SupplierNumber")
Me.txtinfo(1).Text = Me.Adosupplier.Recordset.Fields("SupplierName")
Me.txtinfo(2).Text = Me.Adosupplier.Recordset.Fields("Address")
Me.txtinfo(3).Text = Me.Adosupplier.Recordset.Fields("ContactNumber")
End Sub

Private Sub Grid_Click()
If Me.cmdAdd.Enabled = True Then
Call settext
Else
End If
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Me.cmdAdd.Enabled = True Then
Call settext
Else
End If
End Sub

Private Sub txtSearch_Change()
On Error Resume Next
Me.Adosupplier.Refresh
Me.Adosupplier.Recordset.Find ("SupplierNumber = '" & Me.txtsearch.Text & "'")
Call settext
Call setgrid
End Sub

Private Sub txtSearch_Click(Area As Integer)
On Error Resume Next
Me.Adosupplier.Refresh
Me.Adosupplier.Recordset.Find ("SupplierNumber = '" & Me.txtsearch.Text & "'")
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
