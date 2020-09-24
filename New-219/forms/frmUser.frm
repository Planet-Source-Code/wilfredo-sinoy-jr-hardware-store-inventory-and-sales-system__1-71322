VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUser 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UserAccount"
   ClientHeight    =   2955
   ClientLeft      =   4575
   ClientTop       =   3825
   ClientWidth     =   6045
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
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
   ScaleHeight     =   2955
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtVer 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1680
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtUser 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1680
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "OK"
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
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSAdodcLib.Adodc AdoUser 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
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
   Begin lvButton.lvButtons_H cmdChange 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Change"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   735
      Left            =   0
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.AdoUser.Recordset.Find ("UserName = '" & mdiMain.stbMain.Panels(1).Text & "'")
Me.txtUser.Text = Me.AdoUser.Recordset.Fields("UserName")
Me.txtPassword.Text = Me.AdoUser.Recordset.Fields("Password")
Me.txtVer.Text = Me.AdoUser.Recordset.Fields("Password")
Me.txtPassword.Locked = True
Me.txtUser.Locked = True
Me.txtVer.Locked = True
Me.cmdCancel.Enabled = False
Me.cmdChange.Enabled = True
Me.cmdOk.Enabled = False
End Sub

Private Sub cmdChange_Click()
Me.txtPassword.Locked = False
Me.txtUser.Locked = False
Me.txtVer.Locked = False
Me.cmdCancel.Enabled = True
Me.cmdChange.Enabled = False
Me.cmdOk.Enabled = True
End Sub

Private Sub cmdClose_Click()
If cmdChange.Enabled = False Then
MsgBox "Please finish your transaction before you close.", vbInformation, "Information"
Else
Unload Me
End If
End Sub

Private Sub cmdOk_Click()
If Me.txtPassword.Text = Me.txtVer.Text Then
AdoUser.Refresh
Me.AdoUser.Recordset.Find ("UserName = '" & mdiMain.stbMain.Panels(1).Text & "'")
Me.AdoUser.Recordset.Fields("UserName") = Me.txtUser.Text
Me.AdoUser.Recordset.Fields("Password") = Me.txtPassword.Text
Me.AdoUser.Recordset.Update
Me.txtPassword.Locked = True
Me.txtUser.Locked = True
Me.txtVer.Locked = True
Me.cmdCancel.Enabled = False
Me.cmdChange.Enabled = True
Me.cmdOk.Enabled = False
mdiMain.stbMain.Panels(1).Text = Me.AdoUser.Recordset.Fields("UserName")
Me.Refresh
Else
MsgBox "Verification not match in Password?", vbExclamation, "Verification Error"
SendKeys HiLyt
End If
End Sub

Private Sub Form_Load()
Call SQLDB(AdoUser, "Select * from UserAccount")
AdoUser.Refresh
AdoUser.Refresh
Me.AdoUser.Recordset.Find ("UserName = '" & mdiMain.stbMain.Panels(1).Text & "'")
Me.txtUser.Text = Me.AdoUser.Recordset.Fields("UserName")
Me.txtPassword.Text = Me.AdoUser.Recordset.Fields("Password")
Me.txtVer.Text = Me.AdoUser.Recordset.Fields("Password")
End Sub
