VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log-in"
   ClientHeight    =   2430
   ClientLeft      =   5160
   ClientTop       =   3825
   ClientWidth     =   5415
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
   ScaleHeight     =   2430
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo txtUser 
      Bindings        =   "frmLogin.frx":0000
      Height          =   390
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "UserName"
      Text            =   ""
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
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
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
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
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
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   1455
      Left            =   120
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   0
      Top             =   1800
      Width           =   5415
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
If Me.cmdClose.Caption = "Close" Then
Dim reply
reply = MsgBox("Do you realy want to Quit this program?", vbYesNo + vbQuestion, "Confirm Quit")
    If reply = vbYes Then
    End
    End If
Else
Unload Me
End If
End Sub

Private Sub cmdOk_Click()
Me.AdoUser.Refresh
If Me.cmdClose.Caption = "Close" Then
Me.AdoUser.Recordset.Find ("UserName = '" & txtUser.Text & "'")
    If Me.txtPassword.Text = Me.AdoUser.Recordset.Fields("Password") Then
        If Me.AdoUser.Recordset.Fields("Previledge") = "Manager" Then
        mdiMain.mnuItems.Enabled = True
        mdiMain.mnuOrders.Enabled = True
        mdiMain.stbMain.Panels(1).Text = Me.AdoUser.Recordset.Fields("UserName")
        mdiMain.Show
        End If
        If Me.AdoUser.Recordset.Fields("Previledge") = "Cashier" Then
        mdiMain.mnuItems.Enabled = False
        mdiMain.mnuOrders.Enabled = False
        mdiMain.stbMain.Panels(1).Text = Me.AdoUser.Recordset.Fields("UserName")
        mdiMain.Show
        End If
        Unload Me
    Else
    MsgBox "Invalid Password!!!", vbExclamation, "Invalid Password"
    Me.txtPassword.SetFocus
    SendKeys HiLyt
    End If
    
End If
If Me.cmdClose.Caption = "Cancel" Then
Me.AdoUser.Recordset.Find ("UserName = '" & txtUser.Text & "'")
    If Me.txtPassword.Text = Me.AdoUser.Recordset.Fields("Password") Then
        If Me.AdoUser.Recordset.Fields("Previledge") = "Manager" Then
        mdiMain.mnuItems.Enabled = True
        mdiMain.mnuOrders.Enabled = True
        mdiMain.stbMain.Panels(1).Text = Me.AdoUser.Recordset.Fields("UserName")
        End If
        If Me.AdoUser.Recordset.Fields("Previledge") = "Cashier" Then
        mdiMain.mnuItems.Enabled = False
        mdiMain.mnuOrders.Enabled = False
        mdiMain.stbMain.Panels(1).Text = Me.AdoUser.Recordset.Fields("UserName")
        End If
        Unload Me
    Else
    MsgBox "Invalid Password!!!", vbExclamation, "Invalid Password"
    Me.txtPassword.SetFocus
    SendKeys HiLyt
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SQLDB(AdoUser, "Select * from UserAccount")
AdoUser.Refresh
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk_Click
End If
End Sub

Private Sub txtUser_Click(Area As Integer)
On Error Resume Next
Dim temp
Me.AdoUser.Recordset.Find ("UserName = '" & txtUser.Text & "'")
temp = Me.AdoUser.Recordset.Fields(1)
End Sub
