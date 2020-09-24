VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPayment 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment"
   ClientHeight    =   3660
   ClientLeft      =   7245
   ClientTop       =   4755
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
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
   ScaleHeight     =   3660
   ScaleWidth      =   4965
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtPayment 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Payment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2655
      Left            =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   0
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim reply
reply = MsgBox("Are you sure of this payment?.Payment can be done only once in this transaction.", vbQuestion + vbYesNo, "Confirmation")
If reply = vbYes Then
    If Val(Me.txtPayment) > Val(frmNewSales.txtTotalCost) Then
    Me.txtBalance.Text = "0.00"
    Me.txtChange.Text = Format(Val(Me.txtPayment) - Val(frmNewSales.txtTotalCost), "#,##0.00")
    Me.cmdOK.Enabled = False
    frmNewSales.txtBalance.Text = "0.00"
    frmNewSales.txtPayment.Text = Val(Me.txtPayment)
    frmNewSales.lblChange.Caption = Format(Val(Me.txtPayment) - Val(frmNewSales.txtTotalCost), "#,##0.00")
    frmNewSales.txtAmountPaid.Text = Val(frmNewSales.txtTotalCost)
    frmNewSales.txtAmountPaid.Visible = False
    frmNewSales.cmdReciept.Enabled = True
    frmNewSales.cmdPayment.Enabled = False
    frmNewSales.cmdEditItems.Enabled = False
    frmNewSales.cmdRemove.Enabled = False
    Else
    Me.txtBalance.Text = Format(Val(frmNewSales.txtTotalCost) - Val(Me.txtPayment.Text), "#,##0.00")
    Me.txtChange.Text = "0.00"
    Me.cmdOK.Enabled = False
    frmNewSales.txtBalance.Text = Format(Val(frmNewSales.txtTotalCost) - Val(Me.txtPayment.Text), "#,##0.00")
    frmNewSales.lblChange.Caption = "0.00"
    frmNewSales.txtAmountPaid.Text = Val(Me.txtPayment)
    frmNewSales.txtPayment.Text = Val(Me.txtPayment)
    frmNewSales.cmdReciept.Enabled = True
    frmNewSales.cmdPayment.Enabled = False
    frmNewSales.cmdEditItems.Enabled = False
    frmNewSales.cmdRemove.Enabled = False
    End If
Else
End If
End Sub
