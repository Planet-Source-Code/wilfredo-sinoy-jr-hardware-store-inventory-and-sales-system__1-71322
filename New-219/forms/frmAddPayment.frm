VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddPayment 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment"
   ClientHeight    =   4590
   ClientLeft      =   3930
   ClientTop       =   2355
   ClientWidth     =   5550
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
   ScaleHeight     =   4590
   ScaleWidth      =   5550
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Enter Payment Here:"
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   5295
      Begin VB.TextBox txtPayment 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Payment Information:"
      ForeColor       =   &H8000000E&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      Begin VB.TextBox txtsalesnumber 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtAmmountPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txttotalCost 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Number:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cost:"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   4080
      Width           =   2415
      _ExtentX        =   4260
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
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "Ok"
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
      Left            =   120
      Top             =   3960
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmAddPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Val(txtPayment.Text) = 0 Then
MsgBox "Zero payment!!!", vbExclamation, "Hardware system"
Me.txtPayment.SetFocus
Else
Dim reply
    reply = MsgBox("Are you sure of this payment?", vbYesNo + vbQuestion, "Confirm Payment")
    If reply = vbYes Then
    With frmUpdateBalance
        .AdoSales.Recordset.Fields("AmountPaid") = Val(.AdoSales.Recordset.Fields("AmountPaid")) + Val(txtPayment.Text)
        .AdoSales.Recordset.Fields("Balance") = Val(.AdoSales.Recordset.Fields("Balance")) - Val(txtPayment.Text)
        .txtAmmountPaid.Text = Format(Val(.txtAmmountPaid.Text) + Val(txtPayment.Text), "##0.00")
        .txtBalance.Text = Format(Val(.txtBalance.Text) - Val(txtPayment.Text), "##0,00")
        .AdoSales.Recordset.Update
        .AdoSales.Refresh
        Unload Me
    End With
    End If
End If
End Sub

Private Sub txtPayment_Change()
If frmAddPayment.Visible = True Then
    If Val(txtPayment.Text) > Val(txtBalance.Text) Then
    MsgBox "The payment you entered is too high that its balanced!!!", vbExclamation, "Hardware System"
    Me.txtPayment.Text = "0.00"
    End If
End If
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk_Click
End If
If KeyAscii = 13 Then
cmdOk_Click
End If
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
