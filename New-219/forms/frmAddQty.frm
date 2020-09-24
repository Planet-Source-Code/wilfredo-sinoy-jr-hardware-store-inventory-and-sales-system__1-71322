VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddQty 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Confirmation to Add Quantity"
   ClientHeight    =   3060
   ClientLeft      =   8115
   ClientTop       =   2655
   ClientWidth     =   3840
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
   ScaleHeight     =   3060
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Enter the Quantity:"
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3615
      Begin VB.TextBox txtQuantity 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.TextBox txtinfo 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
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
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Number"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   0
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2055
      Left            =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmAddQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim reply
reply = MsgBox("Are you sure?", vbYesNo + vbQuestion, "Confirm")
If reply = vbYes Then
With frmFItems
    .AdoItems.Refresh
    .AdoItems.Recordset.Find ("ItemNumber ='" & Me.txtinfo(0).Text & "'")
    .AdoItems.Recordset.Fields("TotalQty") = .AdoItems.Recordset.Fields("TotalQty") + Val(Me.txtQuantity.Text)
    .AdoItems.Recordset.Fields("RemainingQty") = .AdoItems.Recordset.Fields("RemainingQty") + Val(Me.txtQuantity.Text)
    .AdoItems.Recordset.Update
    .AdoItems.Refresh
    .Refresh
    .AdoItems.Refresh
    .txtinfo(0).Text = .AdoItems.Recordset.Fields("ItemNumber")
    .txtinfo(1).Text = .AdoItems.Recordset.Fields("ItemName")
    .txtUnit.Text = .AdoItems.Recordset.Fields("Unit")
    .txtinfo(3).Text = .AdoItems.Recordset.Fields("Description")
    .txtinfo(4).Text = Format(.AdoItems.Recordset.Fields("UnitPrice"), "##0.00")
    .txtinfo(5).Text = .AdoItems.Recordset.Fields("TotalQty")
    .txtinfo(6).Text = .AdoItems.Recordset.Fields("RemainingQty")
    Unload Me
End With
End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
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
