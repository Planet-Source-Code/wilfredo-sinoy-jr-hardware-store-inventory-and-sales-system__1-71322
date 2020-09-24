VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateDelivery 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Delivery"
   ClientHeight    =   3045
   ClientLeft      =   4830
   ClientTop       =   3555
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5550
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Payment Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      Begin VB.ComboBox txtDeliverySatus 
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
         ItemData        =   "frmUpdateDelivery.frx":0000
         Left            =   1680
         List            =   "frmUpdateDelivery.frx":000A
         TabIndex        =   5
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtsalesnumber 
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
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTDelivery 
         Height          =   345
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   3285
         _ExtentX        =   5794
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
         Format          =   43384833
         CurrentDate     =   38065
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Status:"
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
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date:"
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
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Number:"
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
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
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
      TabIndex        =   4
      Top             =   2520
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   2400
      Width           =   5295
   End
End
Attribute VB_Name = "frmUpdateDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim reply
    reply = MsgBox("Are you sure of this Delivery Information?", vbYesNo + vbQuestion, "Confirm Delivery Change")
    If reply = vbYes Then
    With frmUpdateBalance
        .AdoSales.Recordset.Fields("DeliveryDate") = Me.DTDelivery.Value
        .AdoSales.Recordset.Fields("DeliveryStatus") = Me.txtDeliverySatus.Text
        .txtDeliveryStatus.Text = Me.txtDeliverySatus.Text
        .txtDeliveryDate.Text = Me.DTDelivery.Value
        .AdoSales.Recordset.Update
        .AdoSales.Refresh
        Unload Me
    End With
    End If
End Sub

Private Sub txtDeliverySatus_KeyPress(KeyAscii As Integer)
KeyAscii = False
End Sub
