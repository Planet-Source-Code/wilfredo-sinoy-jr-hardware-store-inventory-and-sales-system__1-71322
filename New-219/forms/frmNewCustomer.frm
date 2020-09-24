VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmNewCustomer 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Customer"
   ClientHeight    =   3540
   ClientLeft      =   4185
   ClientTop       =   3750
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6285
   Begin VB.TextBox txtinfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   3
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
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
      TabIndex        =   9
      Top             =   3000
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   2880
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   2295
      Left            =   120
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "frmNewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
Dim reply
frmNewSales.adoCustomer.Refresh
frmNewSales.adoCustomer.Recordset.Find ("CustomerName='" & LCase(Me.txtinfo(1).Text) & "'")
If Me.txtinfo(1).Text = "" Or Me.txtinfo(2).Text = "" Then
MsgBox "Please complete empty box/boxes.", vbExclamation, "Hardware System"
Else
    If Me.txtinfo(1).Text <> frmNewSales.adoCustomer.Recordset.Fields("CustomerName") Then
    With frmNewSales.adoCustomer
    .Recordset.AddNew
    .Recordset.Fields("CustomerNumber") = Me.txtinfo(0).Text
    .Recordset.Fields("CustomerName") = Me.txtinfo(1).Text
    .Recordset.Fields("Address") = Me.txtinfo(2).Text
    .Recordset.Fields("ContactNumber") = Me.txtinfo(3).Text
    .Recordset.Update
    .Refresh
    .Refresh
    frmNewSales.txtCustomerName.Text = Me.txtinfo(1).Text
    frmNewSales.txtCustomerNumber.Text = Me.txtinfo(0).Text
    frmNewSales.txtCustomerNumber.Locked = True
    frmNewSales.txtCustomerName.Locked = True
    frmNewSales.cmdNewCustomer.Enabled = False
    Unload Me
    End With
    Else
    reply = MsgBox("The Customer already exist!!!, Do you want to add it to your transaction?", vbYesNo + vbQuestion, "Error")
        If reply = vbYes Then
        frmNewSales.txtCustomerName.Text = frmNewSales.adoCustomer.Recordset.Fields("CustomerName")
        frmNewSales.txtCustomerNumber.Text = frmNewSales.adoCustomer.Recordset.Fields("CustomerNumber")
        Unload Me
        Else
        End If
    End If
End If

End Sub

