VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditItems 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Items"
   ClientHeight    =   4650
   ClientLeft      =   5160
   ClientTop       =   2775
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6150
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3360
      Width           =   4215
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox txtinfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
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
      TabIndex        =   8
      Top             =   4080
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   3960
      Width           =   5895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "UnitPrice:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   3495
      Left            =   120
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmEditItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Val(Me.txtinfo(4).Text) = 0 Then
MsgBox "Please input the Quantity.", vbExclamation, "Hardware System"
Else
    frmNewSales.txtTotalCost.Text = Val(frmNewSales.txtTotalCost.Text) - frmNewSales.AdoSalesInfo.Recordset.Fields("TotalPrice")
    frmNewSales.txtBalance.Text = Val(frmNewSales.txtBalance.Text) - frmNewSales.AdoSalesInfo.Recordset.Fields("TotalPrice")
    frmNewSales.AdoItems.Recordset.Fields("RemainingQty") = frmNewSales.AdoItems.Recordset.Fields("RemainingQty") + frmNewSales.AdoSalesInfo.Recordset.Fields("Qty")
    frmNewSales.AdoItems.Recordset.Update
    frmNewSales.AdoItems.Refresh
        With frmNewSales.AdoSalesInfo
        .Recordset.Fields("SalesNumber") = Me.txtinfo(0).Text
        .Recordset.Fields("Date") = Me.txtinfo(1).Text
        .Recordset.Fields("ItemNumber") = Me.txtinfo(2).Text
        .Recordset.Fields("ItemName") = Me.txtinfo(3).Text
        .Recordset.Fields("Qty") = Me.txtinfo(4).Text
        .Recordset.Fields("UnitPrice") = Me.txtinfo(5).Text
        .Recordset.Fields("TotalPrice") = Me.txtinfo(6).Text
        .Recordset.Update
        .Refresh
        End With
    frmNewSales.AdoItems.Refresh
    frmNewSales.AdoItems.Recordset.Find ("ItemNumber= '" & Me.txtinfo(2).Text & "'")
    frmNewSales.AdoItems.Recordset.Fields("RemainingQty") = Val(frmNewSales.AdoItems.Recordset.Fields("RemainingQty")) - Val(Me.txtinfo(4).Text)
    frmNewSales.AdoItems.Recordset.Update
    frmNewSales.AdoItems.Refresh
    frmNewSales.txtTotalCost.Text = Format(Val(frmNewSales.txtTotalCost.Text) + Val(Me.txtinfo(6).Text), "##0.00")
    frmNewSales.txtBalance.Text = Format(Val(frmNewSales.txtBalance.Text) + Val(Me.txtinfo(6).Text), "##0.00")
    Call SQLDB(frmNewSales.AdoSalesInfo, "Select * from SalesInfo where SalesNumber='" & frmNewSales.txtSalesNumber.Text & "'")
    Call SQLDB3(frmNewSales.AdoItems, "Select * from Items where RemainingQty > 0")
    frmNewSales.AdoItems.Refresh
    frmNewSales.Refresh
    frmNewSales.Grid2.Columns(1).Visible = False
    frmNewSales.Grid2.Columns(2).Visible = False
    frmNewSales.Grid2.Columns(5).NumberFormat = "##0.00"
    frmNewSales.Grid2.Columns(6).NumberFormat = "##0.00"
    frmNewSales.Grid.Columns(0).Visible = False
    frmNewSales.Grid.Columns(2).Visible = False
    frmNewSales.Grid.Columns(5).Visible = False
    Unload Me
End If
End Sub

Private Sub txtinfo_Change(Index As Integer)
If Val(txtinfo(4).Text) > frmNewSales.AdoItems.Recordset("RemainingQty") + frmNewSales.AdoSalesInfo.Recordset.Fields("Qty") Then
MsgBox "Quantity is too high.", vbExclamation, "Hardware System"
Me.txtinfo(4).Text = "0"
Else
Me.txtinfo(6).Text = Val(Me.txtinfo(4).Text) * Val(Me.txtinfo(5).Text)
End If
End Sub
