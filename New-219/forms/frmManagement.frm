VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmManagement 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Management"
   ClientHeight    =   4920
   ClientLeft      =   3975
   ClientTop       =   1860
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6750
   Begin RichTextLib.RichTextBox rtf 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5741
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmManagement.frx":0000
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   4320
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Management"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   4200
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "frmManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.rtf.FileName = App.Path & "\Others\management.rtf"
End Sub
