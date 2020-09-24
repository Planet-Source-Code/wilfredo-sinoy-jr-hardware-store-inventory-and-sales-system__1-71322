VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmResearcher 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Researcher"
   ClientHeight    =   5265
   ClientLeft      =   4725
   ClientTop       =   2055
   ClientWidth     =   5655
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
   ScaleHeight     =   5265
   ScaleWidth      =   5655
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   4680
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
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address: JelaineLoredo@HotMail.com.ph"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number: 09206106083"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Address: Guinhulacan, Bingawan, Iloilo"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Jelaine Loredo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address: MaeDePedro@yahoo.com.ph"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number: 09197115232"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Address: Quinagaringan Grande, Passi City, Iloilo"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mae De Pedro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address: JerolLeysa@yahoo.com.ph"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number: 09095272831"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address: Malitbog Ilaya, Bingawan, Iloilo"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Jerol Leysa:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Researcher "
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
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   615
      Left            =   120
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      Top             =   600
      Width           =   5415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      Top             =   3240
      Width           =   5415
   End
End
Attribute VB_Name = "frmResearcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

