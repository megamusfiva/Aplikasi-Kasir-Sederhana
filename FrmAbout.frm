VERSION 5.00
Begin VB.Form FrmAbout 
   Caption         =   "Tentang"
   ClientHeight    =   6825
   ClientLeft      =   1995
   ClientTop       =   1995
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   Picture         =   "FrmAbout.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   10785
   Begin VB.CommandButton cmdKembali 
      Caption         =   "&Kembali"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   4
      Top             =   6240
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   2880
      Picture         =   "FrmAbout.frx":44A8
      ScaleHeight     =   3465
      ScaleWidth      =   4785
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon: 0821-4251-2288"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat : Jl. Raya Modongan, RT.7/RW.2, Modongan, Kec. Sooko, Mojokerto, Jawa Timur 61361"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   5160
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SISTEM INFORMASI PENYEDIAAN TOKO GROSIR RAFACELL"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKembali_Click()
Unload Me
FrmMain.Show
End Sub

