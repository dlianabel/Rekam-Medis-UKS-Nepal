VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form menu 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleMode       =   0  'User
   ScaleWidth      =   157.667
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "GANTI AKUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BACKUP DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "RIWAYAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PEMERIKSAAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DATA PASIEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DATA OBAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CommandButton tombol 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DATA PETUGAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin Project1.PictureG PictureG8 
      Height          =   9195
      Left            =   9960
      Top             =   -5520
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   16219
      GIF             =   "menu.frx":2C6F1
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG7 
      Height          =   9015
      Left            =   9840
      Top             =   -5280
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   15901
      GIF             =   "menu.frx":73443
      Mirror          =   1
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2024 PPLG"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   7920
      Width           =   3615
   End
   Begin Project1.PictureG PictureG3 
      Height          =   4065
      Left            =   120
      Top             =   7320
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   7170
      GIF             =   "menu.frx":9ACC1
   End
   Begin Project1.PictureG PictureG5 
      Height          =   9990
      Left            =   -3120
      Top             =   2520
      Width           =   19200
      _ExtentX        =   33867
      _ExtentY        =   17621
      GIF             =   "menu.frx":9B3A3
   End
   Begin VB.Label namapemeriksaT 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama pemeriksa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label noteleponT 
      BackStyle       =   0  'Transparent
      Caption         =   "080000000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UKS SMK N 1 PEMALANG"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   9735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "UKS SMK N 1 PEMALANG"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC7F21&
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   9735
   End
   Begin Project1.PictureG PictureG2 
      Height          =   4065
      Left            =   -240
      Top             =   -3120
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   7170
      GIF             =   "menu.frx":D0CB1
   End
   Begin Project1.PictureG PictureG1 
      Height          =   9480
      Left            =   8040
      Top             =   -3720
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   16722
      GIF             =   "menu.frx":D1393
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00B58A57&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   8295
      Left            =   -5880
      Shape           =   2  'Oval
      Top             =   3360
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   10695
      Left            =   7440
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   10575
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
f2dataobat.namapemeriksaT = namapemeriksaT.Caption
f2dataobat.noteleponT = noteleponT.Caption
f2dataobat.Show
End Sub

Private Sub Command2_Click()
f3datapasien.namapemeriksaT = namapemeriksaT.Caption
f3datapasien.noteleponT = noteleponT.Caption
f3datapasien.Show
Unload Me
End Sub

Private Sub Command3_Click()
fpemeriksaan.namapemeriksaT = namapemeriksaT.Caption
fpemeriksaan.noteleponT = noteleponT.Caption
fpemeriksaan.Show
Unload Me
End Sub

Private Sub Command4_Click()
friwayat.namapemeriksaT = namapemeriksaT.Caption
friwayat.noteleponT = noteleponT.Caption
friwayat.Show
End Sub

Private Sub Command5_Click()
login.Show
Unload Me
End Sub

Private Sub Command6_Click()
 Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm
End Sub

Private Sub Command7_Click()
fbackupdata.Show
End Sub

Private Sub tombol_Click()
f1datapetugas.namapemeriksaT = namapemeriksaT.Caption
f1datapetugas.noteleponT = noteleponT.Caption
f1datapetugas.Show
End Sub

