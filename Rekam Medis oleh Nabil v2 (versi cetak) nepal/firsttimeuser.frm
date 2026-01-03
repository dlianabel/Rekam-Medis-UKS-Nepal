VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form firsttimeuser 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   ControlBox      =   0   'False
   Icon            =   "firsttimeuser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton backBTN 
      Caption         =   "KEMBALI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton tmbhBTN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MASUK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   2175
   End
   Begin VB.ComboBox jenkelCMBBX 
      Height          =   315
      ItemData        =   "firsttimeuser.frx":2C6F1
      Left            =   1800
      List            =   "firsttimeuser.frx":2C6FB
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox nomTX 
      Height          =   405
      Left            =   1800
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox namaTX 
      Height          =   405
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox passTX 
      Height          =   405
      Left            =   1800
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox userTX 
      Height          =   405
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1560
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "firsttimeuser.frx":2C715
      Height          =   855
      Left            =   5280
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1508
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "nama"
         Caption         =   "NAMA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "jenkel"
         Caption         =   "JENKEL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "username"
         Caption         =   "USERNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "password"
         Caption         =   "password"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "notelp"
         Caption         =   "TELEPON"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2940,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1620,284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\dbrekam_medis.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\dbrekam_medis.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbpetugas"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Telepon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buat Akun"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
   Begin Project1.PictureG PictureG1 
      Height          =   4740
      Left            =   -2640
      Top             =   7080
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   8361
      GIF             =   "firsttimeuser.frx":2C72A
      Stretch         =   2
      Mirror          =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buat Akun"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5895
   End
   Begin Project1.PictureG PictureG2 
      Height          =   4740
      Left            =   -1440
      Top             =   -2880
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   8361
      GIF             =   "firsttimeuser.frx":2CE0C
      Stretch         =   2
      Mirror          =   1
   End
End
Attribute VB_Name = "firsttimeuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
kosong
End Sub



'==================================KUMPULAN SUB
Sub kosong()
userTX.Text = ""
passTX.Text = ""
namaTX.Text = ""
nomTX.Text = ""
'jenkelCMBBX.Text = ""
jenkelCMBBX.ListIndex = -1
End Sub
Sub tampil()
With Adodc1.Recordset
userTX.Text = !UserName
passTX.Text = !Password
namaTX.Text = !nama
nomTX.Text = !notelp
jenkelCMBBX.Text = !jenkel
cariTX.Text = ""
End With
End Sub
Sub buat()
If userTX.Text <> "" And passTX.Text <> "" And namaTX.Text <> "" And jenkelCMBBX.Text <> "" And nomTX.Text <> "" Then
With Adodc1.Recordset
.AddNew
!UserName = userTX.Text
!Password = passTX.Text
!nama = namaTX.Text
!notelp = nomTX.Text
!jenkel = jenkelCMBBX.Text
.Update
kosong
MsgBox "Akun berhasil dibuat!", , "Sukses"
menu.namapemeriksaT = !nama
menu.noteleponT = !notelp
menu.Show
Unload login
Unload Me
End With
Else
MsgBox "Masih ada yang kosong!"
End If
End Sub
'==============================KUMPULAN SUB END








'=====================================TOMBOL FUNGSIONAL
'TOMBOL CREATE
Private Sub tmbhBTN_Click()
buat
End Sub
'TOMBOL KEMBALI
Private Sub backBTN_Click()
Unload Me
End Sub

'=================================TOMBOL FUNGSIONAL END

'KEYPRESS

Private Sub userTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
passTX.SetFocus
End If
End Sub

Private Sub passTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
namaTX.SetFocus
End If
End Sub

Private Sub namaTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
nomTX.SetFocus
End If
End Sub

Private Sub nomTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
jenkelCMBBX.SetFocus
End If
End Sub

Private Sub jenkelCMBBX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tmbhBTN.SetFocus
End If
End Sub
'KEYPRESS END

