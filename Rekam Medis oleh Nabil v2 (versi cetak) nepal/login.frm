VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form login 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton tombol 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGIN"
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox passTX 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox userTX 
      Height          =   405
      Left            =   5640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2640
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   0
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "login.frx":2C6F1
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2355
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   6600
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin Project1.PictureG PictureG5 
      Height          =   11880
      Left            =   -3120
      Top             =   960
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   20955
      GIF             =   "login.frx":2C706
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG6 
      Height          =   11040
      Left            =   7080
      Top             =   -4440
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   19473
      GIF             =   "login.frx":2FF78
      Mirror          =   3
   End
   Begin Project1.PictureG PictureG2 
      Height          =   10680
      Left            =   -2760
      Top             =   1560
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   18838
      GIF             =   "login.frx":39C92
      Mirror          =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"login.frx":3D050
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   14175
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   8055
      Left            =   -2760
      Shape           =   2  'Oval
      Top             =   7200
      Width           =   8895
   End
   Begin Project1.PictureG PictureG3 
      Height          =   28800
      Left            =   6960
      Top             =   -5640
      Width           =   28800
      _ExtentX        =   50800
      _ExtentY        =   50800
      GIF             =   "login.frx":3D0A2
      Mirror          =   3
   End
   Begin Project1.PictureG PictureG7 
      Height          =   11880
      Left            =   6480
      Top             =   -5040
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   20955
      GIF             =   "login.frx":40914
      Mirror          =   3
   End
   Begin Project1.PictureG PictureG4 
      Height          =   6060
      Left            =   -840
      Top             =   1080
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   10689
      GIF             =   "login.frx":44186
      Stretch         =   2
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG1 
      Height          =   4740
      Left            =   1920
      Top             =   -3360
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   8361
      GIF             =   "login.frx":44868
      Stretch         =   2
      Mirror          =   1
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
kosong
End Sub

Private Sub form_activate()
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Jika belum memiliki akun langsung klik login"
End If
End Sub

Sub kosong()
userTX.Text = ""
passTX.Text = ""
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub passTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tombol_Click
End If
End Sub

Private Sub tombol_Click()
If Adodc1.Recordset.RecordCount = 0 Then
firsttimeuser.Show
Else
Adodc1.Refresh
With Adodc1.Recordset
n = .RecordCount
If (n > 0) Then

If (userTX.Text = "") Or (passTX.Text = "") Then
MsgBox "Masih ada yang kosong"
userTX.SetFocus
Exit Sub
End If

.MoveFirst
.Find "username = '" & userTX.Text & "'"

If .BOF Or .EOF Then
MsgBox "Username tidak ditemukan"
userTX.SetFocus
Exit Sub
End If

'mencari user dan password / apakah sama? begitu
If (!UserName = userTX.Text) And (!Password = passTX.Text) Then
MsgBox "Login Berhasil!"
menu.namapemeriksaT = !nama
menu.noteleponT = !notelp
menu.Show
Unload Me
Else
MsgBox "Username atau Password salah"
userTX.SetFocus
End If

End If
End With
End If
End Sub

Private Sub userTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
passTX.SetFocus
End If
End Sub

