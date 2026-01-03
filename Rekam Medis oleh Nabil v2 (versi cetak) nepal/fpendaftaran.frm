VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form fpendaftaran 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   ControlBox      =   0   'False
   Icon            =   "fpendaftaran.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton backBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "BATAL"
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
      TabIndex        =   16
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton daftarBTN 
      BackColor       =   &H00EB9E3F&
      Caption         =   "DAFTAR"
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
      TabIndex        =   15
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox keluhanTX 
      Height          =   1485
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "fpendaftaran.frx":2C6F1
      Top             =   3480
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "fpendaftaran.frx":2C6F7
      Height          =   1695
      Left            =   1320
      TabIndex        =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2990
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "idrekam"
         Caption         =   "idrekam"
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
         DataField       =   "nokartu"
         Caption         =   "nokartu"
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
         DataField       =   "nama"
         Caption         =   "nama"
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
         DataField       =   "status"
         Caption         =   "status"
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
         DataField       =   "keluhan"
         Caption         =   "keluhan"
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
      BeginProperty Column05 
         DataField       =   "tanggal"
         Caption         =   "tanggal"
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
      BeginProperty Column06 
         DataField       =   "jam"
         Caption         =   "jam"
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
      BeginProperty Column07 
         DataField       =   "pemeriksa"
         Caption         =   "pemeriksa"
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
      BeginProperty Column08 
         DataField       =   "notelpemeriksa"
         Caption         =   "notelpemeriksa"
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
      BeginProperty Column09 
         DataField       =   "jenkel"
         Caption         =   "jenkel"
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
      BeginProperty Column10 
         DataField       =   "umur"
         Caption         =   "umur"
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
      BeginProperty Column11 
         DataField       =   "peran"
         Caption         =   "peran"
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
      BeginProperty Column12 
         DataField       =   "idpasien"
         Caption         =   "idpasien"
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
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3614,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   329,953
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   390,047
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   480,189
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   195,024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
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
      RecordSource    =   "tbpemeriksaan"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran"
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
      TabIndex        =   10
      Top             =   120
      Width           =   5895
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
      Left            =   1560
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Keluhan"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pasien"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pemeriksa"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label pasienT 
      BackStyle       =   0  'Transparent
      Caption         =   "nama pasien"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label idpasienT 
      BackStyle       =   0  'Transparent
      Caption         =   "id pasien"
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
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label nokartuT 
      BackStyle       =   0  'Transparent
      Caption         =   "nokartu"
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
      Left            =   240
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label jenkelT 
      BackStyle       =   0  'Transparent
      Caption         =   "jenkel"
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
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label umurT 
      BackStyle       =   0  'Transparent
      Caption         =   "umur"
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
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label peranT 
      BackStyle       =   0  'Transparent
      Caption         =   "peran"
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
      Left            =   240
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   3615
   End
   Begin Project1.PictureG PictureG1 
      Height          =   9195
      Left            =   240
      Top             =   -120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   16219
      GIF             =   "fpendaftaran.frx":2C70C
      Stretch         =   2
      Mirror          =   1
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   10695
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   -7200
      Width           =   17775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   10695
      Left            =   -13680
      Shape           =   2  'Oval
      Top             =   4800
      Width           =   17775
   End
End
Attribute VB_Name = "fpendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
kosong
End Sub


'==================================KUMPULAN SUB
Sub kosong()
keluhanTX = ""
End Sub
'==============================KUMPULAN SUB END


'===============================KUMPULAN TOMBOL
Private Sub daftarBTN_Click()
If keluhanTX.Text = "" Then
MsgBox "Keluhan belum diisi!"
keluhanTX.SetFocus
Else
Dim q As New fpemeriksaan
wjam = Format(Now, "HH:mm:ss")
With Adodc2.Recordset

maksimalrek = 10000000
If .RecordCount >= maksimalrek Then
MsgBox "Jumlah data sudah mencapai batas maksimal"
Exit Sub
End If
If .EOF Then
urutan = "RM" + "0000001"
Else
.MoveLast
nomorurut = CInt(Mid(!idrekam, 3)) 'contoh : RM1234567
nomorurut = nomorurut + 1
urutan = "RM" + Right("0000000" & nomorurut, 7)
End If

.AddNew
!idrekam = urutan
!nokartu = nokartuT
!nama = pasienT
!Status = "Pemeriksaan"
!keluhan = keluhanTX.Text
!tanggal = Format(Now, "dd/mm/yyyy")
!jam = wjam
!pemeriksa = namapemeriksaT
!jenkel = jenkelT
!umur = umurT
!peran = peranT
'!idpasien = idpasienT
!notelpemeriksa = noteleponT
.Update
Adodc2.Refresh
End With
q.namapemeriksaT = namapemeriksaT.Caption
q.noteleponT = noteleponT.Caption
q.Show
Unload Me
Unload f3datapasien
End If
End Sub
Private Sub backBTN_Click()
Unload Me
End Sub
'===========================KUMPULAN TOMBOL END

