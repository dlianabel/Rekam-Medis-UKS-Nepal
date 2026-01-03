VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form fpemeriksaan 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "fpemeriksaan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton backBTN 
      Caption         =   "MENU"
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
      Left            =   12840
      TabIndex        =   10
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton dtpasienBTN 
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
      Height          =   375
      Left            =   11160
      TabIndex        =   9
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton hpsBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "HAPUS"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton periksaBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "PERIKSA"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox cariTX 
      Height          =   405
      Left            =   9240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00B58A57&
      Caption         =   "CARI"
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "fpemeriksaan.frx":2C6F1
      Height          =   2415
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4260
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "idrekam"
         Caption         =   "ID"
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
         Caption         =   "NK"
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
      BeginProperty Column03 
         DataField       =   "status"
         Caption         =   "STATUS"
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
         Caption         =   "TANGGAL"
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
         Caption         =   "JAM"
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
         Caption         =   "PEMERIKSA"
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
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3360,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1425,26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   14,74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7920
      Top             =   2520
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
      CommandType     =   1
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
      RecordSource    =   "select * from tbpemeriksaan where status=""Pemeriksaan"";"
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
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00422A0D&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   1680
      Top             =   2400
      Width           =   11055
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
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
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
   Begin Project1.PictureG PictureG3 
      Height          =   3345
      Left            =   -360
      Top             =   7320
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "fpemeriksaan.frx":2C706
      Mirror          =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pemeriksaan"
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
      TabIndex        =   5
      Top             =   0
      Width           =   5175
   End
   Begin Project1.PictureG PictureG6 
      Height          =   11475
      Left            =   4200
      Top             =   -3720
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   20241
      GIF             =   "fpemeriksaan.frx":2CDE8
   End
   Begin Project1.PictureG PictureG1 
      Height          =   6090
      Left            =   -1560
      Top             =   480
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   10742
      GIF             =   "fpemeriksaan.frx":FDD22
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pemeriksaan"
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
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
   Begin Project1.PictureG PictureG2 
      Height          =   3345
      Left            =   -480
      Top             =   -2400
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "fpemeriksaan.frx":216D0C
   End
   Begin Project1.PictureG PictureG5 
      Height          =   11670
      Left            =   4320
      Top             =   -3720
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   20585
      GIF             =   "fpemeriksaan.frx":2173EE
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   6735
      Left            =   8520
      Shape           =   2  'Oval
      Top             =   840
      Width           =   7215
   End
   Begin Project1.PictureG PictureG4 
      Height          =   11040
      Left            =   -3960
      Top             =   -1440
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   19473
      GIF             =   "fpemeriksaan.frx":25E140
   End
End
Attribute VB_Name = "fpemeriksaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'==================================KUMPULAN SUB
Sub periksa()
Dim q As New fpemeriksaan_dokter
With Adodc1.Recordset
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Tidak ada data"
Else
If Adodc1.Recordset!pemeriksa = namapemeriksaT.Caption Then

'update 21 oktober 2024
q.Adodc3.Recordset.Filter = "idrekam='" & !idrekam & "'"
'up end

q.idrekamT = .Fields("idrekam").Value
q.nokartuT = .Fields("nokartu").Value
q.pasienT = .Fields("nama").Value
q.statusT = .Fields("Status").Value
q.keluhanTX = .Fields("keluhan").Value
q.tanggalT = !tanggal
q.jamT = !jam
q.namapemeriksaT = !pemeriksa
q.noteleponT = !notelpemeriksa
q.jenkelT = !jenkel
q.umurT = !umur
q.peranT = !peran
'q.idpasienT = !idpasien
q.Show
Unload Me
Else
MsgBox "Pemeriksa Tidak Sesuai!"
End If
End If
End With
End Sub
'==============================KUMPULAN SUB END






'=====================================TOMBOL FUNGSIONAL
Private Sub backBTN_Click()
menu.Show
menu.namapemeriksaT = namapemeriksaT.Caption
menu.noteleponT = noteleponT.Caption
Unload Me
End Sub



Private Sub dtpasienBTN_Click()
f3datapasien.Show
f3datapasien.namapemeriksaT = namapemeriksaT.Caption
f3datapasien.noteleponT = noteleponT.Caption
Unload Me
End Sub

Private Sub hpsBTN_Click()
With Adodc1.Recordset
If .RecordCount = 0 Then
MsgBox "Tidak ada data pemeriksaan!"
Else
respons = MsgBox("Hapus Data Pemeriksaan?" & vbCrLf & "Tindakan ini hanya dilakukan ketika ada kondisi darurat", vbOKCancel)
If respons = vbOK Then
Adodc1.Recordset.Delete
Adodc1.Recordset.Filter = "idrekam <> '" & idrekam & "'"
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
End If
End If
End With
End Sub

Private Sub reBTN_Click()
Adodc1.Recordset.Filter = "idrekam<>"""
End Sub
'=================================TOMBOL FUNGSIONAL END

Private Sub periksaBTN_Click()
Dim q As New fpemeriksaan_dokter
periksa
End Sub


