VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form f2dataobat 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "f2dataobat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox jmldataTX 
      Height          =   405
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   6960
      Width           =   1215
   End
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
      TabIndex        =   21
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cariBTN 
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox cariTX 
      Height          =   405
      Left            =   10200
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton refreshBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "REFRESH"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton lihatBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "LIHAT DETAIL"
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
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton stokobtBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "STOK"
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton hpsBTN 
      BackColor       =   &H00EB9E3F&
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton updBTN 
      BackColor       =   &H00EB9E3F&
      Caption         =   "EDIT"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton tmbhBTN 
      BackColor       =   &H00EB9E3F&
      Caption         =   "TAMBAH"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox stokTX 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.ComboBox satuanCMBBX 
      Height          =   315
      ItemData        =   "f2dataobat.frx":2C6F1
      Left            =   720
      List            =   "f2dataobat.frx":2C707
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox kegunaanTX 
      Height          =   1725
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "f2dataobat.frx":2C73B
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox namaTX 
      Height          =   405
      Left            =   720
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2160
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "f2dataobat.frx":2C741
      Height          =   5055
      Left            =   5280
      TabIndex        =   0
      Top             =   1800
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8916
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
         DataField       =   "idobat"
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
         DataField       =   "nama"
         Caption         =   "NAMA OBAT"
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
         DataField       =   "kegunaan"
         Caption         =   "SATUAN"
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
         DataField       =   "satuan"
         Caption         =   "SATUAN"
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
         DataField       =   "stok"
         Caption         =   "STOK"
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
            ColumnWidth     =   4155,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1184,882
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   1320
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
      BackColor       =   16777215
      ForeColor       =   0
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\dbrekam_medis.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\dbrekam_medis.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbobat"
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stok"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00422A0D&
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   5160
      Top             =   1200
      Width           =   8535
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Obat"
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
      TabIndex        =   4
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Obat"
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
      TabIndex        =   5
      Top             =   360
      Width           =   5175
   End
   Begin Project1.PictureG PictureG5 
      Height          =   11760
      Left            =   -3120
      Top             =   -3600
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   20743
      GIF             =   "f2dataobat.frx":2C756
   End
   Begin Project1.PictureG PictureG2 
      Height          =   3345
      Left            =   -2040
      Top             =   -2040
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "f2dataobat.frx":FD690
   End
   Begin Project1.PictureG PictureG6 
      Height          =   11040
      Left            =   -3000
      Top             =   -1560
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   19473
      GIF             =   "f2dataobat.frx":FDD72
   End
   Begin Project1.PictureG PictureG3 
      Height          =   3345
      Left            =   9000
      Top             =   -3000
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "f2dataobat.frx":103140
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan"
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
      Left            =   720
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kegunaan"
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
      Left            =   720
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label2 
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
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   10695
      Left            =   9000
      Shape           =   2  'Oval
      Top             =   -360
      Width           =   17775
   End
   Begin Project1.PictureG PictureG1 
      Height          =   11760
      Left            =   -3240
      Top             =   -3480
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   20743
      GIF             =   "f2dataobat.frx":103822
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   10695
      Left            =   -14040
      Shape           =   2  'Oval
      Top             =   1080
      Width           =   17775
   End
   Begin Project1.PictureG PictureG4 
      Height          =   3345
      Left            =   2400
      Top             =   7320
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "f2dataobat.frx":14A574
      Mirror          =   1
   End
End
Attribute VB_Name = "f2dataobat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cariBTN_Click()
'With Adodc1.Recordset
'.Filter = "idobat = '" & cariTX.Text & "' OR nama = '" & cariTX.Text & "' OR satuan = '" & cariTX.Text & "' OR kegunaan = '" & cariTX.Text & "'"
'End With
If cariTX.Text = "" Then
    Adodc1.Refresh
Else
    Adodc1.Recordset.Filter = "idobat = '" & cariTX.Text & "' OR nama = '" & cariTX.Text & "' OR satuan = '" & cariTX.Text & "' OR kegunaan = '" & cariTX.Text & "'"
End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
Private Sub cariTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cariBTN_Click
End If
End Sub


Private Sub Form_Load()
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub

'==================================KUMPULAN SUB
Sub kosong()
namaTX.Text = ""
kegunaanTX.Text = ""
'satuanCMBBX.Text = ""
satuanCMBBX.ListIndex = -1
stokTX.Text = ""
cariTX.Text = ""
End Sub
Sub tampil()
With Adodc1.Recordset
namaTX.Text = !nama
kegunaanTX.Text = !kegunaan
satuanCMBBX.Text = !satuan
stokTX.Text = !stok
End With
End Sub
Sub buat()
If namaTX.Text <> "" And kegunaanTX.Text <> "" And satuanCMBBX.Text <> "" Then
With Adodc1.Recordset
maksimalrek = 100000
If .RecordCount >= maksimalrek Then
MsgBox "Jumlah data sudah mencapai batas maksimal"
Exit Sub
End If
If .EOF Then
urutan = "OBT" + "00001"
Else
.MoveLast
nomorurut = CInt(Mid(!idobat, 4)) 'OBT12345
nomorurut = nomorurut + 1
urutan = "OBT" + Right("00000" & nomorurut, 5)
End If

.AddNew
!idobat = urutan
!nama = namaTX.Text
!kegunaan = kegunaanTX.Text
!satuan = satuanCMBBX.Text
!stok = 0
.Update
kosong
'Exit Sub
End With
Else
MsgBox "Masih ada yang kosong!"
End If
End Sub
Sub edit()
With Adodc1.Recordset
respons = MsgBox("Perbarui Data?", vbOKCancel)
If respons = vbOK Then
If namaTX.Text <> "" And kegunaanTX.Text <> "" And satuanCMBBX.Text <> "" Then
!nama = namaTX.Text
!kegunaan = kegunaanTX.Text
!satuan = satuanCMBBX.Text
.Update
Else
MsgBox "Masih ada yang kosong!"
End If
End If
End With
End Sub
'==============================KUMPULAN SUB END






'=====================================TOMBOL FUNGSIONAL
'TOMBOL TUTUP
Private Sub backBTN_Click()
Unload Me
End Sub






Private Sub lihatBTN_Click()
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Masih belum ada Data"
Else
If lihatBTN.Caption = "LIHAT DETAIL" Then
lihatBTN.Caption = "TUTUP DETAIL"
tmbhBTN.Enabled = False
updBTN.Enabled = False
hpsBTN.Enabled = False
Adodc1.Enabled = False
namaTX.Enabled = False
kegunaanTX.Enabled = False
satuanCMBBX.Enabled = False
stokobtBTN.Enabled = False
stokTX.Enabled = False
cariTX.Enabled = False
cariBTN.Enabled = False
refreshBTN.Enabled = False
tampil
ElseIf lihatBTN.Caption = "TUTUP DETAIL" Then
kosong
lihatBTN.Caption = "LIHAT DETAIL"
tmbhBTN.Enabled = True
updBTN.Enabled = True
hpsBTN.Enabled = True
Adodc1.Enabled = True
namaTX.Enabled = True
kegunaanTX.Enabled = True
satuanCMBBX.Enabled = True
stokobtBTN.Enabled = True
stokTX.Enabled = True
cariTX.Enabled = True
cariBTN.Enabled = True
refreshBTN.Enabled = True
kosong
End If
End If
End Sub


Private Sub refreshBTN_Click()
Adodc1.Refresh
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub

Private Sub stokobtBTN_Click()
Dim q As New f2dataobatstok
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Masih belum ada Data"
Exit Sub
Else
'f2dataobatstok.namapemeriksaT = namapemeriksaT.Caption
'f2dataobatstok.noteleponT = noteleponT.Caption
'q.summonobat
q.Adodc2.Recordset.MoveFirst
q.Adodc2.Recordset.Find "nama = '" & Adodc1.Recordset.Fields("nama").Value & "'"
q.namaTX.Text = Adodc1.Recordset.Fields("nama").Value
q.satuanTX.Text = Adodc1.Recordset.Fields("satuan").Value
q.stokawalTX.Text = Adodc1.Recordset.Fields("stok").Value
q.totalstokTX.Text = Adodc1.Recordset.Fields("stok").Value
q.Show
q.namapemeriksaT = namapemeriksaT.Caption
q.noteleponT = noteleponT.Caption
Unload Me
End If
End Sub



'TOMBOL CREATE
Private Sub tmbhBTN_Click()
buat
Adodc1.Recordset.Filter = "idobat <>"""
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
'namaTX.SetFocus
End Sub
'TOMBOL UPDATE
Private Sub updBTN_Click()
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Tidak ada data! Buat data terlebih dahulu!"
Else
If updBTN.Caption = "EDIT" Then
updBTN.Caption = "PERBARUI"
tmbhBTN.Enabled = False
hpsBTN.Enabled = False
Adodc1.Enabled = False
stokobtBTN.Enabled = False
lihatBTN.Enabled = False
cariTX.Enabled = False
cariBTN.Enabled = False
refreshBTN.Enabled = False
tampil
ElseIf updBTN.Caption = "PERBARUI" Then
edit
updBTN.Caption = "EDIT"
tmbhBTN.Enabled = True
hpsBTN.Enabled = True
Adodc1.Enabled = True
stokobtBTN.Enabled = True
lihatBTN.Enabled = True
cariTX.Enabled = True
cariBTN.Enabled = True
refreshBTN.Enabled = True
kosong
End If
End If
End Sub
'TOMBOL DELETE
Private Sub hpsBTN_Click()
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Tidak ada data! Buat data terlebih dahulu!"
Else
respons = MsgBox("Hapus Data?", vbOKCancel)
If respons = vbOK Then
Adodc1.Recordset.Delete
Adodc1.Recordset.Filter = "idobat <> '" & idobat & "'"
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
End If
End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
'=================================TOMBOL FUNGSIONAL END


'=====================================KEYPRESS
Private Sub namaTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
kegunaanTX.SetFocus
End If
End Sub
Private Sub kegunaanTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
satuanCMBBX.SetFocus
End If
End Sub
Private Sub satuanCMBBX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tmbhBTN.SetFocus
End If
End Sub
'=================================KEYPRESS END
