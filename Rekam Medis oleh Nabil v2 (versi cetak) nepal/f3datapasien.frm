VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form f3datapasien 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   2685
   ClientTop       =   1320
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "f3datapasien.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox jmldataTX 
      Height          =   405
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6480
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
      TabIndex        =   24
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton pemeriksaanBTN 
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
      Height          =   375
      Left            =   11040
      TabIndex        =   23
      Top             =   7680
      Width           =   1575
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox cariTX 
      Height          =   405
      Left            =   10560
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1560
      Width           =   2055
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton daftarBTN 
      BackColor       =   &H00B58A57&
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1560
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox jenkelCMBBX 
      Height          =   315
      ItemData        =   "f3datapasien.frx":2C6F1
      Left            =   360
      List            =   "f3datapasien.frx":2C6FB
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5400
      Width           =   2175
   End
   Begin VB.ComboBox peranCMBBX 
      Height          =   315
      ItemData        =   "f3datapasien.frx":2C715
      Left            =   360
      List            =   "f3datapasien.frx":2C725
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox umurTX 
      Height          =   405
      Left            =   360
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox namaTX 
      Height          =   405
      Left            =   360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox nokartuTX 
      Height          =   405
      Left            =   360
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "tes"
      Height          =   375
      Left            =   12840
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "f3datapasien.frx":2C751
      Height          =   4335
      Left            =   3240
      TabIndex        =   1
      Top             =   2040
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7646
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "umur"
         Caption         =   "UMUR"
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
         DataField       =   "peran"
         Caption         =   "PERAN"
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
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5144,882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1890,142
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3240
      Top             =   1560
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
      RecordSource    =   "tbpasien"
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
      Height          =   5535
      Left            =   3120
      Top             =   1440
      Width           =   10815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pasien"
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
      Left            =   8520
      TabIndex        =   9
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Kartu"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Umur"
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
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Peran"
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
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label4 
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
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Width           =   2535
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin Project1.PictureG PictureG5 
      Height          =   9225
      Left            =   -1800
      Top             =   -1200
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   16272
      GIF             =   "f3datapasien.frx":2C766
   End
   Begin Project1.PictureG PictureG3 
      Height          =   8985
      Left            =   -1800
      Top             =   -1080
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   15849
      GIF             =   "f3datapasien.frx":67340
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pasien"
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
      Left            =   8520
      TabIndex        =   10
      Top             =   0
      Width           =   5655
   End
   Begin Project1.PictureG PictureG4 
      Height          =   3585
      Left            =   -120
      Top             =   -2640
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   6324
      GIF             =   "f3datapasien.frx":71752
      Stretch         =   2
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG1 
      Height          =   3585
      Left            =   1800
      Top             =   7200
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   6324
      GIF             =   "f3datapasien.frx":71E34
      Stretch         =   2
   End
   Begin Project1.PictureG PictureG2 
      Height          =   11745
      Left            =   0
      Top             =   -2520
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   20717
      GIF             =   "f3datapasien.frx":72516
      Stretch         =   2
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   6735
      Left            =   -3720
      Shape           =   2  'Oval
      Top             =   4560
      Width           =   7215
   End
End
Attribute VB_Name = "f3datapasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cariBTN_Click()
If cariTX.Text = "" Then
    Adodc1.Recordset.Filter = "" ' Menampilkan semua data
    Adodc1.Refresh
Else
    Adodc1.Recordset.Filter = "nokartu = '" & cariTX.Text & "' OR nama = '" & cariTX.Text & "' OR jenkel = '" & cariTX.Text & "' OR umur = '" & cariTX.Text & "' OR peran = '" & cariTX.Text & "'"
End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
Private Sub cariTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cariBTN_Click
End If
End Sub


Private Sub Command1_Click()
DataReport1.Show
End Sub

Private Sub Form_Load()
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub





'==================================KUMPULAN SUB
Sub kosong()
nokartuTX.Text = ""
namaTX.Text = ""
'jenkelCMBBX.Text = ""
jenkelCMBBX.ListIndex = -1
umurTX.Text = ""
'peranCMBBX.Text = ""
peranCMBBX.ListIndex = -1
cariTX.Text = ""
End Sub
Sub tampil()
With Adodc1.Recordset
nokartuTX.Text = !nokartu
namaTX.Text = !nama
jenkelCMBBX.Text = !jenkel
umurTX.Text = !umur
peranCMBBX.Text = !peran
End With
End Sub

Sub adnewsiswa()
With Adodc1.Recordset
.AddNew
!nokartu = nokartuTX.Text
!nama = namaTX.Text
!jenkel = jenkelCMBBX.Text
!umur = umurTX.Text
!peran = peranCMBBX.Text
.Update
End With
End Sub

Sub buat()
    If nokartuTX.Text <> "" And namaTX.Text <> "" And jenkelCMBBX.Text <> "" And umurTX.Text <> "" And peranCMBBX.Text <> "" Then
        If Val(umurTX.Text) <= 0 Then
                MsgBox "Umur tidak valid!"
                Exit Sub
            End If
                    
        With Adodc1.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "nokartu = '" & nokartuTX.Text & "'"
            
            If Not .EOF Then
                Dim newNokartu As String
                Dim baseNokartu As String
                baseNokartu = nokartuTX.Text
                
                ' Mengambil nomor yang sudah ada dengan akhiran "-X"
                Dim lastNokartu As String
                lastNokartu = nokartuTX.Text
                Dim suffix As Integer
                
                ' Cek jika ada akhiran "-X"
                Do While Not .EOF
                    If InStr(lastNokartu, "-") > 0 Then
                        Dim suffixString As String
                        suffixString = Mid(lastNokartu, InStrRev(lastNokartu, "-") + 1)
                        If IsNumeric(suffixString) Then
                            suffix = CInt(suffixString)
                            lastNokartu = Left(lastNokartu, InStrRev(lastNokartu, "-") - 1)
                        End If
                    End If
                    
                    ' Mengambil nomor kartu baru
                    lastNokartu = baseNokartu & "-" & (suffix + 1)
                    .Find "nokartu = '" & lastNokartu & "'"
                Loop
                
                respons = MsgBox("Nomor Kartu sudah ada. Apakah ingin memberinya nama '" & lastNokartu & "'?", vbOKCancel)
                
                If respons = vbOK Then
                    .AddNew
                    !nokartu = lastNokartu
                    !nama = namaTX.Text
                    !jenkel = jenkelCMBBX.Text
                    !umur = umurTX.Text
                    !peran = peranCMBBX.Text
                    .Update
                    Exit Sub
                Else
                Exit Sub
                End If
            End If
            
           
            
            
            adnewsiswa
            kosong
        Else
        adnewsiswa
        kosong
        End If
        End With
    Else
        MsgBox "Masih ada yang kosong!"
    End If
End Sub
Sub edit()

With Adodc1.Recordset

If Val(umurTX.Text) <= 0 Then
MsgBox "Umur tidak valid!"
Exit Sub
End If

respons = MsgBox("Perbarui Data?", vbOKCancel)
If respons = vbOK Then
!nokartu = nokartuTX.Text
!nama = namaTX.Text
!jenkel = jenkelCMBBX.Text
!umur = umurTX.Text
!peran = peranCMBBX.Text
.Update
End If
End With

End Sub
'==============================KUMPULAN SUB END











Private Sub pemeriksaanBTN_Click()
fpemeriksaan.Show
fpemeriksaan.namapemeriksaT = namapemeriksaT.Caption
fpemeriksaan.noteleponT = noteleponT.Caption
Unload Me
End Sub

Private Sub refreshBTN_Click()
Adodc1.Refresh
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub

'=====================================TOMBOL FUNGSIONAL
'TOMBOL CREATE
Private Sub tmbhBTN_Click()
'If nokartuTX.Text = Adodc1.Recordset!nokartu Then
'MsgBox "Nomor Kartu sudah ada!"
'Else
buat
'End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
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
daftarBTN.Enabled = False
nokartuTX.Enabled = False
cariTX.Enabled = False
cariBTN.Enabled = False
refreshBTN.Enabled = False
jenkelCMBBX.Enabled = False
tampil
ElseIf updBTN.Caption = "PERBARUI" Then

If nokartuTX.Text <> "" And namaTX.Text <> "" And jenkelCMBBX.Text <> "" And umurTX.Text <> "" And peranCMBBX.Text <> "" Then

edit
updBTN.Caption = "EDIT"
tmbhBTN.Enabled = True
hpsBTN.Enabled = True
Adodc1.Enabled = True
daftarBTN.Enabled = True
nokartuTX.Enabled = True
cariTX.Enabled = True
cariBTN.Enabled = True
refreshBTN.Enabled = True
jenkelCMBBX.Enabled = True
kosong

Else
MsgBox "Masih ada yang kosong!"
End If

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
Adodc1.Recordset.Filter = "nokartu <> '" & nokartu & "'"
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
End If
End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
'TOMBOL DAFTAR PASIEN
Private Sub daftarBTN_Click()
Dim q As New fpendaftaran
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Tidak ada data! Buat data terlebih dahulu!"
Else
With Adodc1.Recordset
q.nokartuT.Caption = .Fields("nokartu").Value
q.pasienT.Caption = .Fields("nama").Value
q.jenkelT.Caption = .Fields("jenkel").Value
q.umurT.Caption = .Fields("umur").Value
q.peranT.Caption = .Fields("peran").Value
q.namapemeriksaT = namapemeriksaT.Caption
q.noteleponT = noteleponT.Caption
q.Show
End With
End If
End Sub
'TOMBOL KEMBALI
Private Sub backBTN_Click()
menu.namapemeriksaT = namapemeriksaT.Caption
menu.noteleponT = noteleponT.Caption
menu.Show
Unload Me
End Sub

'=================================TOMBOL FUNGSIONAL END


'=====================================KEYPRESS
Private Sub nokartuTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
namaTX.SetFocus
End If
End Sub
Private Sub namaTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
umurTX.SetFocus
End If
End Sub
Private Sub umurTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
peranCMBBX.SetFocus
End If
End Sub
Private Sub peranCMBBX_KeyPress(KeyAscii As Integer)
If jenkelCMBBX.Enabled = True Then
If KeyAscii = 13 Then
jenkelCMBBX.SetFocus
End If
End If
End Sub
Private Sub jenkelCMBBX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tmbhBTN.SetFocus
End If
End Sub
'=================================KEYPRESS END

