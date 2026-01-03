VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form f1datapetugas 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "f1datapetugas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox jmldataTX 
      Height          =   405
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   6720
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
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox cariTX 
      Height          =   405
      Left            =   10440
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1440
      Width           =   1935
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton konfirmBTN 
      BackColor       =   &H00EB9E3F&
      Caption         =   "CEK"
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
      TabIndex        =   18
      Top             =   7560
      Visible         =   0   'False
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox jenkelCMBBX 
      Height          =   315
      ItemData        =   "f1datapetugas.frx":2C6F1
      Left            =   960
      List            =   "f1datapetugas.frx":2C6FB
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox nomTX 
      Height          =   405
      Left            =   960
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox namaTX 
      Height          =   405
      Left            =   960
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox passTX 
      Height          =   405
      Left            =   960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox userTX 
      Height          =   405
      Left            =   960
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2040
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "f1datapetugas.frx":2C715
      Height          =   4695
      Left            =   4320
      TabIndex        =   0
      Top             =   1920
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
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
            ColumnWidth     =   1620,284
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
      Left            =   4320
      Top             =   1440
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
   Begin Project1.PictureG PictureG1 
      Height          =   3345
      Left            =   2520
      Top             =   7320
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "f1datapetugas.frx":2C72A
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Petugas / Pemeriksa"
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
      Left            =   6600
      TabIndex        =   9
      Top             =   0
      Width           =   7215
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Petugas / Pemeriksa"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   0
      Width           =   7215
   End
   Begin Project1.PictureG PictureG2 
      Height          =   3345
      Left            =   3600
      Top             =   -2280
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "f1datapetugas.frx":2CE0C
      Mirror          =   2
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00422A0D&
      FillStyle       =   0  'Solid
      Height          =   5895
      Left            =   4200
      Top             =   1320
      Width           =   9495
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
      Left            =   960
      TabIndex        =   5
      Top             =   4560
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
      Left            =   960
      TabIndex        =   4
      Top             =   5520
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
      Left            =   960
      TabIndex        =   3
      Top             =   3600
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
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
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
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin Project1.PictureG PictureG4 
      Height          =   11745
      Left            =   240
      Top             =   -1680
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   20717
      GIF             =   "f1datapetugas.frx":2D4EE
   End
   Begin Project1.PictureG PictureG3 
      Height          =   3855
      Left            =   3240
      Top             =   2160
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6800
      GIF             =   "f1datapetugas.frx":328BC
   End
   Begin Project1.PictureG PictureG6 
      Height          =   10650
      Left            =   -360
      Top             =   -2280
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   18785
      GIF             =   "f1datapetugas.frx":328D4
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG5 
      Height          =   10785
      Left            =   -480
      Top             =   -2160
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   19024
      GIF             =   "f1datapetugas.frx":6D4AE
      Mirror          =   1
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E09C48&
      FillStyle       =   0  'Solid
      Height          =   8295
      Left            =   7080
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "f1datapetugas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cariBTN_Click()
If cariTX.Text = "" Then
    Adodc1.Refresh
Else
    Adodc1.Recordset.Filter = "nama = '" & cariTX.Text & "' OR jenkel = '" & cariTX.Text & "' OR username = '" & cariTX.Text & "' OR notelp = '" & cariTX.Text & "'"
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
'updBTN.Enabled = False
'hpsBTN.Enabled = False
End Sub



'==================================KUMPULAN SUB
Sub kosong()
userTX.Text = ""
passTX.Text = ""
namaTX.Text = ""
nomTX.Text = ""
'jenkelCMBBX.Text = ""
jenkelCMBBX.ListIndex = -1
cariTX.Text = ""
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
            If .RecordCount > 0 Then
            .MoveFirst ' Pastikan cursor berada di awal recordset
            .Find "UserName = '" & userTX.Text & "'"
            If Not .EOF Then
                MsgBox "Username sudah ada!"
                Exit Sub
            Else
                .AddNew
                !UserName = userTX.Text
                !Password = passTX.Text
                !nama = namaTX.Text
                !notelp = nomTX.Text
                !jenkel = jenkelCMBBX.Text
                .Update
                kosong
            End If
            Else
                .AddNew
                !UserName = userTX.Text
                !Password = passTX.Text
                !nama = namaTX.Text
                !notelp = nomTX.Text
                !jenkel = jenkelCMBBX.Text
                .Update
                kosong
            End If
        End With
    Else
        MsgBox "Masih ada yang kosong!"
    End If
End Sub
Sub edit()
If Val(nomTX.Text) <= 0 Then
    MsgBox "Nomor telepon tidak valid!"
    Exit Sub
End If
With Adodc1.Recordset
respons = MsgBox("Perbarui Data?", vbOKCancel)
If respons = vbOK Then
!UserName = userTX.Text
!Password = passTX.Text
!nama = namaTX.Text
!notelp = nomTX.Text
!jenkel = jenkelCMBBX.Text
.Update
End If
End With
End Sub
'==============================KUMPULAN SUB END






Private Sub konfirmBTN_Click()
konfirmuser.Show
End Sub

Private Sub refreshBTN_Click()
Adodc1.Refresh
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub

'=====================================TOMBOL FUNGSIONAL
'TOMBOL CREATE
Private Sub tmbhBTN_Click()
If Val(nomTX.Text) <= 0 Then
    MsgBox "Nomor telepon tidak valid!"
    Exit Sub
End If
buat
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
userTX.Enabled = False
cariTX.Enabled = False
cariBTN.Enabled = False
refreshBTN.Enabled = False
jenkelCMBBX.Enabled = False
tampil
ElseIf updBTN.Caption = "PERBARUI" Then

If userTX.Text <> "" And passTX.Text <> "" And namaTX.Text <> "" And jenkelCMBBX.Text <> "" And nomTX.Text <> "" Then

edit
updBTN.Caption = "EDIT"
tmbhBTN.Enabled = True
hpsBTN.Enabled = True
Adodc1.Enabled = True
userTX.Enabled = True
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
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
'TOMBOL DELETE
Private Sub hpsBTN_Click()
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Tidak ada data! Buat data terlebih dahulu!"
Else
respons = MsgBox("Hapus Data?", vbOKCancel)
If respons = vbOK Then
Adodc1.Recordset.Delete
Adodc1.Recordset.Filter = "username <>'" & UserName & "'"
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If
End If
End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
'TOMBOL KEMBALI
Private Sub backBTN_Click()
If Adodc1.Recordset.RecordCount > 0 Then
menu.Show
menu.namapemeriksaT = namapemeriksaT.Caption
menu.noteleponT = noteleponT.Caption
Unload Me
Else
login.Show
Unload menu
Unload Me
End If
End Sub

'=================================TOMBOL FUNGSIONAL END

'=====================================KEYPRESS
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


