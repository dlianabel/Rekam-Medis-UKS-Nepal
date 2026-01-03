VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form f2dataobatstok 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "f2dataobatstok.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox jmldataTX 
      Height          =   405
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   6960
      Width           =   1215
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
      TabIndex        =   22
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton resettotalBTN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RESET TOTAL"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6000
      Top             =   7680
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
   Begin VB.CommandButton hitkuBTN 
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton hittaBTN 
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
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox namaTX 
      Enabled         =   0   'False
      Height          =   405
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox satuanTX 
      Enabled         =   0   'False
      Height          =   405
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cekBTN 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "f2dataobatstok.frx":2C6F1
      Height          =   2655
      Left            =   3120
      TabIndex        =   12
      Top             =   7680
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4683
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
         Caption         =   "idobat"
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
      BeginProperty Column02 
         DataField       =   "kegunaan"
         Caption         =   "kegunaan"
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
         Caption         =   "satuan"
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
         Caption         =   "stok"
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
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1275,024
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox namaCMBBX 
      Height          =   315
      ItemData        =   "f2dataobatstok.frx":2C706
      Left            =   5640
      List            =   "f2dataobatstok.frx":2C708
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox cariTX 
      Height          =   405
      Left            =   10440
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1320
      Width           =   2175
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
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox jmlobatTX 
      Height          =   405
      Left            =   600
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox totalstokTX 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox stokawalTX 
      Enabled         =   0   'False
      Height          =   405
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton tmbhBTN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TAMBAH STOK"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   2175
   End
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
      Left            =   12840
      TabIndex        =   0
      Top             =   7680
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "f2dataobatstok.frx":2C70A
      Height          =   5055
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8916
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "idstokobat"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "stokawal"
         Caption         =   "AWAL"
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
         DataField       =   "stokmasuk"
         Caption         =   "MASUK"
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
         DataField       =   "stokkeluar"
         Caption         =   "KELUAR"
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
         DataField       =   "stoktotal"
         Caption         =   "TOTAL"
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
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3480
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
      RecordSource    =   "tbstokobat"
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stok"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   27
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stok"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label9 
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
      TabIndex        =   25
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label10 
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
      TabIndex        =   26
      Top             =   360
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00422A0D&
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   3360
      Top             =   1200
      Width           =   10575
   End
   Begin Project1.PictureG PictureG5 
      Height          =   11760
      Left            =   7200
      Top             =   -240
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   20743
      GIF             =   "f2dataobatstok.frx":2C71F
      Mirror          =   3
   End
   Begin Project1.PictureG PictureG1 
      Height          =   11760
      Left            =   4320
      Top             =   -3360
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   20743
      GIF             =   "f2dataobatstok.frx":73471
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Stok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Awal"
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
      Left            =   600
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Obat"
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
      Left            =   600
      TabIndex        =   9
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Obat"
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
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
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
   Begin Project1.PictureG PictureG4 
      Height          =   3345
      Left            =   2400
      Top             =   7680
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "f2dataobatstok.frx":BA1C3
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG2 
      Height          =   3345
      Left            =   -2040
      Top             =   -2400
      Width           =   16125
      _ExtentX        =   28443
      _ExtentY        =   5900
      GIF             =   "f2dataobatstok.frx":BA8A5
      Stretch         =   2
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   10695
      Left            =   -13080
      Shape           =   2  'Oval
      Top             =   5280
      Width           =   17775
   End
End
Attribute VB_Name = "f2dataobatstok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cariBTN_Click()
If cariTX.Text = "" Then
    Adodc1.Recordset.Filter = "" ' Menampilkan semua data
    Adodc1.Refresh
Else
Adodc1.Recordset.Filter = "idstokobat = '" & cariTX.Text & "' OR tanggal = '" & cariTX.Text & "' OR jam = '" & cariTX.Text & "' OR nama = '" & cariTX.Text & "' OR satuan = '" & cariTX.Text & "' OR stokawal = '" & cariTX.Text & "' OR stokmasuk = '" & cariTX.Text & "' OR stoktotal = '" & cariTX.Text & "'"
End If
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
Private Sub cariTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cariBTN_Click
End If
End Sub


Private Sub cekBTN_Click()
With Adodc2.Recordset
.MoveFirst
.Find "nama = '" & namaTX.Text & "'"
satuanTX.Text = .Fields("satuan").Value
stokawalTX.Text = .Fields("stok").Value
idobatTX.Caption = .Fields("idobat")
End With
End Sub

Private Sub Form_Load()
kosong
tmbhBTN.Enabled = False
resettotalBTN.Enabled = False
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub

'==================================KUMPULAN SUB
Sub kosong()
'namaCMBBX.Text = ""
namaTX.Text = ""
satuanTX.Text = ""
stokawalTX.Text = ""
jmlobatTX.Text = ""
totalstokTX.Text = ""
cariTX.Text = ""
End Sub
Sub summonobat()
    If Not (Adodc2.Recordset.EOF And Adodc2.Recordset.BOF) Then
        Adodc2.Recordset.MoveFirst
        Do While Not Adodc2.Recordset.EOF
            namaCMBBX.AddItem Adodc2.Recordset.Fields("Nama").Value
            Adodc2.Recordset.MoveNext
        Loop
    End If
End Sub
Sub tambahstok()
With Adodc1.Recordset
wjam = Format(Now, "HH:mm:ss")
maksimalrek = 10000000
If .RecordCount >= maksimalrek Then
MsgBox "Jumlah data sudah mencapai batas maksimal"
Exit Sub
End If
If .EOF Then
urutan = "SO" + "0000001"
Else
.MoveLast
nomorurut = CInt(Mid(!idstokobat, 3)) 'contoh : RM1234567
nomorurut = nomorurut + 1
urutan = "SO" + Right("0000000" & nomorurut, 7)
End If

.AddNew
!idstokobat = urutan
!tanggal = Format(Now, "dd/mm/yyyy")
!jam = wjam
!nama = namaTX.Text
!satuan = satuanTX.Text
!stokawal = stokawalTX.Text
!stokmasuk = jmlobatTX.Text
!stokkeluar = 0
!stoktotal = totalstokTX.Text
.Update
End With
With Adodc2.Recordset
!stok = totalstokTX.Text
.Update
End With
End Sub
Sub kurangstok()
With Adodc1.Recordset
wjam = Format(Now, "HH:mm:ss")
maksimalrek = 10000000
If .RecordCount >= maksimalrek Then
MsgBox "Jumlah data sudah mencapai batas maksimal"
Exit Sub
End If
If .EOF Then
urutan = "SO" + "0000001"
Else
.MoveLast
nomorurut = CInt(Mid(!idstokobat, 3)) 'contoh : RM1234567
nomorurut = nomorurut + 1
urutan = "SO" + Right("0000000" & nomorurut, 7)
End If

.AddNew
!idstokobat = urutan
!tanggal = Format(Now, "dd/mm/yyyy")
!jam = wjam
!nama = namaTX.Text
!satuan = satuanTX.Text
!stokawal = stokawalTX.Text
!stokmasuk = 0
!stokkeluar = jmlobatTX.Text
!stoktotal = totalstokTX.Text
.Update
End With
With Adodc2.Recordset
!stok = totalstokTX.Text
.Update
End With
End Sub
'==============================KUMPULAN SUB END






'=====================================TOMBOL FUNGSIONAL
'TOMBOL TUTUP
Private Sub backBTN_Click()
f2dataobat.Show
f2dataobat.namapemeriksaT = namapemeriksaT.Caption
f2dataobat.noteleponT = noteleponT.Caption
Unload Me
End Sub

Private Sub hittaBTN_Click()
If jmlobatTX.Text = "" Then
Exit Sub
End If
If Val(jmlobatTX.Text) <= 0 Then
MsgBox "Jumlah obat tidak valid!"
Else
totalstokTX.Text = Val(stokawalTX.Text) + Val(jmlobatTX.Text)
tmbhBTN.Caption = "TAMBAH STOK"
tmbhBTN.Enabled = True
resettotalBTN.Enabled = True
End If
End Sub

Private Sub hitkuBTN_Click()
If jmlobatTX.Text = "" Then
Exit Sub
End If
If Val(jmlobatTX.Text) > Val(stokawalTX.Text) Then
MsgBox "Jumlah obat kelebihan!"
ElseIf Val(jmlobatTX.Text) <= 0 Then
MsgBox "Jumlah obat tidak valid!"
Else
totalstokTX.Text = Val(stokawalTX.Text) - Val(jmlobatTX.Text)
tmbhBTN.Caption = "KURANGI STOK"
tmbhBTN.Enabled = True
resettotalBTN.Enabled = True
End If
End Sub


Private Sub stokmasukTX_KeyPress(KeyAscii As Integer)
With Adodc2.Recordset
.MoveFirst
.Find "nama = '" & namaTX.Text & "'"
satuanTX.Text = .Fields("satuan").Value
stokawalTX.Text = .Fields("stok").Value
End With
End Sub

Private Sub refreshBTN_Click()
Adodc1.Refresh
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub

Private Sub resettotalBTN_Click()
totalstokTX.Text = stokawalTX.Text
jmlobatTX.Text = ""
tmbhBTN.Enabled = False
End Sub

Private Sub tmbhBTN_Click()
If tmbhBTN.Caption = "TAMBAH STOK" Then
tambahstok
ElseIf tmbhBTN.Caption = "KURANGI STOK" Then
kurangstok
End If
kosong
tmbhBTN.Enabled = False
resettotalBTN.Enabled = False
hittaBTN.Enabled = False
hitkuBTN.Enabled = False
jmlobatTX.Enabled = False
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub


'=================================TOMBOL FUNGSIONAL END


'=====================================KEYPRESS
'=================================KEYPRESS END

