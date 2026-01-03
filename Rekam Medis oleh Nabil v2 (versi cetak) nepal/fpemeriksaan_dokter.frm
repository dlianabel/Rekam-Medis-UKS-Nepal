VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form fpemeriksaan_dokter 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "fpemeriksaan_dokter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "fpemeriksaan_dokter.frx":2C6F1
      Height          =   1095
      Left            =   4920
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
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
            ColumnWidth     =   374,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   524,976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   585,071
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   840,189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox napasTX 
      Height          =   405
      Left            =   8040
      MultiLine       =   -1  'True
      TabIndex        =   45
      Text            =   "fpemeriksaan_dokter.frx":2C706
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox nadiTX 
      Height          =   405
      Left            =   8040
      MultiLine       =   -1  'True
      TabIndex        =   44
      Text            =   "fpemeriksaan_dokter.frx":2C70C
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox tensiTX 
      Height          =   405
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   43
      Text            =   "fpemeriksaan_dokter.frx":2C712
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox suhuTX 
      Height          =   405
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "fpemeriksaan_dokter.frx":2C718
      Top             =   2640
      Width           =   1215
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
      TabIndex        =   40
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton krgobtBTN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ULANGI INPUT OBAT"
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
      Left            =   2520
      TabIndex        =   39
      Top             =   5880
      Width           =   3375
   End
   Begin VB.CommandButton simpanBTN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SIMPAN"
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
      Left            =   6720
      TabIndex        =   38
      Top             =   6480
      Width           =   5415
   End
   Begin VB.TextBox catTX 
      Height          =   1605
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "fpemeriksaan_dokter.frx":2C71E
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton tmbhBTN 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5400
      TabIndex        =   35
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox jmlobtTX 
      Height          =   405
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "fpemeriksaan_dokter.frx":2C724
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox jmlobtmaxTX 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "fpemeriksaan_dokter.frx":2C72A
      Top             =   5280
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "fpemeriksaan_dokter.frx":2C730
      Left            =   2520
      List            =   "fpemeriksaan_dokter.frx":2C732
      TabIndex        =   32
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox diagnosaTX 
      Height          =   1005
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "fpemeriksaan_dokter.frx":2C734
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox keluhanTX 
      Height          =   645
      Left            =   6720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "fpemeriksaan_dokter.frx":2C73A
      Top             =   1320
      Width           =   5415
   End
   Begin VB.ListBox List2 
      Height          =   840
      ItemData        =   "fpemeriksaan_dokter.frx":2C740
      Left            =   4320
      List            =   "fpemeriksaan_dokter.frx":2C742
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton dtobtBTN 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MUAT ULANG"
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
      Left            =   240
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   7320
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "fpemeriksaan_dokter.frx":2C744
      Height          =   975
      Left            =   1680
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1720
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3600
      Top             =   7320
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
      RecordSource    =   "tbriwayat"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "fpemeriksaan_dokter.frx":2C759
      Height          =   1215
      Left            =   3360
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
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
      ColumnCount     =   16
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
         DataField       =   "diagnosa"
         Caption         =   "diagnosa"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
         DataField       =   "obat"
         Caption         =   "obat"
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
         DataField       =   "catatan"
         Caption         =   "catatan"
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2429,858
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   5280
      Top             =   7320
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   8160
      Top             =   4320
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
      RecordSource    =   "tbtambahobatpasien"
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
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "fpemeriksaan_dokter.frx":2C76E
      Height          =   1575
      Left            =   6720
      TabIndex        =   36
      Top             =   4680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2778
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
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
         DataField       =   "jml"
         Caption         =   "JML"
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
            ColumnWidth     =   1604,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   464,882
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanda-tanda Vital"
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
      Left            =   6720
      TabIndex        =   42
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
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
      Left            =   2520
      TabIndex        =   29
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label5 
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
      Top             =   120
      Width           =   5175
   End
   Begin Project1.PictureG PictureG3 
      Height          =   3345
      Left            =   -120
      Top             =   7200
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   5900
      GIF             =   "fpemeriksaan_dokter.frx":2C783
      Stretch         =   2
   End
   Begin VB.Label pasienT 
      BackStyle       =   0  'Transparent
      Caption         =   "pasien"
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
      Left            =   2640
      TabIndex        =   27
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label2 
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
      Left            =   2520
      TabIndex        =   26
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosa"
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
      Left            =   9480
      TabIndex        =   24
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Resep Obat"
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
      Left            =   6720
      TabIndex        =   23
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Catatan"
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
      Left            =   9480
      TabIndex        =   22
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Obat"
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
      Left            =   2520
      TabIndex        =   21
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
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
      Left            =   9600
      TabIndex        =   20
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label namapemeriksaT 
      Alignment       =   1  'Right Justify
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
      Left            =   9600
      TabIndex        =   19
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   11520
      TabIndex        =   18
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label tanggalT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "dd/mm/yyyy"
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
      Left            =   11520
      TabIndex        =   17
      Top             =   360
      Width           =   2535
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
      Left            =   12840
      TabIndex        =   16
      Top             =   5400
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
      Left            =   12840
      TabIndex        =   15
      Top             =   5160
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
      Left            =   12840
      TabIndex        =   14
      Top             =   4920
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
      Left            =   12840
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label statusT 
      BackStyle       =   0  'Transparent
      Caption         =   "status"
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
      Left            =   12840
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label idrekamT 
      BackStyle       =   0  'Transparent
      Caption         =   "id rekam"
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
      Left            =   12840
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label jamT 
      BackStyle       =   0  'Transparent
      Caption         =   "jam"
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
      Left            =   12840
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   12840
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
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
      Left            =   12480
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label noteleponT 
      Alignment       =   1  'Right Justify
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
      Left            =   12480
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
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
      Top             =   240
      Width           =   5175
   End
   Begin Project1.PictureG PictureG4 
      Height          =   10815
      Left            =   -120
      Top             =   -2760
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   19076
      GIF             =   "fpemeriksaan_dokter.frx":2CE65
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG1 
      Height          =   10935
      Left            =   -360
      Top             =   -2760
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   19288
      GIF             =   "fpemeriksaan_dokter.frx":73BB7
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG2 
      Height          =   3345
      Left            =   -240
      Top             =   -2280
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "fpemeriksaan_dokter.frx":9B435
   End
End
Attribute VB_Name = "fpemeriksaan_dokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBTN_Click()
With Adodc4.Recordset
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
With Adodc1.Recordset
    .MoveFirst ' Pastikan kita mulai dari awal recordset
    If Adodc4.Recordset.RecordCount = 0 Then
    Exit Sub
    Else
    Do While Not .EOF
        If .Fields("nama").Value = Adodc4.Recordset.Fields("nama").Value Then
            .Fields("stok").Value = .Fields("stok").Value + Val(Adodc4.Recordset.Fields("jml").Value)
            .Update
            Exit Do ' Keluar setelah memperbarui
        End If
        .MoveNext
    Loop
Adodc4.Recordset.Delete
End If
End With
.MoveNext
Loop
End If
End With

fpemeriksaan.Show
fpemeriksaan.namapemeriksaT = namapemeriksaT.Caption
fpemeriksaan.noteleponT = noteleponT.Caption
Unload Me
End Sub

Private Sub Command1_Click()
Adodc1.Refresh
Adodc1.Recordset.Filter = "idobat<>"""
End Sub

Private Sub Form_Load()
panggildtobat
kosong
'delall

suhuTX.Text = "Suhu"
suhuTX.ForeColor = &H808080 ' Warna abu-abu
tensiTX.Text = "Tensi"
tensiTX.ForeColor = &H808080
napasTX.Text = "Napas/menit"
napasTX.ForeColor = &H808080
nadiTX.Text = "Nadi/menit"
nadiTX.ForeColor = &H808080

If Adodc1.Recordset.RecordCount = 0 Then
tmbhBTN.Enabled = False
jmlobtTX.Enabled = False
krgobtBTN.Enabled = False
Exit Sub
End If
End Sub

'Tentang Placeholder
Private Sub suhuTX_GotFocus()
If suhuTX.Text = "Suhu" Then
suhuTX.Text = ""
suhuTX.ForeColor = &H0
End If
End Sub
Private Sub suhuTX_LostFocus()
If suhuTX.Text = "" Then
suhuTX.Text = "Suhu"
suhuTX.ForeColor = &H808080
End If
End Sub

Private Sub tensiTX_GotFocus()
If tensiTX.Text = "Tensi" Then
tensiTX.Text = ""
tensiTX.ForeColor = &H0
End If
End Sub
Private Sub tensiTX_LostFocus()
If tensiTX.Text = "" Then
tensiTX.Text = "Tensi"
tensiTX.ForeColor = &H808080
End If
End Sub

Private Sub nadiTX_GotFocus()
If nadiTX.Text = "Nadi/menit" Then
nadiTX.Text = ""
nadiTX.ForeColor = &H0
End If
End Sub
Private Sub nadiTX_LostFocus()
If nadiTX.Text = "" Then
nadiTX.Text = "Nadi/menit"
nadiTX.ForeColor = &H808080
End If
End Sub

Private Sub napasTX_GotFocus()
If napasTX.Text = "Napas/menit" Then
napasTX.Text = ""
napasTX.ForeColor = &H0
End If
End Sub
Private Sub napasTX_LostFocus()
If napasTX.Text = "" Then
napasTX.Text = "Napas/menit"
napasTX.ForeColor = &H808080
End If
End Sub
'tentangplaceholder




'==================================KUMPULAN SUB
Sub kosong()
diagnosaTX.Text = ""
catTX.Text = ""
jmlobtTX.Text = ""
jmlobtmaxTX.Text = ""
End Sub

Sub panggildtobat()
With Adodc1.Recordset
Command1_Click
Do While Not Adodc1.Recordset.EOF
    List1.AddItem Adodc1.Recordset.Fields("Nama").Value
    Adodc1.Recordset.MoveNext
Loop
End With
End Sub

'MEGISI DATA PASIEN (SPESIFIK : STATUS = "Selesai") tabel pemeriksaan
Sub isdtpas()
With Adodc3.Recordset
.MoveLast
            'update 18 oktober 2024
            !catatan = catTX
            !notelpemeriksa = noteleponT
            !jenkel = jenkelT
            !umur = umurT
            !peran = peranT
            !Status = "Selesai"
            !keluhan = keluhanTX
            !diagnosa = diagnosaTX
!Status = "Selesai"
.Update
Adodc3.Refresh
End With
End Sub

Sub delall()
With Adodc4.Recordset
n = .RecordCount
If n > 0 Then
.MoveFirst
Do While Not .EOF
.Delete
.MoveNext
Loop
End If
End With
End Sub
'==============================KUMPULAN SUB END



'===============================KUMPULAN TOMBOL
'PERGI KE FORM DATA OBAT
Private Sub dtobtBTN_Click()
fdataobat.Show
End Sub





Private Sub List1_Click()
jmlobtmaxTX.Text = List1.Text
With Adodc1.Recordset
.MoveFirst
.Find "nama = '" & List1.Text & "'"
jmlobtmaxTX.Text = !stok & " (Max)"
End With
End Sub



'TAMBAH OBAT
Private Sub tmbhBTN_Click()
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "nama = '" & List1.List(List1.ListIndex) & "'"
If jmlobtTX.Text = "" Then
MsgBox "Jumlah masih kosong!"
Exit Sub
End If
If Val(jmlobtTX.Text) > Adodc1.Recordset!stok Then
MsgBox "Jumlah obat kelebihan!"
ElseIf Val(jmlobtTX.Text) <= 0 Then
MsgBox "Jumlah obat tidak valid!"
Else
If jmlobtTX.Text = "" Then
MsgBox "Jumlah obat belum diisi!"
Else

With Adodc4.Recordset
    If Not .EOF And Not .BOF Then
        .MoveFirst
    End If
    Dim obatAda As Boolean
    obatAda = False

    ' Cek apakah nama obat sudah ada
    Do While Not .EOF And Not .BOF
        If .Fields("nama").Value = List1.List(List1.ListIndex) Then
            ' Jika sudah ada, tambah jumlah
            .Fields("jml").Value = .Fields("jml").Value + Val(jmlobtTX.Text)
            .Update
            obatAda = True
            Exit Do
        End If
        .MoveNext
    Loop

    ' Jika obat belum ada, maka tambahkan yang baru
    If Not obatAda Then
        .AddNew
        !nama = List1.List(List1.ListIndex)
        !jml = Val(jmlobtTX.Text)
        .Update
    End If
End With

' Mengupdate stok di Adodc1
With Adodc1.Recordset
    .MoveFirst
    Do While Not .EOF
        If .Fields("nama").Value = List1.List(List1.ListIndex) Then
            .Fields("stok").Value = .Fields("stok").Value - Val(jmlobtTX.Text)
            .Update
            Exit Do
        End If
        .MoveNext
    Loop
End With
End If
End If
List1_Click
End Sub
'KURANGI OBAT
Private Sub krgobtBTN_Click()
With Adodc1.Recordset
    .MoveFirst ' Pastikan kita mulai dari awal recordset
    If Adodc4.Recordset.RecordCount = 0 Then
    Exit Sub
    Else
    Do While Not .EOF
        If .Fields("nama").Value = Adodc4.Recordset.Fields("nama").Value Then
            .Fields("stok").Value = .Fields("stok").Value + Val(Adodc4.Recordset.Fields("jml").Value)
            .Update
            Exit Do ' Keluar setelah memperbarui
        End If
        .MoveNext
    Loop
Adodc4.Recordset.Delete
End If
End With
List1_Click
End Sub
'TOMBOL SIMPAN = PERGI KE FORM RIWAYAT
Private Sub simpanBTN_Click()
If diagnosaTX.Text = "" Then
        MsgBox "Diagnosa belum diisi!"
        Exit Sub
    End If
    
    Dim q As New friwayat

With Adodc2.Recordset
If Adodc4.Recordset.RecordCount = 0 Then
            .AddNew
            !idrekam = idrekamT
            !nokartu = nokartuT
            '!nama = pasienT
            '!tanggal = tanggalT
            '!jam = jamT
            '!pemeriksa = namapemeriksaT
            !obat = ""
            !jmlobat = ""
            '!idpasien = idpasienT
            'update 16 oktober 2024
            !suhu = suhuTX.Text
            !tensi = tensiTX.Text
            !nadi = nadiTX.Text
            !napas = napasTX.Text
            .Update
    With Adodc3.Recordset
            .MoveFirst
            !diagnosa = diagnosaTX.Text
            !catatan = catTX.Text
            .Update
    End With
    
Else
    If Not Adodc4.Recordset.EOF Then
        Adodc4.Recordset.MoveFirst
        Do While Not Adodc4.Recordset.EOF
            .AddNew
            !idrekam = idrekamT
            !nokartu = nokartuT
            '!tanggal = tanggalT
            '!jam = jamT
            !obat = Adodc4.Recordset!nama ' Menggunakan nama dari Adodc4
            !jmlobat = Adodc4.Recordset!jml ' Menggunakan jml dari Adodc4
            '!idpasien = idpasienT
            'update 16 oktober 2024
            !suhu = suhuTX.Text
            !tensi = tensiTX.Text
            !nadi = nadiTX.Text
            !napas = napasTX.Text
            .Update
            
            ' Pindah ke record berikutnya di Adodc4
            Adodc4.Recordset.MoveNext
        Loop
        With Adodc3.Recordset
            .MoveFirst
            !diagnosa = diagnosaTX.Text
            !catatan = catTX.Text
            .Update
        End With
    End If
End If
End With
delall
    Adodc2.Refresh
    isdtpas
    q.Show
    q.namapemeriksaT = namapemeriksaT.Caption
    q.noteleponT = noteleponT.Caption
    q.Adodc1.Recordset.MoveLast
    Unload Me
End Sub
'===========================KUMPULAN TOMBOL END

Private Sub jmlobtTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

End If
End Sub


