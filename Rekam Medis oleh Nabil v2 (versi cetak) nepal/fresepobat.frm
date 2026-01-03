VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form fresepobat 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "fresepobat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox keluhanTX 
      Height          =   2205
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "fresepobat.frx":2C6F1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox napasTX 
      Height          =   405
      Left            =   9240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "fresepobat.frx":2C6F7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox nadiTX 
      Height          =   405
      Left            =   9240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "fresepobat.frx":2C6FD
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox tensiTX 
      Height          =   405
      Left            =   7920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "fresepobat.frx":2C703
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox suhuTX 
      Height          =   405
      Left            =   7920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "fresepobat.frx":2C709
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox catTX 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   11760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "fresepobat.frx":2C70F
      Top             =   4680
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "fresepobat.frx":2C715
      Left            =   10200
      List            =   "fresepobat.frx":2C717
      TabIndex        =   26
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "fresepobat.frx":2C719
      Left            =   7800
      List            =   "fresepobat.frx":2C71B
      TabIndex        =   25
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox diagnosaTX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   10800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "fresepobat.frx":2C71D
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cetakBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "CETAK RESEP OBAT"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6960
      Width           =   2295
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
      TabIndex        =   28
      Top             =   7680
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "fresepobat.frx":2C723
      Height          =   2055
      Left            =   9720
      TabIndex        =   0
      Top             =   -1080
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      ColumnCount     =   17
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
         Caption         =   "OBAT"
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
         DataField       =   "jmlobat"
         Caption         =   "JUMLAH"
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
      BeginProperty Column16 
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
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2670,236
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   14,74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   12960
      Top             =   960
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7800
      Top             =   480
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "fresepobat.frx":2C738
      Height          =   1575
      Left            =   5880
      TabIndex        =   38
      Top             =   -1080
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2778
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
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Tensi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   255
      Left            =   7920
      TabIndex        =   35
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Nadi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   255
      Left            =   9240
      TabIndex        =   36
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Suhu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   255
      Left            =   7920
      TabIndex        =   34
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      Left            =   11760
      TabIndex        =   5
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label17 
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
      Left            =   7920
      TabIndex        =   30
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah obat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   255
      Left            =   8880
      TabIndex        =   22
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Obat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label tanggalT 
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
      Left            =   4320
      TabIndex        =   20
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label13 
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
      Left            =   4320
      TabIndex        =   19
      Top             =   1920
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
      Left            =   4200
      TabIndex        =   18
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label11 
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
      Left            =   4200
      TabIndex        =   17
      Top             =   3360
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
      Left            =   4200
      TabIndex        =   16
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label9 
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
      Left            =   4200
      TabIndex        =   15
      Top             =   2760
      Width           =   2535
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
      Left            =   4200
      TabIndex        =   14
      Top             =   4440
      Width           =   2535
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
      Left            =   4200
      TabIndex        =   13
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label10 
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
      Left            =   4200
      TabIndex        =   11
      Top             =   5040
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
      Left            =   4200
      TabIndex        =   10
      Top             =   5280
      Width           =   2895
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
      Left            =   4200
      TabIndex        =   9
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "No Kartu"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "UKS SMK N 1 Pemalang"
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
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jl. Gatot Subroto No.31, Bojongbata, Kec. Pemalang, Kabupaten Pemalang, Jawa Tengah 52319"
      ForeColor       =   &H00422A0D&
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   1440
      Width           =   9375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Left            =   7800
      TabIndex        =   4
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label Label5 
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
      Left            =   10800
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Napas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   255
      Left            =   9240
      TabIndex        =   37
      Top             =   3120
      Width           =   2535
   End
   Begin Project1.PictureG PictureG2 
      Height          =   11535
      Left            =   -360
      Top             =   -3240
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   20346
      GIF             =   "fresepobat.frx":2C74D
      Mirror          =   1
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resep Obat"
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
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resep Obat"
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
      TabIndex        =   12
      Top             =   0
      Width           =   5055
   End
   Begin Project1.PictureG PictureG1 
      Height          =   9195
      Left            =   1920
      Top             =   -480
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   16219
      GIF             =   "fresepobat.frx":113663
      Stretch         =   2
      Mirror          =   3
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   10695
      Left            =   7680
      Shape           =   2  'Oval
      Top             =   -5880
      Width           =   17775
   End
   Begin Project1.PictureG PictureG4 
      Height          =   11760
      Left            =   -600
      Top             =   -3360
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   20743
      GIF             =   "fresepobat.frx":118C39
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG5 
      Height          =   6060
      Left            =   -840
      Top             =   7080
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   10689
      GIF             =   "fresepobat.frx":15F98B
      Stretch         =   2
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG3 
      Height          =   3825
      Left            =   -3840
      Top             =   -2520
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   6747
      GIF             =   "fresepobat.frx":16006D
   End
End
Attribute VB_Name = "fresepobat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub tampil()
'update 18 oktober 2024
With Adodc1.Recordset
namapemeriksaT.Caption = !pemeriksa
catTX.Text = !catatan
keluhanTX.Text = !keluhan
diagnosaTX.Text = !diagnosa
noteleponT.Caption = !notelpemeriksa
pasienT.Caption = !nama
peranT.Caption = !peran
nokartuT.Caption = !nokartu
tanggalT.Caption = !tanggal
End With

With Adodc2.Recordset
n = .RecordCount
If n > 0 Then
.MoveFirst
For i = 1 To n
If Not .EOF Then
List2.AddItem !obat
List1.AddItem !jmlobat
'update 16 oktober 2024
suhuTX.Text = !suhu
tensiTX.Text = !tensi
nadiTX.Text = !nadi
napasTX.Text = !napas
.MoveNext
End If
Next i
End If

End With
End Sub

Private Sub backBTN_Click()
Unload Me
End Sub


Private Sub cetakBTN_Click()
respons = MsgBox("Cetak obat pasien dengan nama '" & pasienT.Caption & "'?", vbOKCancel)
If respons = vbOK Then
Static WDOCX As Word.Application
Static WDOCX1 As Word.Document
Set WDOCX = New Word.Application
WDOCX.Visible = True
WDOCX.Activate
Set WDOCX1 = WDOCX.Documents.Add(App.Path & "\cetak_uks.dotx")

With WDOCX1



.FormFields("w_tanggal").Range = tanggalT.Caption
.FormFields("w_pemeriksa").Range = namapemeriksaT.Caption
.FormFields("w_notelp").Range = noteleponT.Caption
.FormFields("w_pasien").Range = pasienT.Caption
.FormFields("w_peran").Range = peranT.Caption
.FormFields("w_nk").Range = nokartuT.Caption

.FormFields("w_suhu").Range = suhuTX.Text
.FormFields("w_tensi").Range = tensiTX.Text
.FormFields("w_nadi").Range = nadiTX.Text
.FormFields("w_napas").Range = napasTX.Text

.FormFields("w_diagnosa").Range = diagnosaTX.Text
.FormFields("w_catatan").Range = catTX.Text


obatList = "" ' Mulai dengan string kosong
' Menggabungkan semua item dari List2 (obat) menjadi satu string
        For i = 0 To List2.ListCount - 1
            If i > 0 Then
                obatList = obatList & vbCrLf ' Menambahkan baris baru setelah setiap obat
            End If
            ' Menggabungkan nama obat dan jumlah obat (misalnya "Obat 1 - 3 pcs")
            obatList = obatList & List2.List(i) & " - " & List1.List(i)
        Next i
.FormFields("w_obat").Range = obatList
        
End With
End If



 'DataReport1.Sections("Section2").Controls("labelpasienR").Caption = pasienT.Caption
        'DataReport1.Sections("Section2").Controls("labelpemeriksaR").Caption = namapemeriksaT.Caption
'Adodc2.Recordset.Filter = "idrekam = '" & Adodc1.Recordset!idRekam & "'"
'DataReport1.Show



'With Adodc1.Recordset
'respons = MsgBox("Cetak obat pasien dengan nama '" & pasienT.Caption & "'?", vbOKCancel)
'If respons = vbOK Then
'Printer.Print
'Me.PrintForm
'Printer.EndDoc
'End If
'End With
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub

