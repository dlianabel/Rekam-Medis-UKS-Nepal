VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form friwayat 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "friwayat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox jmldataTX 
      Height          =   405
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton sortirBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "SORTIR"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox tglakhirTX 
      Height          =   405
      Left            =   3000
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox tglawalTX 
      Height          =   405
      Left            =   1680
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6480
      Width           =   975
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   7680
      Width           =   1455
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox cariTX 
      Height          =   405
      Left            =   9240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1200
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton lhtBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "LIHAT"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   12720
      Top             =   4560
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
      Bindings        =   "friwayat.frx":2C6F1
      Height          =   735
      Left            =   12720
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1620,284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   390,047
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   450,142
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   975,118
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "friwayat.frx":2C706
      Height          =   4695
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
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
            ColumnWidth     =   3495,118
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
      Left            =   1680
      Top             =   1200
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
      CommandType     =   8
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
      RecordSource    =   "select * from tbpemeriksaan where status=""Selesai"";"
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
   Begin VB.Label jmlDataL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "jml"
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
      Left            =   7440
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2520
      TabIndex        =   15
      Top             =   6360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00422A0D&
      FillStyle       =   0  'Solid
      Height          =   6015
      Left            =   1560
      Top             =   1080
      Width           =   11175
   End
   Begin Project1.PictureG PictureG6 
      Height          =   9405
      Left            =   -120
      Top             =   -1440
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   16589
      GIF             =   "friwayat.frx":2C71B
      Mirror          =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Riwayat"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Riwayat"
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
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
   Begin Project1.PictureG PictureG3 
      Height          =   3810
      Left            =   1440
      Top             =   7200
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   6720
      GIF             =   "friwayat.frx":672F5
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
      Height          =   6930
      Left            =   8520
      Top             =   0
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   12224
      GIF             =   "friwayat.frx":679D7
   End
   Begin Project1.PictureG PictureG1 
      Height          =   3345
      Left            =   3720
      Top             =   -2280
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "friwayat.frx":1809C1
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG4 
      Height          =   11040
      Left            =   6720
      Top             =   -1440
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   19473
      GIF             =   "friwayat.frx":1810A3
   End
   Begin Project1.PictureG PictureG2 
      Height          =   9225
      Left            =   120
      Top             =   -1200
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   16272
      GIF             =   "friwayat.frx":186471
      Mirror          =   1
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7F21&
      FillStyle       =   0  'Solid
      Height          =   6735
      Left            =   -1320
      Shape           =   2  'Oval
      Top             =   600
      Width           =   7455
   End
End
Attribute VB_Name = "friwayat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"

'tentang placeholder
tglawalTX.Text = "Tgl Awal"
tglawalTX.ForeColor = &H808080 ' Warna abu-abu
tglakhirTX.Text = "Tgl Akhir"
tglakhirTX.ForeColor = &H808080
End Sub

Sub kosong()
cariTX.Text = ""
tglawalTX.Text = ""
tglakhirTX.Text = ""
End Sub

Private Sub backBTN_Click()
menu.Show
menu.namapemeriksaT = namapemeriksaT.Caption
menu.noteleponT = noteleponT.Caption
Unload Me
Unload Me
End Sub



Private Sub cariBTN_Click()
Dim tanggalCari As Date

If cariTX.Text = "" Then
Adodc1.Refresh ' Menampilkan semua data
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
Exit Sub
End If

If IsDate(cariTX.Text) Then
tanggalCari = CDate(cariTX.Text)
End If
Adodc1.Recordset.Filter = "idrekam = '" & cariTX.Text & "' OR " & _
                          "nokartu = '" & cariTX.Text & "' OR " & _
                          "nama = '" & cariTX.Text & "' OR " & _
                          "jenkel = '" & cariTX.Text & "' OR " & _
                          "umur = '" & cariTX.Text & "' OR " & _
                          "peran = '" & cariTX.Text & "' OR " & _
                          "keluhan = '" & cariTX.Text & "' OR " & _
                          "diagnosa = '" & cariTX.Text & "' OR " & _
                          "catatan = '" & cariTX.Text & "' OR " & _
                          "tanggal = #" & Format(tanggalCari, "MM/DD/YYYY") & "# OR " & _
                          "jam = '" & cariTX.Text & "' OR " & _
                          "pemeriksa = '" & cariTX.Text & "' OR " & _
                          "notelpemeriksa = '" & cariTX.Text & "'"
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub
Private Sub cariTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cariBTN_Click
End If
End Sub



Private Sub sortirBTN_Click()
If tglawalTX.Text = "" And tglakhirTX.Text = "" Then
MsgBox "Kolom sortir masih ada yang kosong!"
tglawalTX.SetFocus
Exit Sub
End If

With Adodc1.Recordset
IsValid = IsDate(tglawalTX.Text) And IsDate(tglakhirTX.Text)
If Not IsValid Then
MsgBox "Format tanggal tidak valid! Harap masukkan dalam format dd/mm/yyyy."
Exit Sub
End If

'jika lebih dari bulan 12
tglawal = CInt(Mid(tglawalTX.Text, InStr(tglawalTX.Text, "/") + 1, InStrRev(tglawalTX.Text, "/") - InStr(tglawalTX.Text, "/") - 1))
tglakhir = CInt(Mid(tglakhirTX.Text, InStr(tglakhirTX.Text, "/") + 1, InStrRev(tglakhirTX.Text, "/") - InStr(tglakhirTX.Text, "/") - 1))
If tglawal < 1 Or tglawal > 12 Then
MsgBox "Bulan salah!"
Exit Sub
End If
If tglakhir < 1 Or tglakhir > 12 Then
MsgBox "Bulan salah!"
Exit Sub
End If


.Filter = "tanggal >= #" & tglawalTX.Text & "# AND tanggal <= #" & tglakhirTX.Text & "#"
End With
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub


Private Sub dtpasienBTN_Click()
f3datapasien.Show
f3datapasien.namapemeriksaT = namapemeriksaT.Caption
f3datapasien.noteleponT = noteleponT.Caption
Unload Me
End Sub

Private Sub lhtBTN_Click()
Dim q As New fresepobat
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "Tidak ada data! Buat data terlebih dahulu!"
Else
q.Adodc2.Recordset.Filter = "idrekam='" & Adodc1.Recordset!idrekam & "'"
q.Adodc1.Recordset.Filter = "idrekam='" & Adodc1.Recordset!idrekam & "'"
q.Show
q.tampil
End If
End Sub



Private Sub refreshBTN_Click()
Adodc1.Refresh
kosong
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub





Private Sub tglawalTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tglakhirTX.SetFocus
End If
End Sub

Private Sub tglakhirTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sortirBTN_Click
End If
End Sub


'Tentang Placeholder
Private Sub tglawalTX_GotFocus()
If tglawalTX.Text = "Tgl Awal" Then
tglawalTX.Text = ""
tglawalTX.ForeColor = &H0
End If
End Sub
Private Sub tglawalTX_LostFocus()
If tglawalTX.Text = "" Then
tglawalTX.Text = "Tgl Awal"
tglawalTX.ForeColor = &H808080
End If
End Sub


Private Sub tglakhirTX_GotFocus()
If tglakhirTX.Text = "Tgl Akhir" Then
tglakhirTX.Text = ""
tglakhirTX.ForeColor = &H0
End If
End Sub
Private Sub tglakhirTX_LostFocus()
If tglakhirTX.Text = "" Then
tglakhirTX.Text = "Tgl Akhir"
tglakhirTX.ForeColor = &H808080
End If
End Sub

