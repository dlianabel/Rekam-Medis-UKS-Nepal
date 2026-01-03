VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form fbackupdata 
   BackColor       =   &H00E09C48&
   Caption         =   "UKS SMK N 1 PEMALANG"
   ClientHeight    =   8250
   ClientLeft      =   4110
   ClientTop       =   1605
   ClientWidth     =   14190
   ControlBox      =   0   'False
   Icon            =   "fbackupdata.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton tbperiksaobtnyaBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "TABEL RIWAYAT PEMERIKSAAN"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CommandButton tbpemeriksaanBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "TABEL PEMERIKSAAN"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton tbpasienBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "TABEL PASIEN"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CommandButton tbstokobtBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "TABEL STOK OBAT"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton tbobatBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "TABEL OBAT"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton tbpetugasBTN 
      BackColor       =   &H00B58A57&
      Caption         =   "TABEL PETUGAS"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid6 
      Bindings        =   "fbackupdata.frx":2C6F1
      Height          =   4815
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
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
      ColumnCount     =   8
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
      BeginProperty Column03 
         DataField       =   "jmlobat"
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
      BeginProperty Column04 
         DataField       =   "suhu"
         Caption         =   "SUHU"
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
         DataField       =   "tensi"
         Caption         =   "TENSI"
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
         DataField       =   "nadi"
         Caption         =   "NADI"
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
         DataField       =   "napas"
         Caption         =   "NAPAS"
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
            ColumnWidth     =   1470,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1470,047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1184,882
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
   Begin VB.TextBox jmldataTX 
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6600
      Width           =   1215
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1320
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   2640
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   3960
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   5280
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   6600
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
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "fbackupdata.frx":2C706
      Height          =   4815
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
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
            ColumnWidth     =   2369,764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049,953
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
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "fbackupdata.frx":2C71B
      Height          =   4815
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
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
            ColumnWidth     =   4094,929
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "fbackupdata.frx":2C730
      Height          =   4815
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
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
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "fbackupdata.frx":2C745
      Height          =   4815
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
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
            ColumnWidth     =   3240
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "fbackupdata.frx":2C75A
      Height          =   4815
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
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
            ColumnWidth     =   5204,977
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1289,764
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menggunakan EXCEL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8280
      TabIndex        =   12
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menggunakan EXCEL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   615
      Left            =   8280
      TabIndex        =   10
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menggunakan EXCEL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00422A0D&
      Height          =   615
      Left            =   8280
      TabIndex        =   9
      Top             =   600
      Width           =   5895
   End
   Begin Project1.PictureG PictureG1 
      Height          =   3345
      Left            =   3000
      Top             =   7320
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "fbackupdata.frx":2C76F
   End
   Begin VB.Shape shapemenu 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00422A0D&
      FillStyle       =   0  'Solid
      Height          =   4335
      Left            =   10200
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Shape shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00422A0D&
      FillStyle       =   0  'Solid
      Height          =   5655
      Left            =   120
      Top             =   1560
      Width           =   9855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Tabel"
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
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Tabel"
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
      Left            =   8280
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
   Begin Project1.PictureG PictureG3 
      Height          =   3345
      Left            =   2760
      Top             =   -2280
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   5900
      GIF             =   "fbackupdata.frx":2CE51
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG6 
      Height          =   11700
      Left            =   240
      Top             =   -3480
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   20638
      GIF             =   "fbackupdata.frx":2D533
      Mirror          =   1
   End
   Begin Project1.PictureG PictureG4 
      Height          =   11760
      Left            =   0
      Top             =   -3360
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   20743
      GIF             =   "fbackupdata.frx":FE46D
      Mirror          =   1
   End
End
Attribute VB_Name = "fbackupdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub kosong()
DataGrid1.Visible = False
DataGrid2.Visible = False
DataGrid3.Visible = False
DataGrid4.Visible = False
DataGrid5.Visible = False
DataGrid6.Visible = False
jmldataTX.Text = ""
End Sub
Sub btnsemula()
tbpetugasBTN.Caption = "TABEL PETUGAS"
tbobatBTN.Caption = "TABEL OBAT"
tbstokobtBTN.Caption = "TABEL STOK OBAT"
tbpasienBTN.Caption = "TABEL PASIEN"
tbpemeriksaanBTN.Caption = "TABEL PEMERIKSAAN"
tbperiksaobtnyaBTN.Caption = "TABEL RIWAYAT PEMERIKSAAN"
End Sub

Private Sub backBTN_Click()
menu.Show
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\dbrekam_medis.mdb;"
Adodc1.Refresh

kosong
DataGrid1.Visible = True
jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
End Sub







Private Sub tbpetugasBTN_Click()
    On Error GoTo ErrorHandler
    
    ' Simpan current directory sebelum CommonDialog dibuka
    Dim currentDir As String
    currentDir = CurDir$

    If tbpetugasBTN.Caption = "TABEL PETUGAS" Then
        kosong
        btnsemula
        DataGrid1.Visible = True
        tbpetugasBTN.Caption = "(BACKUP) TABEL PETUGAS"
    Else
        With Adodc1.Recordset
            respons = MsgBox("Ingin mengekspor tabel ke Excel?", vbOKCancel)
            If respons = vbOK Then
                Dim xlApp As Object
                Dim xlWorkbook As Object
                Dim xlWorksheet As Object
                Dim folderPath As String
                Dim currentDate As String

                ' Buat instance Excel
                Set xlApp = CreateObject("Excel.Application")
                Set xlWorkbook = xlApp.Workbooks.Add
                Set xlWorksheet = xlWorkbook.Sheets(1)

                ' Format tanggal saat ini untuk nama file
                currentDate = Format(Now, "dd-mm-yyyy")
                CommonDialog1.DialogTitle = "Pilih Lokasi untuk Menyimpan"
                CommonDialog1.FileName = "Tabel petugas (Rekam Medis) " & currentDate & ".xlsx"
                CommonDialog1.ShowSave
                folderPath = CommonDialog1.FileName

                ' Jika dialog dibatalkan
                If folderPath = "" Then
                    ' Tutup Excel dengan benar jika pengguna membatalkan
                    xlWorkbook.Close False
                    xlApp.Quit
                    Set xlWorksheet = Nothing
                    Set xlWorkbook = Nothing
                    Set xlApp = Nothing
                    MsgBox "Penyimpanan dibatalkan."
                    Exit Sub
                End If

                ' Tambahkan ekstensi jika pengguna tidak memasukkan ekstensi
                If Right(folderPath, 5) <> ".xlsx" Then
                    folderPath = folderPath & ".xlsx"
                End If

                ' Ekspor data ke Excel
                For i = 1 To .Fields.Count
                    xlWorksheet.Cells(1, i).Value = .Fields(i - 1).Name
                Next i

                .MoveFirst
                Row = 2
                Do While Not .EOF
                    For j = 1 To .Fields.Count
                        xlWorksheet.Cells(Row, j).Value = .Fields(j - 1).Value
                    Next j
                    .MoveNext
                    Row = Row + 1
                Loop

                ' Simpan file Excel
                xlWorkbook.SaveAs folderPath
                xlWorkbook.Close
                xlApp.Quit

                ' Bersihkan objek Excel
                Set xlWorksheet = Nothing
                Set xlWorkbook = Nothing
                Set xlApp = Nothing

                'MsgBox "Tabel berhasil diekspor ke Excel!"
            End If
        End With
        tbpetugasBTN.Caption = "TABEL PETUGAS"
    End If

    ' Kembalikan current directory ke kondisi semula
    ChDrive currentDir
    ChDir currentDir

    jmldataTX.Text = Adodc1.Recordset.RecordCount & " Data"
    Exit Sub

ErrorHandler:
    ' Tutup Excel jika ada error
    If Not xlApp Is Nothing Then
        xlWorkbook.Close False
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
    End If
    MsgBox "Terjadi kesalahan: " & Err.Description
    Resume Next
End Sub







Private Sub tbobatBTN_Click()
    On Error GoTo ErrorHandler
    
    ' Simpan current directory sebelum CommonDialog dibuka
    Dim currentDir As String
    currentDir = CurDir$

    If tbobatBTN.Caption = "TABEL OBAT" Then
        kosong
        btnsemula
        DataGrid2.Visible = True
        tbobatBTN.Caption = "(BACKUP) TABEL OBAT"
    Else
        With Adodc2.Recordset
            respons = MsgBox("Ingin mengekspor tabel ke Excel?", vbOKCancel)
            If respons = vbOK Then
                Dim xlApp As Object
                Dim xlWorkbook As Object
                Dim xlWorksheet As Object
                Dim folderPath As String
                Dim currentDate As String

                ' Buat instance Excel
                Set xlApp = CreateObject("Excel.Application")
                Set xlWorkbook = xlApp.Workbooks.Add
                Set xlWorksheet = xlWorkbook.Sheets(1)

                ' Format tanggal saat ini untuk nama file
                currentDate = Format(Now, "dd-mm-yyyy")
                CommonDialog1.DialogTitle = "Pilih Lokasi untuk Menyimpan"
                CommonDialog1.FileName = "Tabel obat (Rekam Medis) " & currentDate & ".xlsx"
                CommonDialog1.ShowSave
                folderPath = CommonDialog1.FileName

                ' Jika dialog dibatalkan
                If folderPath = "" Then
                    ' Tutup Excel dengan benar jika pengguna membatalkan
                    xlWorkbook.Close False
                    xlApp.Quit
                    Set xlWorksheet = Nothing
                    Set xlWorkbook = Nothing
                    Set xlApp = Nothing
                    MsgBox "Penyimpanan dibatalkan."
                    Exit Sub
                End If

                ' Tambahkan ekstensi jika pengguna tidak memasukkan ekstensi
                If Right(folderPath, 5) <> ".xlsx" Then
                    folderPath = folderPath & ".xlsx"
                End If

                ' Ekspor data ke Excel
                For i = 1 To .Fields.Count
                    xlWorksheet.Cells(1, i).Value = .Fields(i - 1).Name
                Next i

                .MoveFirst
                Row = 2
                Do While Not .EOF
                    For j = 1 To .Fields.Count
                        xlWorksheet.Cells(Row, j).Value = .Fields(j - 1).Value
                    Next j
                    .MoveNext
                    Row = Row + 1
                Loop

                ' Simpan file Excel
                xlWorkbook.SaveAs folderPath
                xlWorkbook.Close
                xlApp.Quit

                ' Bersihkan objek Excel
                Set xlWorksheet = Nothing
                Set xlWorkbook = Nothing
                Set xlApp = Nothing

                'MsgBox "Tabel berhasil diekspor ke Excel!"
            End If
        End With
        tbobatBTN.Caption = "TABEL OBAT"
    End If

    ' Kembalikan current directory ke kondisi semula
    ChDrive currentDir
    ChDir currentDir

    jmldataTX.Text = Adodc2.Recordset.RecordCount & " Data"
    Exit Sub

ErrorHandler:
    ' Tutup Excel jika ada error
    If Not xlApp Is Nothing Then
        xlWorkbook.Close False
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
    End If
    MsgBox "Terjadi kesalahan: " & Err.Description
    Resume Next
End Sub







Private Sub tbstokobtBTN_Click()
    On Error GoTo ErrorHandler
    
    ' Simpan current directory sebelum CommonDialog dibuka
    Dim currentDir As String
    currentDir = CurDir$

    If tbstokobtBTN.Caption = "TABEL STOK OBAT" Then
        kosong
        btnsemula
        DataGrid3.Visible = True
        tbstokobtBTN.Caption = "(BACKUP) TABEL STOK OBAT"
    Else
        With Adodc3.Recordset
            respons = MsgBox("Ingin mengekspor tabel ke Excel?", vbOKCancel)
            If respons = vbOK Then
                Dim xlApp As Object
                Dim xlWorkbook As Object
                Dim xlWorksheet As Object
                Dim folderPath As String
                Dim currentDate As String

                ' Buat instance Excel
                Set xlApp = CreateObject("Excel.Application")
                Set xlWorkbook = xlApp.Workbooks.Add
                Set xlWorksheet = xlWorkbook.Sheets(1)

                ' Format tanggal saat ini untuk nama file
                currentDate = Format(Now, "dd-mm-yyyy")
                CommonDialog1.DialogTitle = "Pilih Lokasi untuk Menyimpan"
                CommonDialog1.FileName = "Tabel stok obat (Rekam Medis) " & currentDate & ".xlsx"
                CommonDialog1.ShowSave
                folderPath = CommonDialog1.FileName

                ' Jika dialog dibatalkan
                If folderPath = "" Then
                    ' Tutup Excel dengan benar jika pengguna membatalkan
                    xlWorkbook.Close False
                    xlApp.Quit
                    Set xlWorksheet = Nothing
                    Set xlWorkbook = Nothing
                    Set xlApp = Nothing
                    MsgBox "Penyimpanan dibatalkan."
                    Exit Sub
                End If

                ' Tambahkan ekstensi jika pengguna tidak memasukkan ekstensi
                If Right(folderPath, 5) <> ".xlsx" Then
                    folderPath = folderPath & ".xlsx"
                End If

                ' Ekspor data ke Excel
                For i = 1 To .Fields.Count
                    xlWorksheet.Cells(1, i).Value = .Fields(i - 1).Name
                Next i

                .MoveFirst
                Row = 2
                Do While Not .EOF
                    For j = 1 To .Fields.Count
                        xlWorksheet.Cells(Row, j).Value = .Fields(j - 1).Value
                    Next j
                    .MoveNext
                    Row = Row + 1
                Loop

                ' Simpan file Excel
                xlWorkbook.SaveAs folderPath
                xlWorkbook.Close
                xlApp.Quit

                ' Bersihkan objek Excel
                Set xlWorksheet = Nothing
                Set xlWorkbook = Nothing
                Set xlApp = Nothing

                'MsgBox "Tabel berhasil diekspor ke Excel!"
            End If
        End With
        tbstokobtBTN.Caption = "TABEL STOK OBAT"
    End If

    ' Kembalikan current directory ke kondisi semula
    ChDrive currentDir
    ChDir currentDir

    jmldataTX.Text = Adodc3.Recordset.RecordCount & " Data"
    Exit Sub

ErrorHandler:
    ' Tutup Excel jika ada error
    If Not xlApp Is Nothing Then
        xlWorkbook.Close False
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
    End If
    MsgBox "Terjadi kesalahan: " & Err.Description
    Resume Next
End Sub







Private Sub tbpasienBTN_Click()
    On Error GoTo ErrorHandler
    
    ' Simpan current directory sebelum CommonDialog dibuka
    Dim currentDir As String
    currentDir = CurDir$

    If tbpasienBTN.Caption = "TABEL PASIEN" Then
        kosong
        btnsemula
        DataGrid4.Visible = True
        tbpasienBTN.Caption = "(BACKUP) TABEL PASIEN"
    Else
        With Adodc4.Recordset
            respons = MsgBox("Ingin mengekspor tabel ke Excel?", vbOKCancel)
            If respons = vbOK Then
                Dim xlApp As Object
                Dim xlWorkbook As Object
                Dim xlWorksheet As Object
                Dim folderPath As String
                Dim currentDate As String

                ' Buat instance Excel
                Set xlApp = CreateObject("Excel.Application")
                Set xlWorkbook = xlApp.Workbooks.Add
                Set xlWorksheet = xlWorkbook.Sheets(1)

                ' Format tanggal saat ini untuk nama file
                currentDate = Format(Now, "dd-mm-yyyy")
                CommonDialog1.DialogTitle = "Pilih Lokasi untuk Menyimpan"
                CommonDialog1.FileName = "Tabel pasien (Rekam Medis) " & currentDate & ".xlsx"
                CommonDialog1.ShowSave
                folderPath = CommonDialog1.FileName

                ' Jika dialog dibatalkan
                If folderPath = "" Then
                    ' Tutup Excel dengan benar jika pengguna membatalkan
                    xlWorkbook.Close False
                    xlApp.Quit
                    Set xlWorksheet = Nothing
                    Set xlWorkbook = Nothing
                    Set xlApp = Nothing
                    MsgBox "Penyimpanan dibatalkan."
                    Exit Sub
                End If

                ' Tambahkan ekstensi jika pengguna tidak memasukkan ekstensi
                If Right(folderPath, 5) <> ".xlsx" Then
                    folderPath = folderPath & ".xlsx"
                End If

                ' Ekspor data ke Excel
                For i = 1 To .Fields.Count
                    xlWorksheet.Cells(1, i).Value = .Fields(i - 1).Name
                Next i

                .MoveFirst
                Row = 2
                Do While Not .EOF
                    For j = 1 To .Fields.Count
                        xlWorksheet.Cells(Row, j).Value = .Fields(j - 1).Value
                    Next j
                    .MoveNext
                    Row = Row + 1
                Loop

                ' Simpan file Excel
                xlWorkbook.SaveAs folderPath
                xlWorkbook.Close
                xlApp.Quit

                ' Bersihkan objek Excel
                Set xlWorksheet = Nothing
                Set xlWorkbook = Nothing
                Set xlApp = Nothing

                'MsgBox "Tabel berhasil diekspor ke Excel!"
            End If
        End With
        tbpasienBTN.Caption = "TABEL PASIEN"
    End If

    ' Kembalikan current directory ke kondisi semula
    ChDrive currentDir
    ChDir currentDir

    jmldataTX.Text = Adodc4.Recordset.RecordCount & " Data"
    Exit Sub

ErrorHandler:
    ' Tutup Excel jika ada error
    If Not xlApp Is Nothing Then
        xlWorkbook.Close False
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
    End If
    MsgBox "Terjadi kesalahan: " & Err.Description
    Resume Next
End Sub







Private Sub tbpemeriksaanBTN_Click()
    On Error GoTo ErrorHandler
    
    ' Simpan current directory sebelum CommonDialog dibuka
    Dim currentDir As String
    currentDir = CurDir$

    If tbpemeriksaanBTN.Caption = "TABEL PEMERIKSAAN" Then
        kosong
        btnsemula
        DataGrid5.Visible = True
        tbpemeriksaanBTN.Caption = "(BACKUP) TABEL PEMERIKSAAN"
    Else
        With Adodc5.Recordset
            respons = MsgBox("Ingin mengekspor tabel ke Excel?", vbOKCancel)
            If respons = vbOK Then
                Dim xlApp As Object
                Dim xlWorkbook As Object
                Dim xlWorksheet As Object
                Dim folderPath As String
                Dim currentDate As String

                ' Buat instance Excel
                Set xlApp = CreateObject("Excel.Application")
                Set xlWorkbook = xlApp.Workbooks.Add
                Set xlWorksheet = xlWorkbook.Sheets(1)

                ' Format tanggal saat ini untuk nama file
                currentDate = Format(Now, "dd-mm-yyyy")
                CommonDialog1.DialogTitle = "Pilih Lokasi untuk Menyimpan"
                CommonDialog1.FileName = "Tabel pemeriksaan (Rekam Medis) " & currentDate & ".xlsx"
                CommonDialog1.ShowSave
                folderPath = CommonDialog1.FileName

                ' Jika dialog dibatalkan
                If folderPath = "" Then
                    ' Tutup Excel dengan benar jika pengguna membatalkan
                    xlWorkbook.Close False
                    xlApp.Quit
                    Set xlWorksheet = Nothing
                    Set xlWorkbook = Nothing
                    Set xlApp = Nothing
                    MsgBox "Penyimpanan dibatalkan."
                    Exit Sub
                End If

                ' Tambahkan ekstensi jika pengguna tidak memasukkan ekstensi
                If Right(folderPath, 5) <> ".xlsx" Then
                    folderPath = folderPath & ".xlsx"
                End If

                ' Ekspor data ke Excel
                For i = 1 To .Fields.Count
                    xlWorksheet.Cells(1, i).Value = .Fields(i - 1).Name
                Next i

                .MoveFirst
                Row = 2
                Do While Not .EOF
                    For j = 1 To .Fields.Count
                        xlWorksheet.Cells(Row, j).Value = .Fields(j - 1).Value
                    Next j
                    .MoveNext
                    Row = Row + 1
                Loop

                ' Simpan file Excel
                xlWorkbook.SaveAs folderPath
                xlWorkbook.Close
                xlApp.Quit

                ' Bersihkan objek Excel
                Set xlWorksheet = Nothing
                Set xlWorkbook = Nothing
                Set xlApp = Nothing

                'MsgBox "Tabel berhasil diekspor ke Excel!"
            End If
        End With
        tbpemeriksaanBTN.Caption = "TABEL PEMERIKSAAN"
    End If

    ' Kembalikan current directory ke kondisi semula
    ChDrive currentDir
    ChDir currentDir

    jmldataTX.Text = Adodc5.Recordset.RecordCount & " Data"
    Exit Sub

ErrorHandler:
    ' Tutup Excel jika ada error
    If Not xlApp Is Nothing Then
        xlWorkbook.Close False
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
    End If
    MsgBox "Terjadi kesalahan: " & Err.Description
    Resume Next
End Sub

Private Sub tbperiksaobtnyaBTN_Click()
    On Error GoTo ErrorHandler
    
    ' Simpan current directory sebelum CommonDialog dibuka
    Dim currentDir As String
    currentDir = CurDir$

    If tbperiksaobtnyaBTN.Caption = "TABEL RIWAYAT PEMERIKSAAN" Then
        kosong
        btnsemula
        DataGrid6.Visible = True
        tbperiksaobtnyaBTN.Caption = "(BACKUP) TABEL RIWAYAT PEMERIKSAAN"
    Else
        With Adodc6.Recordset
            respons = MsgBox("Ingin mengekspor tabel ke Excel?", vbOKCancel)
            If respons = vbOK Then
                Dim xlApp As Object
                Dim xlWorkbook As Object
                Dim xlWorksheet As Object
                Dim folderPath As String
                Dim currentDate As String

                ' Buat instance Excel
                Set xlApp = CreateObject("Excel.Application")
                Set xlWorkbook = xlApp.Workbooks.Add
                Set xlWorksheet = xlWorkbook.Sheets(1)

                ' Format tanggal saat ini untuk nama file
                currentDate = Format(Now, "dd-mm-yyyy")
                CommonDialog1.DialogTitle = "Pilih Lokasi untuk Menyimpan"
                CommonDialog1.FileName = "Tabel riwayat pemeriksaan (Rekam Medis) " & currentDate & ".xlsx"
                CommonDialog1.ShowSave
                folderPath = CommonDialog1.FileName

                ' Jika dialog dibatalkan
                If folderPath = "" Then
                    ' Tutup Excel dengan benar jika pengguna membatalkan
                    xlWorkbook.Close False
                    xlApp.Quit
                    Set xlWorksheet = Nothing
                    Set xlWorkbook = Nothing
                    Set xlApp = Nothing
                    MsgBox "Penyimpanan dibatalkan."
                    Exit Sub
                End If

                ' Tambahkan ekstensi jika pengguna tidak memasukkan ekstensi
                If Right(folderPath, 5) <> ".xlsx" Then
                    folderPath = folderPath & ".xlsx"
                End If

                ' Ekspor data ke Excel
                For i = 1 To .Fields.Count
                    xlWorksheet.Cells(1, i).Value = .Fields(i - 1).Name
                Next i

                .MoveFirst
                Row = 2
                Do While Not .EOF
                    For j = 1 To .Fields.Count
                        xlWorksheet.Cells(Row, j).Value = .Fields(j - 1).Value
                    Next j
                    .MoveNext
                    Row = Row + 1
                Loop

                ' Simpan file Excel
                xlWorkbook.SaveAs folderPath
                xlWorkbook.Close
                xlApp.Quit

                ' Bersihkan objek Excel
                Set xlWorksheet = Nothing
                Set xlWorkbook = Nothing
                Set xlApp = Nothing

                'MsgBox "Tabel berhasil diekspor ke Excel!"
            End If
        End With
        tbperiksaobtnyaBTN.Caption = "TABEL RIWAYAT PEMERIKSAAN"
    End If

    ' Kembalikan current directory ke kondisi semula
    ChDrive currentDir
    ChDir currentDir

    jmldataTX.Text = Adodc6.Recordset.RecordCount & " Data"
    Exit Sub

ErrorHandler:
    ' Tutup Excel jika ada error
    If Not xlApp Is Nothing Then
        xlWorkbook.Close False
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
    End If
    MsgBox "Terjadi kesalahan: " & Err.Description
    Resume Next
End Sub
