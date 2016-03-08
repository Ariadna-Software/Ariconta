VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmAsientos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAsientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaldoHco 
      Height          =   495
      Index           =   0
      Left            =   10560
      Picture         =   "frmAsientos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Saldos en historico"
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSaldoHco 
      Height          =   495
      Index           =   1
      Left            =   11160
      Picture         =   "frmAsientos.frx":6B5C
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Extractos"
      Top             =   6720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Tag             =   "Nº asiento|N|S|0||cabapu|numasien||S|"
      Text            =   "Text1"
      Top             =   810
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   3
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "Obser|T|S|||cabapu|obsdiari|||"
      Text            =   "frmAsientos.frx":D3AE
      Top             =   1320
      Width           =   5775
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "Text4"
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   7380
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Nº asiento predefinido|N|S|0||cabapu|numaspre|||"
      Text            =   "commor"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha entrada|F|N|||cabapu|fechaent|dd/mm/yyyy|S|"
      Text            =   "commor"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "Text4"
      Top             =   840
      Width           =   2835
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   15
      Top             =   6360
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10680
      TabIndex        =   14
      Top             =   7440
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   1
      Left            =   1080
      TabIndex        =   36
      Top             =   6360
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   6
      Top             =   6360
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   7
      Top             =   6360
      Width           =   885
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   8
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   5
      Left            =   6480
      MaxLength       =   30
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   6
      Left            =   8340
      TabIndex        =   10
      Top             =   6360
      Width           =   1125
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   7
      Left            =   9480
      TabIndex        =   11
      Top             =   6360
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   8
      Left            =   10620
      MaxLength       =   4
      TabIndex        =   12
      Top             =   6360
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   795
      Left            =   7440
      TabIndex        =   22
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2940
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   4
         Left            =   2940
         TabIndex        =   28
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "HABER"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   27
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "DEBE"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2640
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   582
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10680
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Tag             =   "numero diario|N|N|0||cabapu|numdiari||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   435
      Left            =   120
      TabIndex        =   16
      Top             =   7320
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   7440
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAsientos.frx":D3B5
      Height          =   4455
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar asiento"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7680
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   495
      Left            =   5400
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Frame framelineas 
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   6480
      Width           =   10275
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   7800
         TabIndex        =   34
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   33
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   420
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   0
         Picture         =   "frmAsientos.frx":D3CA
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   8580
         Picture         =   "frmAsientos.frx":DDCC
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   5160
         Picture         =   "frmAsientos.frx":E7CE
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1980
         Picture         =   "frmAsientos.frx":F1D0
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "C. coste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   7800
         TabIndex        =   35
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   32
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Cta. Contrapartida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame frameextras 
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   6480
      Width           =   10215
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nomctapar"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text3"
         Top             =   420
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nombreconcepto"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   4
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "centrocoste"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   3
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Cta. Contrapartida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   45
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   44
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "C. coste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7800
         TabIndex        =   43
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   2
      Left            =   7920
      Picture         =   "frmAsientos.frx":FBD2
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   1
      Left            =   3960
      Picture         =   "frmAsientos.frx":105D4
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   0
      Left            =   2040
      Picture         =   "frmAsientos.frx":10FD6
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Asiento predefinido"
      Height          =   195
      Index           =   9
      Left            =   8520
      TabIndex        =   52
      Top             =   600
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Asiento"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   51
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   50
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   195
      Index           =   6
      Left            =   7380
      TabIndex        =   48
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   195
      Index           =   5
      Left            =   1500
      TabIndex        =   47
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Diario"
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   20
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Cod Diario"
      Height          =   195
      Index           =   0
      Left            =   3120
      TabIndex        =   18
      Top             =   600
      Width           =   735
   End
   Begin VB.Menu mnOpcionesAsiPre 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "Lineas"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Public ASIENTO As String  'Con pipes numdiari|fechanormal|numasien
Public vLinapu As Integer

Private Const NO = "No encontrado"
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDi As frmTiposDiario
Attribute frmDi.VB_VarHelpID = -1
Private WithEvents frmPre As frmAsiPre
Attribute frmPre.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busquedaa
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'   Nuevo modo --> Modificando lineas
'  5.- Modificando lineas

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private SQL As String
Dim I As Integer
Dim ancho As Integer
'Dim colMes As Integer

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

'-------------------------------------------------------------
Dim AntiguoText1 As String

'Cuando la cuenta lleva contrapartida
Private LlevaContraPartida As Boolean
'Para pasar de lineas a cabeceras
Dim Linliapu As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar


Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean
Dim VieneDeDesactualizar As Boolean
Dim ActualizandoAsiento As Boolean   'Para k no devuelv el contador

Dim PosicionGrid As Integer

Private CadenaAmpliacion As String

Private DiarioPorDefecto As String 'Si solo tiene un diario que lo ponga


Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    Dim Limp As Boolean
    Dim Mc As Contadores
    Dim B As Boolean
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            Set Mc = New Contadores
            I = FechaCorrecta2(CDate(Text1(1).Text))
            If Mc.ConseguirContador("0", (I = 0), False) = 0 Then
                cmdCancelar.Caption = "Cancelar"
                'COMPROBAR NUMERO ASIENTO
                Text1(4).Text = Mc.Contador
                If ComprobarNumeroAsiento((I = 0)) Then
                    B = InsertarDesdeForm(Me)
                Else
                    B = False
                End If
                If B Then
                    Set Mc = Nothing
                    
                    'El LOG
                    'vLog.Insertar 1, vUsu.Codigo, Text1(4).Text & ":" & Text1(1).Text
                    
                    'Ponemos la cadena consulta
                    If SituarData1(True) Then
                        PonerModo 5
                        'Haremos como si pulsamo el boton de insertar nuevas lineas
                        cmdCancelar.Caption = "Cabecera"
                        
                        If Text1(2).Text <> "" Then
                            CopiaLineasAsiento
                            CargaGrid True
                        End If
                        ModificandoLineas = 0
                        AnyadirLinea True
                    Else
                        SQL = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                        SQL = SQL & vbCrLf & " CLAVE: FrmAsientos. cmdAceptar. SituarData1"
                        MsgBox SQL, vbCritical
                        Exit Sub
                    End If
                    
                Else
                    'SI NO INSERTA debemos devolver el contador

                    Mc.DevolverContador "0", (I = 0), Mc.Contador
                End If
            End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos modificar
                'PreparaBloquear
                Limp = Modificar
                'TerminaBloquear
                If Limp Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1(False) Then
                        lblIndicador.Caption = ""
                        PonerModo 2
                    Else
                        PonerModo 0
                    End If
                    DesBloqAsien   'Desbloqueamos el asiento
                    
                    'Ahora, si viene de actualizar
                    If vParam.AsienActAuto Then
                        If Not Adodc1.Recordset.EOF Then
                            If Text2(2).Text = "" Then
                               SQL = "El asiento esta cuadrado. ¿Desea actualizar?"
                                If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                                    If ActualizarASiento Then
                                        If VieneDeDesactualizar Then
                                            lblIndicador.Caption = ""
                                        Else
                                            lblIndicador.Caption = "" ' Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                                        End If
                                        PonerModo 2
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                Else
                    PonerCampos
                End If
            End If
            
    Case 5
        Cad = AuxOK
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
        Else
            'Insertaremos, o modificaremos
            If InsertarModificar Then
                'Reestablecemos los campos
                'y ponemos el grid
                cmdAceptar.Visible = False
                DataGrid1.AllowAddNew = False
                
                If Not Adodc1.Recordset.EOF Then PosicionGrid = DataGrid1.FirstRow
                CargaGrid True
                Limp = True
                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    'Si ha puesto contrapartida borramos
                    If txtaux(3).Text <> "" Then
                        If LlevaContraPartida Then
                            'Ya lleva la contra partida, luego no hacemos na
                            LlevaContraPartida = False
                        Else
                            Cad = "Generar asiento de la contrapartida?"
                            If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                                FijarContraPartida
                                Limp = False
                                LlevaContraPartida = True
                            End If
                        End If
                    Else
                        LlevaContraPartida = False
                    End If
                    txtaux(8).Text = ""
                    Text3(2).Text = ""
                    If Limp Then
                        For I = 0 To 2
                            Text3(I).Text = ""
                        Next I
                        For I = 0 To 7
                            txtaux(I).Text = ""
                        Next I
                    End If
                    ModificandoLineas = 0
                    cmdAceptar.Visible = True
                    cmdCancelar.Caption = "C&abecera"
                    AnyadirLinea False
                    If Limp Then
                        PonerFoco txtaux(0)
                    Else
                        PonerFoco txtaux(2)
                    End If
                Else
                    ModificandoLineas = 0
                    
                    'Intentamos poner el grid donde toca
                    PonerLineaModificadaSeleccionada
                    CamposAux False, 0, False
                    cmdCancelar.Caption = "Cabecera"
                End If
            End If
        End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub PonerLineaModificadaSeleccionada()
    On Error GoTo E1
   ' While Not Adodc1.Recordset.EOF
   '     If CStr(Adodc1.Recordset.Fields(1)) = CStr(Linliapu) Then Exit Sub
   '     Adodc1.Recordset.MoveNext
   ' Wend
   
    
   Adodc1.Recordset.Find "linliapu =" & Linliapu
 
   
   If Adodc1.Recordset.RecordCount - Adodc1.Recordset.AbsolutePosition < DataGrid1.VisibleRows Then
        'Estoy en la utlimo trozo. No habra scroll
   Else
        I = PosicionGrid - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
    End If
    Exit Sub
E1:
    Err.Clear
End Sub



Private Sub cmdAux_Click(Index As Integer)
    cmdAux(0).Tag = 0
    LlamaContraPar
    'txtAux_LostFocus Index
    If txtaux(0).Text <> "" Then PonerFoco txtaux(2)
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3
        LimpiarCampos
        PonerModo 0
    Case 4
        lblIndicador.Caption = ""
        PonerModo 2
        PonerCampos
        DesBloqAsien
    Case 5
        CamposAux False, 0, False
        frameextras.Visible = True
        framelineas.Visible = False
        LlevaContraPartida = False
        If Adodc1.Recordset.EOF Then
              cmdSaldoHco(1).Visible = False
              cmdSaldoHco(0).Visible = False
        End If
        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
            If Adodc1.Recordset.EOF Then
                SQL = "El asiento no tiene lineas. Desea salir igualmente?"
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            Else
                'Si el asiento esta descuadrado hbar que dar una notificacion
                If Text2(2).Text <> "" Then
                    SQL = "El asiento esta descuadrado. Seguro que desea salir de la edición de lineas de asiento ?"
                    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                Else
                    'Si asiento cuadrado y actualizar automaticamente
                    'lanzamos actualizacion
                    If vParam.AsienActAuto Then
                        SQL = "El asiento esta cuadrado. ¿Desea actualizar?"
                        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                            If ActualizarASiento Then
                                If VieneDeDesactualizar Then
                                    PulsadoSalir = True
                                    Unload Me
                                    Exit Sub
                                Else
                                    lblIndicador.Caption = "" ' Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                                End If
                                PonerModo 2
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
            I = DesbloquearAsiento(Text1(4).Text, Text1(0).Text, Format(Text1(1).Text, FormatoFecha))
            If I = 0 Then MsgBox "Error desbloqueando el asiento", vbExclamation
                
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
                 DataGrid1.Refresh
            End If
            frameextras.Visible = Not Adodc1.Recordset.EOF
            cmdAceptar.Visible = False
            cmdCancelar.Caption = "Cabeceras"
            ModificandoLineas = 0
        End If
    End Select
End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1(Insertar As Boolean) As Boolean
    Dim SQL As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        SQL = "Select * from cabapu WHERE numasien =" & Text1(4).Text
        SQL = SQL & " AND fechaent='" & Format(Text1(1).Text, FormatoFecha) & "' AND numdiari = " & Text1(0).Text
        Data1.RecordSource = SQL
    'End If
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!Numasien) = Text1(4).Text Then
                If CStr(.Fields!NumDiari) = Text1(0).Text Then
                    If Format(CStr(.Fields!fechaent), "dd/mm/yyyy") = Text1(1).Text Then
                        SituarData1 = True
                        Exit Function
                    End If
                End If
            End If
            .MoveNext
        Wend
    End With
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    PonerCadenaBusqueda True
    
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    cmdSaldoHco(0).Visible = False
    cmdSaldoHco(1).Visible = False
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    
    '###A mano
    If DiarioPorDefecto <> "" Then
        Text1(0).Text = RecuperaValor(DiarioPorDefecto, 1)
        Text4.Text = RecuperaValor(DiarioPorDefecto, 2)
    End If
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(1)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(4)
        Text1(4).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                PonerFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda False
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
If Data1.Recordset.EOF Then Exit Sub
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
    Case 1
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 3
        Data1.Recordset.MoveLast
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    
    
    
    
    'Comprobamos que la fecha es de ejerccio actual
    If Not AmbitoDeFecha(True) Then Exit Sub
       
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdCancelar.Caption = "Cancelar"
    cmdAceptar.Caption = "&Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano

    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    PonerFoco Text1(0)
End Sub

Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim I As Integer
    Dim Mc As Contadores
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
     'Comprobamos que la fecha es de ejerccio actual
    If Not AmbitoDeFecha(True) Then Exit Sub
       
    
    
    
    If Not EliminarDesdeActualizar Then
        If BloqAsien Then Exit Sub  'Bloqueamos el asiento, para ver si no esta bloqueado por otro
        '### a mano
        SQL = "Cabecera de apuntes." & vbCrLf
        SQL = SQL & "-----------------------------" & vbCrLf & vbCrLf
        SQL = SQL & "Va a eliminar el asiento:"
        SQL = SQL & vbCrLf & "Nº Asiento   :   " & Data1.Recordset.Fields(2)
        SQL = SQL & vbCrLf & "Fecha ent    :   " & CStr(Data1.Recordset.Fields(1))
        SQL = SQL & vbCrLf & "Diario           :   " & Text1(0).Text & " - " & Text4.Text & vbCrLf & vbCrLf
        SQL = SQL & "      ¿Desea continuar ? "
        I = MsgBox(SQL, vbQuestion + vbYesNoCancel)
        'Borramos
        If I <> vbYes Then
            DesBloqAsien
            Exit Sub
        End If
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
        
    End If
    'Devolvemos contador, si no estamos actualizando
    If Not ActualizandoAsiento Then
        I = FechaCorrecta2(CDate(Data1.Recordset.Fields(1)))
        Set Mc = New Contadores
        NumRegElim = Data1.Recordset.Fields(2)
        Mc.DevolverContador "0", I = 0, NumRegElim
        Set Mc = Nothing
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    DataGrid1.Enabled = False
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid False
        PonerModo 0
        Else
            If NumRegElim > Data1.Recordset.RecordCount Then
                Data1.Recordset.MoveLast
            Else
                Data1.Recordset.MoveFirst
                Data1.Recordset.Move NumRegElim - 1
            End If
            PonerCampos
            DataGrid1.Enabled = True
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Not EliminarDesdeActualizar Then
            If Not Data1.Recordset.EOF Then DesBloqAsien
        Else
           If VieneDeDesactualizar Then
                PulsadoSalir = True
                Unload Me
           End If
        End If
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

'If Data1.Recordset.EOF Then
'    MsgBox "Ningún registro devuelto.", vbExclamation
'    Exit Sub
'End If
'
'Cad = ""
'i = 0
'Do
'    j = i + 1
'    i = InStr(j, DatosADevolverBusqueda, "|")
'    If i > 0 Then
'        AUX = Mid(DatosADevolverBusqueda, j, i - j)
'        j = Val(AUX)
'        Cad = Cad & Text1(j).Text & "|"
'    End If
'Loop Until i = 0
'RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub



Private Sub cmdSaldoHco_Click(Index As Integer)
Dim Cta As String
    If Modo = 5 And ModificandoLineas > 0 Then
        If txtaux(0).Text = "" Then
            MsgBox "Seleccione una cuenta", vbExclamation
            Exit Sub
        End If
        SQL = txtaux(0).Text
        Cta = txtaux(1).Text
    Else
        If Adodc1.Recordset.EOF Then
            MsgBox "Ningún registro activo.", vbExclamation
            Exit Sub
        End If
        SQL = Adodc1.Recordset!codmacta
        Cta = DBLet(Adodc1.Recordset!nommacta)
    End If
    If Index = 0 Then
        SaldoHistorico SQL, "", Cta, False
    Else
        If VieneDeDesactualizar Then
            MsgBox "Acaba de desactualizar asientos. No puede hacer consulta desde aqui.", vbExclamation
        Else
            Screen.MousePointer = vbHourglass
            frmConExtr.EjerciciosCerrados = False
            frmConExtr.Cuenta = SQL
            frmConExtr.Show vbModal
        End If
    End If
End Sub



Private Sub Form_Activate()
Dim B As Boolean
  
    If PrimeraVez Then
        B = False
        PrimeraVez = False
        If ASIENTO <> "" Then
            B = True
            Modo = 2
            SQL = "Select * from cabapu "
            SQL = SQL & " WHERE numasien = " & RecuperaValor(ASIENTO, 3)
            SQL = SQL & " AND numdiari =" & RecuperaValor(ASIENTO, 1)
            SQL = SQL & " AND fechaent= '" & Format(RecuperaValor(ASIENTO, 2), FormatoFecha) & "'"
            CadenaConsulta = SQL
            Modo = 2
            PonerCadenaBusqueda False
            'BOTON lineas
            
        Else
            FijarDiarioPorDefecto
            Modo = 0
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE numasien = -1"
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            
            
        End If
        PonerModo CInt(Modo)
        VieneDeDesactualizar = B
        CargaGrid (Modo = 2)
        If Modo <> 2 Then
            
            'ESTO LO HE CAMBIADO HOY 9 FEB 2006
            'Antes no estaba el IF
            If ASIENTO <> "" Then
                'CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
                'Data1.RecordSource = CadenaConsulta
                MsgBox "Proceso de sistema. Stop. Frm_Activate"
            End If
        Else
            'Viene de HCO
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(6).Enabled = False
            DespalzamientoVisible False
        End If
        If ASIENTO <> "" Then
            If vLinapu > 0 Then
                If Not (Adodc1.Recordset Is Nothing) Then
                    If Not Adodc1.Recordset.EOF Then
                        Adodc1.Recordset.Find "linliapu = " & vLinapu
                        If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
                    End If
                End If
            End If
            
            'Pulso botono pasar a lineas
            HacerToolBar 10
        End If
        Toolbar1.Enabled = True
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False
    ActualizandoAsiento = False
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .Enabled = False
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 17
        .Buttons(13).Image = 16
        .Buttons(14).Image = 15
        .Buttons(16).Image = 6
        .Buttons(17).Image = 7
        .Buttons(18).Image = 8
        .Buttons(19).Image = 9
    End With
    
    Caption = "Introducción de apuntes (" & vEmpresa.nomresum & ")"
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
       ' Me.Width = 12000
       ' Me.Height = Screen.Height
    End If
    Me.Height = 8625
    'Los campos auxiliares
    CamposAux False, 0, True
    
    'Si no es analitica no mostramos el label, texto ni IMAGEN
    Text3(2).Visible = vParam.autocoste
    Label2(2).Visible = vParam.autocoste
    Image1(2).Visible = vParam.autocoste
    
    DiarioPorDefecto = ""
    '## A mano
    NombreTabla = "cabapu"
    Ordenacion = " ORDER BY numasien"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
'    Data1.UserName = vUsu.Login
'    Data1.Password = vUsu.Passwd
'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login

    
    'Maxima longitud cuentas
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    txtaux(3).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
'    Data1.CursorType = adOpenDynamic
'    Data1.LockType = adLockPessimistic
    'CadAncho = False
    PulsadoSalir = False
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
End Sub


'Private Sub Form_Resize()
'If Me.WindowState <> 0 Then Exit Sub
'If Me.Width < 11610 Then Me.Width = 11610
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim B As Boolean
    
    If Modo > 2 Then
        B = True
    Else
        B = VieneDeDesactualizar
    End If
    If B Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 3)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
Dim vFe As String

    'Cuentas
    vFe = RecuperaValor(CadenaSeleccion, 3)
    If vFe <> "" Then
        vFe = RecuperaValor(CadenaSeleccion, 1)
        If EstaLaCuentaBloqueada(vFe, CDate(Text1(1).Text)) Then
            MsgBox "Cuenta bloqueada: " & vFe, vbExclamation
            Exit Sub
        End If
    End If
    If cmdAux(0).Tag = 0 Then
        'Cuenta normal
        txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
        txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
        
        'Habilitaremos el ccoste
        HabilitarCentroCoste
        
    Else
        'contrapartida
        txtaux(3).Text = RecuperaValor(CadenaSeleccion, 1)
        Text3(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
'Centro de coste
txtaux(8).Text = RecuperaValor(CadenaSeleccion, 1)
Text3(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
Dim RC As Byte
'Concepto
txtaux(4).Text = RecuperaValor(CadenaSeleccion, 1)
Text3(1).Text = RecuperaValor(CadenaSeleccion, 2)
txtaux(5).Text = RecuperaValor(CadenaSeleccion, 2) & " "
'Habilitamos importes
RC = CByte(Val(RecuperaValor(CadenaSeleccion, 3)))
HabilitarImportes RC
End Sub

Private Sub frmDi_DatoSeleccionado(CadenaSeleccion As String)
Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
Text4.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmPre_DatoSeleccionado(CadenaSeleccion As String)
Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
Text5.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    'Cta contrapartida
    cmdAux(0).Tag = 1
    LlamaContraPar
    PonerFoco txtaux(4)
Case 1
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|"
    frmCon.Show vbModal
    Set frmCon = Nothing
Case 2
    If txtaux(8).Enabled Then
        Set frmCC = New frmCCoste
        frmCC.DatosADevolverBusqueda = "0|1|"
        frmCC.Show vbModal
        Set frmCC = Nothing
    End If
Case 3
    'Como si hubeiran pulsado sobre el cmd +
    cmdAux(0).Tag = 0
    LlamaContraPar
    PonerFoco txtaux(2)
End Select
End Sub

Private Sub imgppal_Click(Index As Integer)
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
    Case 1
        'Tipos diario
        Set frmDi = New frmTiposDiario
        frmDi.DatosADevolverBusqueda = "0"
        frmDi.Show vbModal
        Set frmDi = Nothing
    Case 2
        'ASiento predefinido
        If Modo = 3 Then
            'Solo si es nuevo
            Set frmPre = New frmAsiPre
            frmPre.DatosADevolverBusqueda = "0"
            frmPre.Show vbModal
            Set frmPre = Nothing
        End If
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    'BotonEliminar False
    HacerToolBar 8
End Sub

Private Sub mnLineas_Click()
Dim B As Button
    Set B = Toolbar1.Buttons(10)
    Toolbar1_ButtonClick B
    Set B = Nothing
End Sub

Private Sub mnModificar_Click()
    'BotonModificar
    HacerToolBar 7
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    'Condiciones para NO salir
    If Modo = 5 Then Exit Sub
    If VieneDeDesactualizar Then
        SQL = "Viene de modificar del historico. ¿Desea dejar el asiento en la introducción de apuntes?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    Else
        'Comprobar apuntes sueltos
        '-------------------------------------
        RevisarIntroduccion = RevisarIntroduccion + 1
        If RevisarIntroduccion = 3 Then
            
        Else
            If RevisarIntroduccion > 10 Then RevisarIntroduccion = 0
        End If
    End If
    PulsadoSalir = True
    Screen.MousePointer = vbHourglass
    DataGrid1.Enabled = False
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
        Else
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim RC As Byte
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite  '&H80000018
    End If
    
    'Si estamos insertando o modificando o buscando
    If Modo = 3 Or Modo = 4 Then
        If Text1(Index).Text = "" Then
            If Index = 0 Then
                Text4.Text = ""
            Else
                If Index = 2 Then Text5.Text = ""
            End If
            Exit Sub
        End If
        Select Case Index
        Case 0
            'Tipo diario
            If Not IsNumeric(Text1(0).Text) Then
                MsgBox "Tipo de diario no es numérico: " & Text1(0).Text, vbExclamation
                Text1(0).Text = ""
                Text4.Text = ""
                PonerFoco Text1(0)
                Exit Sub
            End If
             SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(0).Text, "N")
             If SQL = "" Then
                    SQL = "Diario no encontrado: " & Text1(0).Text
                    Text1(0).Text = ""
                    Text4.Text = ""
                    MsgBox SQL, vbExclamation
                    PonerFoco Text1(0)
            End If
            Text1(0).Text = Val(Text1(0))
            Text4.Text = SQL
        Case 1
            SQL = ""
            If Not EsFechaOK(Text1(1)) Then
                MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
                SQL = "mal"
            Else
                RC = FechaCorrecta2(CDate(Text1(1).Text))
                'Text1(1).Text = Format(Text1(1).Text, "dd/mm/yyyy")
                SQL = ""
                If RC > 1 Then
                    If RC = 2 Then
                        SQL = varTxtFec
                    
                    Else
                        If RC = 3 Then
                             SQL = "El ejercicio al que pertenece la fecha: " & Text1(Index).Text & " está cerrado."
                        Else
                            SQL = "Ejercicio para: " & Text1(Index).Text & " todavía no activo"
                        End If
                    End If
                    MsgBox SQL, vbExclamation
                 Else
                    'Fecha correcta. Si tiene valor DiarioPorDefecto entonces NO paso por ese campo
                    'Y me voy directamente al siguiente
                    If DiarioPorDefecto <> "" Then PonerFoco Text1(2)
                 End If
            End If
            If SQL <> "" Then
                Text1(1).Text = ""
                PonerFoco Text1(1)
            End If
            
        Case 2
            SQL = DevuelveDesdeBD("nomaspre", "cabasipre", "numaspre", Text1(2).Text, "N")
            If SQL = "" Then
                Text1(2).Text = "-1"
                SQL = NO
            End If
            Text5.Text = SQL
        End Select
    End If
End Sub

Private Sub HacerBusqueda()
    Dim Cad As String
    Dim CadB As String
    CadB = ObtenerBusqueda(Me)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
        Else
            'Se muestran en el mismo form
            If CadB <> "" Then
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda False
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(4), 20, "Nº Asiento:")
        Cad = Cad & ParaGrid(Text1(1), 30, "Fecha Entrada")
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Diario")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Asientos"
            frmB.vSelElem = 0
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                'If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
               ' Text1(kCampo).SetFocus
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda(Insertando As Boolean)
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Insertando Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla de Apuntes", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
        Else
            PonerModo 2
            'Data1.Recordset.MoveLast
            Data1.Recordset.MoveFirst
            PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
        MuestraError Err.Number, "PonerCadenaBusqueda"
        PonerModo 0
        Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True
    'Cargamos datos extras
    SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(0).Text, "N")
    If SQL = "" Then SQL = "Error en nº de diario"
    Text4.Text = SQL
    
    If Text1(2).Text = "" Then
        SQL = ""
    Else
        SQL = DevuelveDesdeBD("nomaspre", "cabasipre", "numaspre", Text1(2).Text, "N")
        If SQL = "" Then SQL = "Error en nº de asiento predefinido"
    End If
    Text5.Text = SQL
    frameextras.Visible = Not Adodc1.Recordset.EOF
    cmdSaldoHco(0).Visible = Not Adodc1.Recordset.EOF
    cmdSaldoHco(1).Visible = Not Adodc1.Recordset.EOF
    If Modo = 2 Then lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean


    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
'        For i = 0 To Text1.Count - 1
'            Text1(i).BackColor = vbWhite
'            'Text1(0).BackColor = &H80000018
'        Next i
'        'chkVistaPrevia.Visible = False
        'Reestablecemos el color del nuº asien
        Text1(4).BackColor = &HFEF7E4
    End If
    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nuevo apunte diario"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar apunte diario"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar apunte diario"
    End If
    

        
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea apunte diario"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea apunte diario"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea apunte diario"
    End If
    PonerOpcionesMenuGeneral Me
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    frameextras.Visible = B
    If B Then framelineas.Visible = False
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B
    Toolbar1.Buttons(11).Enabled = B
    If Not B Then frameextras.Visible = False
        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    cmdAceptar.Visible = B Or Modo = 1
    'PRueba###
    


    '
    B = B Or (Modo = 5)
    mnOpcionesAsiPre.Enabled = Not B
    B = B And Not VieneDeDesactualizar
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    
   
   
        'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5

    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.Visible = (Modo = 2)
'    Else
'        cmdRegresar.Visible = False
'    End If
    
    '
    Text1(4).Enabled = (Modo = 1)
    Text1(2).Enabled = (Modo = 3 Or Modo = 1) 'Solo insertar
    B = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    Text1(0).Enabled = B
    Text1(1).Enabled = B
    Text1(3).Enabled = B
    'El text
    B = (Modo = 2) Or (Modo = 5)
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B

   If Modo < 2 Then
     Me.cmdSaldoHco(0).Visible = False
     Me.cmdSaldoHco(1).Visible = False
    End If
   
   
    If Modo <= 2 Then
         Me.cmdAceptar.Caption = "Aceptar"
         Me.cmdCancelar.Caption = "Cancelar"
    End If
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    
    B = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = B
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    
    For I = 6 To 11
        If I <> 9 Then Me.Toolbar1.Buttons(I).Enabled = Me.Toolbar1.Buttons(I).Enabled And vUsu.Nivel < 3
    Next I
    
    
    Me.mnNuevo.Enabled = Me.Toolbar1.Buttons(6).Enabled
    Me.mnEliminar.Enabled = Me.Toolbar1.Buttons(7).Enabled
    Me.mnModificar.Enabled = Me.Toolbar1.Buttons(8).Enabled
    Me.mnLineas.Enabled = Me.Toolbar1.Buttons(10).Enabled
    'Si viene de actualizar solo mostramos no dejamos k toque los menus
     If VieneDeDesactualizar Then Me.mnOpcionesAsiPre.Enabled = False
End Sub


Private Function DatosOk() As Boolean
    Dim RS As ADODB.Recordset
    Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function
    
    ' NOV 2007
    ' NUEVA ambitode fecha activa
    '       0 .- Año actual
    '       1 .- Siguiente
    '       2 .- Fuera de ambito  !!! NUEVO !!!
    '       2 .- Anterior al inicio
    '       3 .- Posterior al fin
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            MsgBox "La fecha no pertenece al ejercicio actual ni al siguiente", vbExclamation
        End If
        B = False

    End If
    DatosOk = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub


Private Sub HacerToolBar(Boton As Integer)

    'Si viene desde hco solo podemos MODIFCAR, ELIMINAR, LINEAS, ACTUALIZAR,SALIR
    If VieneDeDesactualizar Then
        I = Boton
        SQL = ""
        If I < 6 Then
            SQL = "NO"
        Else
            If I > 15 Then
                SQL = "NO"
            Else
                'INSERTAR, pero no estamos en edicion lineas
                If I = 6 And Modo <> 5 Then
                    SQL = "NO"
                End If
            End If
        End If
        If SQL <> "" Then
            MsgBox "Esta modificando el asiento de historico. Finalice primero este proceso.", vbExclamation
            Exit Sub
        End If
    End If
    Select Case Boton
    Case 1
        BotonBuscar
    Case 2
        BotonVerTodos
    Case 6
        If Modo <> 5 Then
            BotonAnyadir
        Else
            'AÑADIR linea factura
            AnyadirLinea True
        End If
    Case 7
        If Modo <> 5 Then
            'Intentamos bloquear la cuenta
            If Data1.Recordset Is Nothing Then Exit Sub
            If Data1.Recordset.EOF Then Exit Sub
            If BloqAsien Then Exit Sub
            BotonModificar
        Else
            'MODIFICAR linea factura
            ModificarLinea
        End If
    Case 8
        If Modo <> 5 Then
            BotonEliminar False
        Else
            'ELIMINAR linea factura
            EliminarLineaFactura
        End If
    Case 10
    
    
        'Fechas
        'Comprobamos que la fecha es de ejerccio actual
        If Not AmbitoDeFecha(False) Then Exit Sub
        
        'If RecodsetVacio Then Exit Sub
        If BloqAsien Then Exit Sub
        'Nuevo Modo
        PonerModo 5
        'Fuerzo que se vean las lineas
        frameextras.Visible = True
        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
        
    Case 11
        'ACtualizar asiento
        If Data1.Recordset.EOF Then
            MsgBox "Ningún asiento para actualizar.", vbExclamation
            Exit Sub
        End If
        If Adodc1 Is Nothing Then Exit Sub
        If Adodc1.Recordset.EOF Then
            MsgBox "No hay lineas insertadas para este asiento", vbExclamation
            Exit Sub
        End If
        
        'Comprobamos que la fecha es de ejerccio actual
        If Not AmbitoDeFecha(False) Then Exit Sub
        
        If BloqAsien Then Exit Sub
        ActualizandoAsiento = True
        If ActualizarASiento Then
            'Si viene de HCO salimos
            If VieneDeDesactualizar Then I = 0
        Else
            I = 1
        End If
        ActualizandoAsiento = False
        If I = 0 Then
            PulsadoSalir = True
            Unload Me
            Exit Sub
        End If
    Case 13
        'Imprimir asientos
        Screen.MousePointer = vbHourglass
        frmActualizar.NUmSerie = Text1(4).Text
        frmActualizar.OpcionActualizar = 4
        frmActualizar.Show vbModal
    
    Case 14
        'SALIR
        If Modo < 3 Then mnSalir_Click
    Case 16 To 19
        Desplazamiento (Boton - 16)
    Case Else
    
    End Select
End Sub



Private Function BloqAsien() As Boolean
Dim B As Byte
'Lo bloqueamos
        B = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        BloqAsien = True
        SQL = ""
        I = (BloquearAsiento(Text1(4).Text, Text1(0).Text, Format(Text1(1).Text, FormatoFecha), SQL))
        'Bloqueamos el registro
        If I = 0 Then
            If SQL = "" Then SQL = "Asiento bloqueado"
            MsgBox SQL, vbExclamation
        
        Else
            BloqAsien = False
        End If
        Screen.MousePointer = B
End Function


Private Sub DesBloqAsien()
Dim B As Byte
'Lo bloqueamos
        B = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        I = (DesbloquearAsiento(Text1(4).Text, Text1(0).Text, Format(Text1(1).Text, FormatoFecha)))
        Screen.MousePointer = B
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    For I = 16 To 19
        Toolbar1.Buttons(I).Enabled = bol
        Toolbar1.Buttons(I).Visible = bol
    Next I
End Sub

'--- A mano // control de devoluciones de prismáticos
Private Sub FrmB1_DatoSeleccionado(CadenaSeleccion As String) '-- Proveedores

End Sub


Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = MontaSQLCarga(Enlaza)
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockPessimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    DataGrid1.Tag = "Asignando"
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-numaspre  1.- Lin aspre
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    'Claves lineas asientos predefinidos
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Visible = False

    'Cuenta
    DataGrid1.Columns(2).Caption = "Cuenta"
    DataGrid1.Columns(2).Width = 1005
    
    DataGrid1.Columns(3).Caption = "Denominación"
    DataGrid1.Columns(3).Width = 2395


    DataGrid1.Columns(4).Caption = "Docu."
    DataGrid1.Columns(4).Width = 1005

    DataGrid1.Columns(5).Caption = "Contra."
    DataGrid1.Columns(5).Width = 1005
    
    DataGrid1.Columns(6).Caption = "Cto."
    DataGrid1.Columns(6).Width = 465
    
    DataGrid1.Columns(7).Visible = False
    

        
    DataGrid1.Columns(8).Caption = "Ampliación"
    DataGrid1.Columns(8).Width = 2400

    'Cuenta contrapartida
    DataGrid1.Columns(9).Visible = False
    
    If vParam.autocoste Then
        ancho = 0
    Else
        ancho = 255 'Es la columna del centro de coste divida entre dos
    End If
    
    DataGrid1.Columns(10).Caption = "Debe"
    DataGrid1.Columns(10).NumberFormat = FormatoImporte
    DataGrid1.Columns(10).Width = 1154 + ancho
    DataGrid1.Columns(10).Alignment = dbgRight
            
    DataGrid1.Columns(11).Caption = "Haber"
    DataGrid1.Columns(11).NumberFormat = FormatoImporte
    DataGrid1.Columns(11).Width = 1154 + ancho
    DataGrid1.Columns(11).Alignment = dbgRight
            
            
    If vParam.autocoste Then
        DataGrid1.Columns(12).Caption = "C.C."
        DataGrid1.Columns(12).Width = 510
    Else
        DataGrid1.Columns(12).Visible = False
    End If
    
    DataGrid1.Columns(13).Visible = False
    DataGrid1.Columns(14).Visible = False
    DataGrid1.Columns(15).Visible = False
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        anc = 323
        txtaux(0).Left = DataGrid1.Left + 330
        txtaux(0).Width = DataGrid1.Columns(2).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(3).Left + 90
                
        txtaux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 6
        txtaux(1).Width = DataGrid1.Columns(3).Width - 180
    
        txtaux(2).Left = DataGrid1.Columns(4).Left + 150
        txtaux(2).Width = DataGrid1.Columns(4).Width - 30
    
        txtaux(3).Left = DataGrid1.Columns(5).Left + 150
        txtaux(3).Width = DataGrid1.Columns(5).Width - 45

        
        'Concepto
        txtaux(4).Left = DataGrid1.Columns(6).Left + 150
        txtaux(4).Width = DataGrid1.Columns(6).Width - 45
        
        txtaux(5).Left = DataGrid1.Columns(8).Left + 150
        txtaux(5).Width = DataGrid1.Columns(8).Width - 45
        
        txtaux(6).Left = DataGrid1.Columns(10).Left + 150
        txtaux(6).Width = DataGrid1.Columns(10).Width - 30
        
       
        txtaux(7).Left = DataGrid1.Columns(11).Left + 150
        txtaux(7).Width = DataGrid1.Columns(11).Width - 30
       
        txtaux(8).Left = DataGrid1.Columns(12).Left + 150
        txtaux(8).Width = DataGrid1.Columns(12).Width - 45
       
        CadAncho = True
    End If
        
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
    
    DataGrid1.Tag = "Calculando"
    'Obtenemos las sumas
    ObtenerSumas
    
    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub ObtenerSumas()
    Dim Deb As Currency
    Dim hab As Currency
    Dim RS As ADODB.Recordset
    
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If Adodc1.Recordset Is Nothing Then Exit Sub
    
    If Adodc1.Recordset.EOF Then Exit Sub
    
    
    Set RS = New ADODB.Recordset
    SQL = "SELECT Sum(linapu.timporteD) AS SumaDetimporteD, Sum(linapu.timporteH) AS SumaDetimporteH"
    SQL = SQL & " ,linapu.numdiari,linapu.fechaent,linapu.numasien"
    SQL = SQL & " From linapu GROUP BY linapu.numdiari, linapu.fechaent, linapu.numasien "
    SQL = SQL & " HAVING (((linapu.numdiari)=" & Data1.Recordset!NumDiari
    SQL = SQL & ") AND ((linapu.fechaent)='" & Format(Data1.Recordset!fechaent, FormatoFecha)
    SQL = SQL & "') AND ((linapu.numasien)=" & Data1.Recordset!Numasien
    SQL = SQL & "));"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Deb = 0
    hab = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Deb = RS.Fields(0)
        If Not IsNull(RS.Fields(1)) Then hab = RS.Fields(1)
    End If
    RS.Close
    Set RS = Nothing
    Text2(0).Text = Format(Deb, FormatoImporte): Text2(1).Text = Format(hab, FormatoImporte)
    'Metemos en DEB el total
    Deb = Deb - hab
    If Deb < 0 Then
        Text2(2).ForeColor = vbRed
        Else
        Text2(2).ForeColor = vbBlack
    End If
    If Deb <> 0 Then Text2(2).Text = Format(Deb, FormatoImporte)
End Sub


Private Function MontaSQLCarga(Enlaza As Boolean) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    ' Si ENLAZA -> Enlaza con el data1
    '           -> Si no lo cargamos sin enlazar a nngun campo
    '--------------------------------------------------------------------
    Dim SQL As String
    SQL = "SELECT linapu.numasien, linapu.linliapu, linapu.codmacta, cuentas.nommacta,"
    SQL = SQL & " linapu.numdocum, linapu.ctacontr, linapu.codconce, conceptos.nomconce as nombreconcepto, linapu.ampconce, cuentas_1.nommacta as nomctapar,"
    SQL = SQL & " linapu.timporteD, linapu.timporteH, linapu.codccost, cabccost.nomccost as centrocoste,"
    SQL = SQL & " linapu.numdiari, linapu.fechaent"
    SQL = SQL & " FROM (((linapu LEFT JOIN cuentas AS cuentas_1 ON linapu.ctacontr ="
    SQL = SQL & " cuentas_1.codmacta) LEFT JOIN cabccost ON linapu.codccost = cabccost.codccost)"
    SQL = SQL & " INNER JOIN cuentas ON linapu.codmacta = cuentas.codmacta) INNER JOIN"
    SQL = SQL & " conceptos ON linapu.codconce = conceptos.codconce"
    If Enlaza Then
        SQL = SQL & " WHERE numasien = " & Data1.Recordset!Numasien
        SQL = SQL & " AND numdiari =" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent= '" & Format(Data1.Recordset!fechaent, FormatoFecha) & "'"
        Else
        SQL = SQL & " WHERE numasien = -1"
    End If
    SQL = SQL & " ORDER BY linapu.linliapu"
    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea(Limpiar As Boolean)
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    Linliapu = ObtenerSigueinteNumeroLinea
    'Situamos el grid al final
    
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If
    cmdAceptar.Caption = "Aceptar"
    LLamaLineas anc, 1, Limpiar
    If Limpiar Then HabilitarImportes 0
    'Ponemos el foco
    PonerFoco txtaux(0)
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    Linliapu = Adodc1.Recordset!Linliapu
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    txtaux(0).Text = Adodc1.Recordset.Fields!codmacta
    txtaux(1).Text = Adodc1.Recordset.Fields!nommacta
    txtaux(2).Text = DataGrid1.Columns(4).Text
    txtaux(3).Text = DataGrid1.Columns(5).Text
    txtaux(4).Text = DataGrid1.Columns(6).Text
    txtaux(5).Text = DataGrid1.Columns(8).Text
    Cad = DBLet(Adodc1.Recordset.Fields!timported)
    If Cad <> "" Then
        txtaux(6).Text = Format(Cad, "0.00")
    Else
        txtaux(6).Text = Cad
    End If
    Cad = DBLet(Adodc1.Recordset.Fields!timporteH)
    If Cad <> "" Then
        txtaux(7).Text = Format(Cad, "0.00")
    Else
        txtaux(7).Text = Cad
    End If
    txtaux(8).Text = DBLet(Adodc1.Recordset.Fields!codccost)
    HabilitarImportes 3
    HabilitarCentroCoste
    Text3(0).Text = Text3(5).Text
    Text3(1).Text = Text3(4).Text
    Text3(2).Text = Text3(3).Text
    LLamaLineas anc, 2, False
    PonerFoco txtaux(0)
End Sub

Private Sub EliminarLineaFactura()
Dim P As Integer

    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Lineas de apuntes." & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar la linea: "
    SQL = SQL & Adodc1.Recordset.Fields(3) & " - " & DataGrid1.Columns(10).Text & " " & DataGrid1.Columns(11).Text
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        P = Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from linapu"
        SQL = SQL & " WHERE linapu.linliapu = " & Adodc1.Recordset!Linliapu
        SQL = SQL & " AND linapu.numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND linapu.fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND linapu.numasien=" & Data1.Recordset!Numasien & ";"
        DataGrid1.Enabled = False
        Conn.Execute SQL
        CargaGrid (Not Data1.Recordset.EOF)
        DataGrid1.Enabled = True
        PosicionaLineas P
    End If
End Sub

Private Sub PosicionaLineas(Pos As Integer)
    On Error GoTo EPosicionaLineas
    If Pos > 1 Then
        If Pos > Adodc1.Recordset.RecordCount Then Pos = Adodc1.Recordset.RecordCount - 1
        Adodc1.Recordset.Move Pos
    End If
    
    Exit Sub
EPosicionaLineas:
    Err.Clear
End Sub

Private Function ObtenerSigueinteNumeroLinea() As Long
    Dim RS As ADODB.Recordset
    Dim I As Long
    
    Set RS = New ADODB.Recordset
    SQL = "SELECT Max(linliapu) FROM linapu"
    SQL = SQL & " WHERE linapu.numdiari=" & Data1.Recordset!NumDiari
    SQL = SQL & " AND linapu.fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
    SQL = SQL & "' AND linapu.numasien=" & Data1.Recordset!Numasien & ";"
    RS.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then I = RS.Fields(0)
    End If
    RS.Close
    ObtenerSigueinteNumeroLinea = I + 1
End Function



'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------

Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtaux(0).Text <> "" Then
            hab = HayKHabilitarCentroCoste(txtaux(0).Text)
        Else
            txtaux(8).Text = ""
        End If
        If hab Then
            txtaux(8).BackColor = &H80000005
            Else
            txtaux(8).BackColor = &H80000018
            txtaux(8).Text = ""
        End If
    End If
    txtaux(8).Enabled = hab
    Image1(2).Enabled = hab
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid DataGrid1
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)
    framelineas.Visible = Not B
    frameextras.Visible = B
    'Habilitamos los botones de cuenta
    cmdSaldoHco(1).Visible = Not B
    cmdSaldoHco(0).Visible = Not B
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    frameextras.Visible = Not B
    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim I As Integer
    Dim J As Integer
    
    DataGrid1.Enabled = Not Visible
    If vParam.autocoste Then
        J = txtaux.Count - 1
        Else
        J = txtaux.Count - 2
        txtaux(8).Visible = False
    End If
    For I = 0 To J
        txtaux(I).Visible = Visible
        txtaux(I).Top = Altura
    Next I
        cmdAux(0).Visible = Visible
        cmdAux(0).Top = Altura
    If Limpiar Then
        For I = 0 To J
            txtaux(I).Text = ""
        Next I
        For I = 0 To 3
            Text3(I).Text = ""
        Next I
    End If
    
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
With txtaux(Index)
    AntiguoText1 = .Text
    If Index <> 5 Then
         .SelStart = 0
        .SelLength = Len(.Text)
    Else
        .SelStart = Len(.Text)
    End If
End With

End Sub

Private Sub txtaux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        'Esto sera k hemos pulsado el ENTER
        txtAux_LostFocus Index
        cmdAceptar_Click
    Else
        If KeyCode = 113 Then
            'Esto sera k pedimos la calculadora
            PideCalculadora
        Else
            'Ha pulsado F5. Ponemos linea anterior
            Select Case KeyCode
            Case 116
                PonerLineaAnterior (Index)
                
            Case 117
                'F6
                'Si es el primer campo , y ha pulsado f6
                'cogera la linea de arriba y la pondra en los txtaux
                If Not Adodc1.Recordset Is Nothing Then
                    If Not Adodc1.Recordset.EOF Then
                        Screen.MousePointer = vbHourglass
                        HacerF6
                        Screen.MousePointer = vbDefault
                    End If
                End If
                
            Case Else
                If (Shift And vbCtrlMask) > 0 Then
                    If UCase(Chr(KeyCode)) = "B" Then
                        'OK. Ha pulsado Control + B
                        '----------------------------------------------------
                        '----------------------------------------------------
                        '
                        ' Dependiendo de index lanzaremos una opcion uotra
                        '
                        '----------------------------------------------------
                        
                        'De momento solo para el 5. Cliente
                        Select Case Index
                        Case 4
                            txtaux(4).Text = ""
                            Image1_Click 1
                        Case 8
                            txtaux(8).Text = ""
                            Image1_Click 2
                        End Select
                     End If
                End If
            End Select
        End If
    End If
End Sub


''Desplegaremos su formulario asociado
'Private Function PulsadoMas(Index As Integer, KeyAscii) As Boolean
'    Select Case Index
'    Case 0
'        'Voy a poner la modificacion del "+"
'        'Es que quiere que le mostremos su formulario de regresar
'        txtaux(0).Text = ""
'        cmdAux_Click 0
'        PulsadoMas = True
'    Case 3
'        txtaux(0).Text = ""
'        Image1_Click 0
'        PulsadoMas = True
'    Case 4
'        txtaux(4).Text = ""
'        Image1_Click 1
'        PulsadoMas = True
'
'    End Select
'
'End Function



Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 112 Then
        
        End If
        'If KeyAscii = 43 Then
        '    If PulsadoMas(Index, KeyAscii) Then KeyAscii = 0
        'End If
    End If
End Sub


'1.-Debe    2.-Haber   3.-Decide en asiento
Private Sub HabilitarImportes(tipoConcepto As Byte)
    Dim bDebe As Boolean
    Dim bHaber As Boolean
    
    'Vamos a hacer .LOCKED =
    Select Case tipoConcepto
    Case 1
        bDebe = False
        bHaber = True
    Case 2
        bDebe = True
        bHaber = False
    Case 3
        bDebe = False
        bHaber = False
    Case Else
        bDebe = True
        bHaber = True
    End Select
    
    txtaux(6).Enabled = Not bDebe
    txtaux(7).Enabled = Not bHaber
    
    If bDebe Then
        txtaux(6).BackColor = &H80000018
        Else
        txtaux(6).BackColor = &H80000005
    End If
    If bHaber Then
        txtaux(7).BackColor = &H80000018
        Else
        txtaux(7).BackColor = &H80000005
    End If
End Sub

Private Sub txtaux_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo <> 1 Then
        If KeyCode = 107 Or KeyCode = 187 Then
                KeyCode = 0
                LanzaPantalla Index
        End If
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Importe As Currency
        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtaux(Index).Text = Trim(txtaux(Index).Text)
    
    
        'Comun a todos
        If txtaux(Index).Text = "" Then
            Select Case Index
            Case 0
                HabilitarCentroCoste
                txtaux(1).Text = ""
            Case 3
                Text3(0).Text = ""
            Case 4
                HabilitarImportes 0
            End Select
            Exit Sub
        End If
        
        If txtaux(Index).Text = AntiguoText1 Then
             Exit Sub
        End If
        
        Select Case Index
        Case 0
            RC = txtaux(0).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(0).Text = RC
                If EstaLaCuentaBloqueada(RC, CDate(Text1(1).Text)) Then
                    MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                Else
                    txtaux(1).Text = SQL
                    RC = ""
                End If
            Else
                If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                    'NO EXISTE LA CUENTA
                    SQL = SQL & " ¿Desea crearla?"
                    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                        CadenaDesdeOtroForm = RC
                        cmdAux(0).Tag = Index
                        Set frmC = New frmColCtas
                        frmC.DatosADevolverBusqueda = "0|1|"
                        frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                        frmC.Show vbModal
                        Set frmC = Nothing
                        If txtaux(0).Text = RC Then SQL = "" 'Para k no los borre
                    End If
                Else
                    MsgBox SQL, vbExclamation
                End If
                    
                If SQL <> "" Then
                  txtaux(0).Text = ""
                  txtaux(1).Text = ""
                  RC = "NO"
                End If
            End If
            HabilitarCentroCoste
            If RC <> "" Then PonerFoco txtaux(0)
            
        Case 3
        
            'Contrapartida
        
            RC = txtaux(3).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(3).Text = RC
                Text3(0).Text = SQL
            Else
            
                If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                    'NO EXISTE LA CUENTA
                    SQL = SQL & " ¿Desea crearla?"
                    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                        CadenaDesdeOtroForm = RC
                        cmdAux(0).Tag = Index
                        Set frmC = New frmColCtas
                        frmC.DatosADevolverBusqueda = "0|1|"
                        frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                        frmC.Show vbModal
                        Set frmC = Nothing
                        If txtaux(3).Text = RC Then SQL = "" 'Para k no los borre
                    End If
                Else
                    MsgBox SQL, vbExclamation
                End If
                If SQL <> "" Then
                    txtaux(3).Text = ""
                    Text3(0).Text = ""
                    PonerFoco txtaux(3)
                End If
            End If
            
        Case 4
             If Not IsNumeric(txtaux(4).Text) Then
                    MsgBox "El concepto debe de ser numérico", vbExclamation
                    PonerFoco txtaux(4)
                    Exit Sub
                End If
                
                If Val(txtaux(4).Text) >= 900 Then
                    If vUsu.Nivel > 1 Then
                        MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
                        Text3(1).Text = ""
                        txtaux(4).Text = ""
                        PonerFoco txtaux(4)
                        Exit Sub
                    Else
                        If Me.Tag = "" Then
                            MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
                            Me.Tag = "0"
                        End If
                    End If
                End If
                
                
                
                CadenaAmpliacion = ""
                If Text3(1).Text <> "" Then
                    'Tenia concepto anterior
                    If InStr(1, txtaux(5).Text, Text3(1).Text) > 0 Then CadenaAmpliacion = Trim(Mid(txtaux(5).Text, Len(Text3(1).Text) + 1))
                End If
                
                RC = "tipoconce"
                SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtaux(4).Text, "N", RC)
                If SQL = "" And RC = "tipoconce" Then
                    MsgBox "Concepto NO encontrado: " & txtaux(4).Text, vbExclamation
                    txtaux(4).Text = ""
                    RC = "0"
                End If
                HabilitarImportes CByte(Val(RC))
                Text3(1).Text = SQL
                txtaux(5).Text = SQL
                If txtaux(5).Text <> "" Then txtaux(5).Text = txtaux(5).Text & " "
                txtaux(5).Text = txtaux(5).Text & CadenaAmpliacion
                If RC = "0" Then PonerFoco txtaux(4)
                
        Case 6, 7
                'LOS IMPORTES
                If Not EsNumerico(txtaux(Index).Text) Then
                    MsgBox "Importes deben ser numéricos.", vbExclamation
                    On Error Resume Next
                    txtaux(Index).Text = ""
                    PonerFoco txtaux(Index)
                    Exit Sub
                End If
                
                
                'Es numerico
                SQL = TransformaPuntosComas(txtaux(Index).Text)
                If CadenaCurrency(SQL, Importe) Then
                    txtaux(Index).Text = Format(Importe, "0.00")
                    'Ponemos el otro campo a ""
                    If Index = 6 Then
                        txtaux(7).Text = ""
                    Else
                        txtaux(6).Text = ""
                    End If
                End If
        Case 8
                txtaux(8).Text = UCase(txtaux(8).Text)
                RC = "idsubcos"
                SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtaux(8).Text, "T", RC)
                If SQL = "" Then
                    MsgBox "Concepto NO encontrado: " & txtaux(8).Text, vbExclamation
                    txtaux(8).Text = ""
                End If
                Text3(2).Text = SQL
        End Select
End Sub


Private Function AuxOK() As String
    
    'Cuenta
    If txtaux(0).Text = "" Then
        AuxOK = "Cuenta no puede estar vacia."
        Exit Function
    End If
    
    If Not IsNumeric(txtaux(0).Text) Then
        AuxOK = "Cuenta debe ser numrica"
        Exit Function
    End If
    
    If txtaux(1).Text = NO Then
        AuxOK = "La cuenta debe estar dada de alta en el sistema"
        Exit Function
    End If
    
    If Not EsCuentaUltimoNivel(txtaux(0).Text) Then
        AuxOK = "La cuenta no es de último nivel"
        Exit Function
    End If
    
    
    'Contrapartida
    If txtaux(3).Text <> "" Then
        If Not IsNumeric(txtaux(3).Text) Then
            AuxOK = "Cuenta contrapartida debe ser numérica"
            Exit Function
        End If
        If Text3(0).Text = NO Then
            AuxOK = "La cta. contrapartida no esta dada de alta en el sistema."
            Exit Function
        End If
        If Not EsCuentaUltimoNivel(txtaux(3).Text) Then
            AuxOK = "La cuenta contrapartida no es de último nivel"
            Exit Function
        End If
    End If
        
    'Concepto
    If txtaux(4).Text = "" Then
        AuxOK = "El concepto no puede estar vacio"
        Exit Function
    End If
        
    If txtaux(4).Text <> "" Then
        If Not IsNumeric(txtaux(4).Text) Then
            AuxOK = "El concepto debe de ser numérico."
            Exit Function
        End If
    End If
    
    'Importe
    If txtaux(6).Text <> "" Then
        If Not EsNumerico(txtaux(6).Text) Then
            AuxOK = "El importe DEBE debe ser numérico"
            Exit Function
        End If
    End If
    
    If txtaux(7).Text <> "" Then
        If Not EsNumerico(txtaux(7).Text) Then
            AuxOK = "El importe HABER debe ser numérico"
            Exit Function
        End If
    End If
    
    If Not (txtaux(6).Text = "" Xor txtaux(7).Text = "") Then
        AuxOK = "Solo el debe, o solo el haber, tiene que tener valor"
        Exit Function
    End If
    
    
    'cENTRO DE COSTE
    If txtaux(8).Enabled Then
        If txtaux(8).Text = "" Then
            AuxOK = "Centro de coste no puede ser nulo"
            Exit Function
        End If
    End If
    
                                            'Fecha del asiento
    If EstaLaCuentaBloqueada(txtaux(0).Text, CDate(Text1(1).Text)) Then
        AuxOK = "Cuenta bloqueada: " & txtaux(0).Text
        Exit Function
    End If
    
    'Si lleva contrapartida
    If txtaux(3).Text <> "" Then
        If EstaLaCuentaBloqueada(txtaux(3).Text, CDate(Text1(1).Text)) Then
            AuxOK = "Cuenta contrapartida bloqueada: " & txtaux(0).Text
            Exit Function
        End If
    End If
    AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS
        'INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab) VALUES (1, '2003-01-16', 1, 2, '5720001', 'doc', 1, NULL, 1600, NULL, NULL, NULL, NULL)
        SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
        SQL = SQL & "codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada) VALUES ("
        'Nudiari, fechaentra y numasien es fijo
        SQL = SQL & Data1.Recordset!NumDiari & ",'"
        SQL = SQL & Format(Data1.Recordset!fechaent, FormatoFecha) & "'," & Data1.Recordset!Numasien & ","
        SQL = SQL & Linliapu & ",'"
        SQL = SQL & txtaux(0).Text & "','"
        SQL = SQL & DevNombreSQL(txtaux(2).Text) & "',"
        SQL = SQL & txtaux(4).Text & ",'"
        SQL = SQL & DevNombreSQL(txtaux(5).Text) & "',"
        If txtaux(6).Text = "" Then
          SQL = SQL & ValorNulo & "," & TransformaComasPuntos(txtaux(7).Text) & ","
          Else
          SQL = SQL & TransformaComasPuntos(txtaux(6).Text) & "," & ValorNulo & ","
        End If
        'Centro coste
        If txtaux(8).Text = "" Then
          SQL = SQL & ValorNulo & ","
          Else
          SQL = SQL & "'" & txtaux(8).Text & "',"
        End If
        
        'Contrapartida
        If txtaux(3).Text = "" Then
          SQL = SQL & ValorNulo & ","
          Else
          SQL = SQL & "'" & txtaux(3).Text & "',"
        End If
        'Marca de entrada manual de datos
        SQL = SQL & "'contab',0)"
        
    Else
    
        'MODIFICAR
        'UPDATE linasipre SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
        '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
        SQL = "UPDATE linapu SET "
        
        SQL = SQL & " codmacta = '" & txtaux(0).Text & "',"
        SQL = SQL & " numdocum = '" & DevNombreSQL(txtaux(2).Text) & "',"
        SQL = SQL & " codconce = " & txtaux(4).Text & ","
        SQL = SQL & " ampconce = '" & DevNombreSQL(txtaux(5).Text) & "',"
        If txtaux(6).Text = "" Then
          SQL = SQL & " timporteD = " & ValorNulo & "," & " timporteH = " & TransformaComasPuntos(txtaux(7).Text) & ","
          Else
          SQL = SQL & " timporteD = " & TransformaComasPuntos(txtaux(6).Text) & "," & " timporteH = " & ValorNulo & ","
        End If
        'Centro coste
        If txtaux(8).Text = "" Then
          SQL = SQL & " codccost = " & ValorNulo & ","
          Else
          SQL = SQL & " codccost = '" & txtaux(8).Text & "',"
        End If
        
        If txtaux(3).Text = "" Then
          SQL = SQL & " ctacontr = " & ValorNulo
          Else
          SQL = SQL & " ctacontr = '" & txtaux(3).Text & "'"
        End If
    
        'Sigue punteada
        'SQL = SQL & " ,punteada = 0"

        
        SQL = SQL & " WHERE linapu.linliapu = " & Linliapu
        SQL = SQL & " AND linapu.numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND linapu.fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND linapu.numasien=" & Data1.Recordset!Numasien & ";"
    
    End If
    Conn.Execute SQL
    InsertarModificar = True
    Exit Function
EInsertarModificar:
        MuestraError Err.Number, "InsertarModificar linea asiento.", Err.Description
End Function
 

Private Sub LlamaContraPar()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing
    
End Sub

'Private Sub DeseleccionaGrid()
'    On Error GoTo EDeseleccionaGrid
'
'    While DataGrid1.SelBookmarks.Count > 0
'        DataGrid1.SelBookmarks.Remove 0
'    Wend
'    Exit Sub
'EDeseleccionaGrid:
'        Err.Clear
'End Sub

Private Sub FijarContraPartida()
    Dim Cad As String
    'Hay contrapartida
    'Reasignamos campos de cuentas
    Cad = txtaux(0).Text
    txtaux(0).Text = txtaux(3).Text
    txtaux(3).Text = Cad
    HabilitarCentroCoste
    Cad = txtaux(1).Text
    txtaux(1).Text = Text3(0).Text
    Text3(0).Text = Cad
    
    'Los importes
    HabilitarImportes 3
    Cad = txtaux(6).Text
    txtaux(6).Text = txtaux(7).Text
    txtaux(7).Text = Cad
End Sub





Private Sub CopiaLineasAsiento()
    Dim Concepto As String
    Dim RS As Recordset
    
    'Preguntamos si desea añadir algo a la linea concepto
    SQL = "Si desea añadir texto a la ampliación de concepto escribalo en el cuadro siguiente."
    Concepto = InputBox(SQL, "Ampliación concepto")
    If Concepto <> "" Then Concepto = " " & Trim(Concepto)
    'Utilizaremos los txtaux y la funcion InsertarModificar
    Set RS = New ADODB.Recordset
    SQL = "Select * from linasipre where numaspre= " & Text1(2).Text & " ORDER BY linlapre"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Linliapu = 0
    I = 0
    ModificandoLineas = 1
    While Not RS.EOF
        Linliapu = Linliapu + 1
        txtaux(0).Text = RS!codmacta
        txtaux(2).Text = RS!numdocum
        txtaux(3).Text = DBLet(RS!ctacontr)
        txtaux(4).Text = RS!codconce
        txtaux(5).Text = RS!ampconce & Concepto
        txtaux(6).Text = DBLet(RS!timported)
        txtaux(7).Text = DBLet(RS!timporteH)
        txtaux(8).Text = DBLet(RS!codccost)
        If Not InsertarModificar Then I = I + 1
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargaGrid(Enlaza As Boolean)
Dim B As Boolean
    B = DataGrid1.Enabled
    
    DataGrid1.Enabled = False
    DoEvents
    CargaGrid2 Enlaza
    DoEvents
    DataGrid1.Enabled = B
    
End Sub

Private Function ActualizarASiento() As Boolean
Dim B As Boolean

If Trim(Text2(2).Text) = "" Then
    Screen.MousePointer = vbHourglass
    frmActualizar.OpcionActualizar = 1
    frmActualizar.NumAsiento = CLng(Text1(4).Text)
    frmActualizar.FechaAsiento = CDate(Text1(1).Text)
    frmActualizar.NumDiari = CInt(Text1(0).Text)
    AlgunAsientoActualizado = False
    frmActualizar.Show vbModal
    Me.Refresh
    ActualizandoAsiento = True
    If AlgunAsientoActualizado Then
        Screen.MousePointer = vbHourglass
        If ASIENTO = "" Then
            B = vParam.emitedia
        Else
            B = vParam.listahco
        End If
        'Emite diario al actualizar
        If B Then
            'imprimimos, para ello borramos la tabla auxiliar
            SQL = "Delete FROM Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
            Conn.Execute SQL
            espera 0.2
            SQL = Text1(4).Text & "|" & Format(Text1(1).Text, FormatoFecha) & "|" & Text1(0).Text & "|"
            If IHcoApuntesAlActualizarModificar(SQL) Then
                With frmImprimir
                    SQL = "Actualización del dia " & Format(Now, "dd/mm/yyyy") & " a las " & Format(Now, "hh:mm") & "."
                    .OtrosParametros = "Fechas= """ & SQL & """|Cuenta= """"|FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
                    .NumeroParametros = 3
                    .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                    .SoloImprimir = True
                    'Opcion dependera del combo
                    .Opcion = 12
                    .Show vbModal
                End With
            End If
        End If
        DataGrid1.Enabled = False
        BotonEliminar True
    Else
        'Desbloquear este asiento
        If Not DesbloquearAsiento(Text1(4).Text, Text1(0).Text, Format(Text1(1).Text, FormatoFecha)) Then
            MsgBox "Error desbloqueando el asiento.", vbExclamation
        End If
    End If
    ActualizandoAsiento = False
    ActualizarASiento = True
Else
    MsgBox "Asiento descuadrado.", vbExclamation
    ActualizarASiento = False
    DesBloqAsien
End If
End Function



Private Function Eliminar() As Boolean
On Error GoTo FinEliminar
        Conn.BeginTrans
        SQL = " WHERE  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien & ";"
        
        'Lineas
        Conn.Execute "Delete  from linapu " & SQL
        
        'Cabeceras
        Conn.Execute "Delete  from cabapu " & SQL
        
                
        'El LOG
        vLog.Insertar 3, vUsu, SQL
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        
        
        
        
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Function Modificar() As Boolean
Dim B1 As Boolean
Dim VC As Contadores

    On Error GoTo EModificar
     Modificar = False
     
        '-----------------------------------------------
        ' ABRIL 2006
        '
        ' Si cambia de ejercicio le ofertaremos un nuevo numero de ASIENTO
        '
        B1 = False
        If Data1.Recordset!fechaent <> CDate(Text1(1).Text) Then
            'HAN CAMBIADO DE FECHA
            
            
            SQL = ""
            'Estabamos(pasado) en ejercicio actual
            If Data1.Recordset!fechaent <= vParam.fechafin Then SQL = "A"
                
                
            B1 = False 'Hay que preguntar cambio de contador. De momento NO
            If CDate(Text1(1).Text) <= vParam.fechafin Then
                'La nueva fecha es del actual
                'Si la otra era del siguiente hay que preguntar
                If SQL = "" Then B1 = True
            Else
                If SQL <> "" Then B1 = True
            End If
            
            If B1 Then
                SQL = "Ha cambiado de ejercicios la fecha del asiento." & vbCrLf & " ¿Desea obtener nuevo numero de asiento?"
                SQL = MsgBox(SQL, vbQuestion + vbYesNoCancel)
                If CByte(SQL) = vbCancel Then Exit Function
                
                If CByte(SQL) = vbNo Then B1 = False
                
            End If
        End If
        Set VC = New Contadores
        If B1 Then
            'Obtengo nuevo contador
            If VC.ConseguirContador("0", (CDate(Text1(1).Text) <= vParam.fechafin), False) > 0 Then Exit Function
        Else
            VC.Contador = Data1.Recordset!Numasien
        End If
                    
                    
                    
        Conn.BeginTrans
        'Comun
        
        SQL = " WHERE  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien
        
        'BLoqueamos
        Conn.Execute "Select * from cabapu " & SQL & " FOR UPDATE"
        
        'Añadimos tb el nunmero de asiento
        SQL = " numasien = " & VC.Contador & " , numdiari= " & Text1(0).Text & " , fechaent ='" & Format(Text1(1).Text, FormatoFecha) & "'" & SQL
        
        
       'Las lineas de apuntes
        Conn.Execute "UPDATE linapu SET " & SQL
      
        
        'Modificamos la cabecera
        If Text1(3).Text = "" Then
            SQL = "obsdiari = NULL," & SQL
        Else
            SQL = "Obsdiari ='" & DevNombreSQL(Text1(3).Text) & "'," & SQL
        End If

        Conn.Execute "UPDATE cabapu SET " & SQL
        
  
EModificar:
        If Err.Number <> 0 Then
            MuestraError Err.Number
            Conn.RollbackTrans
            Modificar = False
            B1 = False
        Else
            Conn.CommitTrans
            Modificar = True
        End If
        
        'Si habia que devolver contador
        If B1 Then
            Text1(4).Text = VC.Contador
            Set VC = Nothing
            Set VC = New Contadores
            VC.DevolverContador "0", (Data1.Recordset!fechaent <= vParam.fechafin), Data1.Recordset!Numasien
            
        End If
        Set VC = Nothing
End Function

Private Sub PideCalculadora()
On Error GoTo EPideCalculadora
    Shell App.path & "\arical.exe", vbNormalFocus
    Exit Sub
EPideCalculadora:
    Err.Clear
End Sub


Private Function ComprobarNumeroAsiento(Actual As Boolean) As Boolean
Dim Cad As String
Dim RT As ADODB.Recordset
        Cad = " WHERE numasien=" & Text1(4).Text
        If Actual Then
            I = 0
        Else
            I = 1
        End If
        Cad = Cad & " AND fechaent >='" & Format(DateAdd("yyyy", I, vParam.fechaini), FormatoFecha)
        Cad = Cad & "' AND fechaent <='" & Format(DateAdd("yyyy", I, vParam.fechafin), FormatoFecha) & "'"
        Set RT = New ADODB.Recordset
        ComprobarNumeroAsiento = True
        I = 0
        RT.Open "Select numasien from linapu" & Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.EOF) Then
                ComprobarNumeroAsiento = False
            End If
        End If
        RT.Close
        If ComprobarNumeroAsiento Then
            I = 1
            RT.Open "Select numasien from hlinapu" & Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RT.EOF Then
                If Not IsNull(RT.EOF) Then
                    ComprobarNumeroAsiento = False
                End If
            End If
            RT.Close
        End If
        Set RT = Nothing
        If Not ComprobarNumeroAsiento Then
            Cad = "Verifique los contadores. Ya exsite el asiento; " & Text1(4).Text & vbCrLf
            If I = 0 Then
                Cad = Cad & " en la introducción de apuntes"
            Else
                Cad = Cad & " en el histórico."
            End If
            MsgBox Cad, vbExclamation
        End If
End Function



Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLineaAnterior(Indice As Integer)
Dim RT As ADODB.Recordset
Dim C As String
On Error GoTo EponerLineaAnterior

    'Si no estamos insertando,modificando lineas
    
    If Modo <> 5 Then Exit Sub
    

    If Adodc1.Recordset.EOF Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    
    'Todos los casos menos la ampliacion del concepto
    If Indice <> 5 Then
        SQL = "SELECT "
        Select Case Indice
        Case 0
            C = "codmacta"
            I = 2
        Case 2
            C = "numdocum"
            I = 3
        Case 3
            C = "ctacontr"
            I = 4
        Case 4
            C = "codconce"
            I = 5
        Case 8
            C = "codccost"
            I = -1
        Case Else
            C = ""
        End Select
        If C <> "" Then
            SQL = SQL & C & "  FROM linapu"
            SQL = SQL & " WHERE numdiari=" & Data1.Recordset!NumDiari
            SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
            SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien
            If ModificandoLineas = 2 Then SQL = SQL & " AND linliapu <" & Linliapu
            SQL = SQL & " ORDER BY linliapu DESC"
            Set RT = New ADODB.Recordset
            RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            C = ""
            If Not RT.EOF Then C = DBLet(RT.Fields(0))
            
            'Lo ponemos en txtaux
            If C <> "" Then
                txtaux(Indice).Text = C
                If I >= 0 Then
                    PonerFoco txtaux(I)
                End If
            End If
            RT.Close
        End If





    Else
        SQL = "Select linliapu,ampconce,nomconce FROM linapu,conceptos"
        SQL = SQL & " WHERE conceptos.codconce=linapu.codconce AND  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien
        If ModificandoLineas = 2 Then SQL = SQL & " AND linliapu <" & Linliapu
           
        SQL = SQL & " ORDER BY linliapu DESC"
        Set RT = New ADODB.Recordset
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        C = ""
        If Not RT.EOF Then
            SQL = DBLet(RT.Fields(1))
            C = DBLet(RT.Fields(2))
        End If
        
        'Lo ponemos en txtaux
        If SQL <> "" Then
            If C <> "" Then
                I = InStr(1, SQL, C)
                If I > 0 Then SQL = Trim(Mid(SQL, Len(C) + 1))
            End If
            txtaux(5).Text = txtaux(5).Text & SQL & " "
            txtaux(5).SelStart = Len(txtaux(5).Text)
            PonerFoco txtaux(6)
        End If
        RT.Close

    
    End If
    
EponerLineaAnterior:
    If Err.Number <> 0 Then Err.Clear
    Set RT = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Function RecodsetVacio() As Boolean
    RecodsetVacio = True
    If Not Adodc1.Recordset Is Nothing Then
        If Not Adodc1.Recordset.EOF Then RecodsetVacio = False
    End If
End Function


Private Sub HacerRevisarIntroduccion()
    'VACIO DE MOMENTO

End Sub


Private Sub LanzaPantalla(Index As Integer)
Dim miI As Integer
        '----------------------------------------------------
        '----------------------------------------------------
        '
        ' Dependiendo de index lanzaremos una opcion uotra
        '
        '----------------------------------------------------
        
        'De momento solo para el 5. Cliente
        miI = -1
        Select Case Index
        Case 0
            txtaux(0).Text = ""
            miI = 3
        Case 3
            txtaux(3).Text = ""
            miI = 0
        Case 4
            txtaux(4).Text = ""
            miI = 1
            
        Case 8
            txtaux(8).Text = ""
            miI = 2
        End Select
        If miI >= 0 Then Image1_Click miI
End Sub



Private Sub HacerF6()
Dim RsF6 As ADODB.Recordset
Dim C As String

    On Error GoTo EHacerF6
    
    Set RsF6 = New ADODB.Recordset
            
    
    C = "SELECT linapu.numasien, linapu.linliapu, linapu.codmacta, cuentas.nommacta,"
    C = C & " linapu.numdocum, linapu.ctacontr, linapu.codconce, conceptos.nomconce as nombreconcepto, linapu.ampconce, cuentas_1.nommacta as nomctapar,"
    C = C & " linapu.timporteD, linapu.timporteH, linapu.codccost, cabccost.nomccost as centrocoste,"
    C = C & " linapu.numdiari, linapu.fechaent"
    C = C & " FROM (((linapu LEFT JOIN cuentas AS cuentas_1 ON linapu.ctacontr ="
    C = C & " cuentas_1.codmacta) LEFT JOIN cabccost ON linapu.codccost = cabccost.codccost)"
    C = C & " INNER JOIN cuentas ON linapu.codmacta = cuentas.codmacta) INNER JOIN"
    C = C & " conceptos ON linapu.codconce = conceptos.codconce"
    C = C & " WHERE numasien = " & Data1.Recordset!Numasien
    C = C & " AND numdiari =" & Data1.Recordset!NumDiari
    C = C & " AND fechaent= '" & Format(Data1.Recordset!fechaent, FormatoFecha) & "'"
    C = C & " ORDER BY linapu.linliapu DESC"
    
    
    
    
    
    RsF6.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RsF6.EOF Then
        C = " numasiento = " & Data1.Recordset!Numasien & vbCrLf
        C = " fecha= " & Format(Data1.Recordset!fechaent, "dd/mm/yyyy")
    
        MsgBox "No se ha encontrado las lineas: " & vbCrLf & C, vbExclamation
    Else
        'Ya tengo la ultima linea
        txtaux(0).Text = RsF6!codmacta
        
        txtaux(0).Text = RsF6!codmacta
        txtaux(1).Text = RsF6!nommacta
        txtaux(2).Text = DBLet(RsF6!numdocum, "T")
        txtaux(3).Text = DBLet(RsF6!ctacontr, "T")
        txtaux(4).Text = RsF6!codconce
        txtaux(5).Text = DBLet(RsF6!ampconce, "T")
        C = DBLet(RsF6!timported, "T")
        If C <> "" Then
            txtaux(6).Text = Format(C, "0.00")
        Else
            txtaux(6).Text = C
        End If
        C = DBLet(RsF6!timporteH, "T")
        If C <> "" Then
            txtaux(7).Text = Format(C, "0.00")
        Else
            txtaux(7).Text = C
        End If
        txtaux(8).Text = DBLet(RsF6!codccost, "T")
        HabilitarImportes 3
        HabilitarCentroCoste
        Text3(0).Text = DBLet(RsF6!nomctapar, "T")
        Text3(1).Text = RsF6!nombreconcepto
        Text3(2).Text = DBLet(RsF6!centrocoste, "T")
        
    End If
    RsF6.Close
    Set RsF6 = Nothing
    Exit Sub
EHacerF6:
    MuestraError Err.Number, Err.Description
    Set RsF6 = Nothing
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub



'##### Nuevo para el ambito de fechas
Private Function AmbitoDeFecha(DesbloqueAsiento As Boolean) As Boolean
        AmbitoDeFecha = False
        varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
        If varFecOk > 1 Then
            If varFecOk = 2 Then
                MsgBox varTxtFec, vbExclamation
            Else
                MsgBox "El asiento pertenece a un ejercicio cerrado.", vbExclamation
            End If
        Else
            AmbitoDeFecha = True
        End If
    
        If DesbloqueAsiento Then DesBloqAsien
End Function

Private Sub FijarDiarioPorDefecto()
    
    AntiguoText1 = "Select * from tiposdiario"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open AntiguoText1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        DiarioPorDefecto = miRsAux!NumDiari & "|" & miRsAux!desdiari & "|"
        miRsAux.MoveNext
        If Not miRsAux.EOF Then AntiguoText1 = ""
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    'Si hay mas de un diario, NO hago nada
    If AntiguoText1 = "" Then DiarioPorDefecto = ""
    AntiguoText1 = ""
        
End Sub
