VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEltoInmo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Elementos de inmovilizado"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   Icon            =   "frmEltoInmo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   20
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   64
      Tag             =   "repartos|N|N|0||sinmov|repartos|||"
      Text            =   "commor"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmEltoInmo.frx":000C
      Left            =   7200
      List            =   "frmEltoInmo.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Situación|N|N|1||sinmov|situacio|||"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   19
      Left            =   6840
      MaxLength       =   30
      TabIndex        =   21
      Tag             =   "Fecha Venta/baja|F|S|||sinmov|fecventa|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   18
      Left            =   5040
      MaxLength       =   30
      TabIndex        =   20
      Tag             =   "Años vida|N|N|0||sinmov|anovidas|||"
      Text            =   "commor"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   17
      Left            =   6840
      MaxLength       =   30
      TabIndex        =   19
      Tag             =   "Años máximo|N|N|0||sinmov|anomaxim|||"
      Text            =   "commor"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   16
      Left            =   5040
      MaxLength       =   30
      TabIndex        =   18
      Tag             =   "Años minimos|N|N|0||sinmov|anominim|||"
      Text            =   "commor"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   15
      Left            =   6840
      MaxLength       =   30
      TabIndex        =   17
      Tag             =   "Valor venta/baja|N|S|||sinmov|impventa|#,###,##0.00||"
      Text            =   "commor"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   14
      Left            =   5040
      MaxLength       =   30
      TabIndex        =   16
      Tag             =   "Valor residual|N|N|0||sinmov|valorres|#,###,##0.00||"
      Text            =   "commor"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   13
      Left            =   6840
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "Amortización acumulada|N|N|0||sinmov|amortacu|#,###,##0.00||"
      Text            =   "commor"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   12
      Left            =   5040
      MaxLength       =   30
      TabIndex        =   14
      Tag             =   "Valor adquisición|N|N|0||sinmov|valoradq|#,###,##0.00||"
      Text            =   "commor"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   120
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "Cta de amortizacion|N|N|||sinmov|codmact3|||"
      Text            =   "commor"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   3
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "Text4"
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   120
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "Cta de gastos|N|N|||sinmov|codmact2|||"
      Text            =   "commor"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "Text4"
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   120
      MaxLength       =   30
      TabIndex        =   11
      Tag             =   "Cta inmovilizado|N|N|||sinmov|codmact1|||"
      Text            =   "commor"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "Text4"
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "Porcentaje|N|N|0|100|sinmov|coeficie|0.00||"
      Text            =   "commor"
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmEltoInmo.frx":0048
      Left            =   5760
      List            =   "frmEltoInmo.frx":0058
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Situación|N|N|1||sinmov|tipoamor|||"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   120
      MaxLength       =   30
      TabIndex        =   9
      Tag             =   "Centro de coste|T|S|||sinmov|codccost|||"
      Text            =   "commor"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "Text3"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Text4"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   2760
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Concepto|N|N|0||sinmov|conconam|||"
      Text            =   "commor"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Num|T|S|||sinmov|numserie|||"
      Text            =   "commor"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   120
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Fecha adquisición|F|N|||sinmov|fechaadq|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   7200
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Fact. proveedor|T|S|||sinmov|factupro|||"
      Text            =   "commor"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Proveedor|N|S|||sinmov|codprove|||"
      Text            =   "commor"
      Top             =   840
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5760
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc2"
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Cod|N|N|0||sinmov|codinmov|000000|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   960
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||sinmov|nominmov|||"
      Text            =   "commor"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Text4"
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   33
      Top             =   6240
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   6960
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   60
      TabIndex        =   22
      Top             =   6240
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
      TabIndex        =   32
      Top             =   6240
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   3420
      MaxLength       =   4
      TabIndex        =   23
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   3
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   24
      Top             =   6240
      Width           =   885
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   6000
      MaxLength       =   5
      TabIndex        =   25
      Top             =   6240
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6960
      Top             =   0
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
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   28
      Top             =   6720
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      Top             =   6960
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEltoInmo.frx":0085
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   5040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2778
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
      TabIndex        =   34
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
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
            Object.Visible         =   0   'False
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
         Left            =   7440
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7440
      TabIndex        =   30
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   6600
      Picture         =   "frmEltoInmo.frx":009A
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   7920
      Picture         =   "frmEltoInmo.frx":0A9C
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmEltoInmo.frx":0B27
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image imgConcep 
      Height          =   240
      Left            =   3600
      Picture         =   "frmEltoInmo.frx":0BB2
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image imgCC 
      Height          =   240
      Left            =   1320
      Picture         =   "frmEltoInmo.frx":15B4
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   3
      Left            =   2280
      Picture         =   "frmEltoInmo.frx":1FB6
      Top             =   3600
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   2
      Left            =   1200
      Picture         =   "frmEltoInmo.frx":29B8
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   1
      Left            =   1560
      Picture         =   "frmEltoInmo.frx":33BA
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   0
      Left            =   4320
      Picture         =   "frmEltoInmo.frx":3DBC
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Situación"
      Height          =   255
      Index           =   21
      Left            =   7200
      TabIndex        =   63
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo amort."
      Height          =   255
      Index           =   20
      Left            =   5760
      TabIndex        =   62
      Top             =   1320
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   8520
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   4680
      X2              =   4680
      Y1              =   2160
      Y2              =   4920
   End
   Begin VB.Label Label1 
      Caption         =   "F. Venta/Baja"
      Height          =   195
      Index           =   19
      Left            =   6840
      TabIndex        =   61
      Top             =   4320
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Años vida"
      Height          =   255
      Index           =   18
      Left            =   5040
      TabIndex        =   60
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Años máximos"
      Height          =   255
      Index           =   17
      Left            =   6840
      TabIndex        =   59
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Años mínimos"
      Height          =   255
      Index           =   16
      Left            =   5040
      TabIndex        =   58
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Valor Venta/Baja"
      Height          =   255
      Index           =   15
      Left            =   6840
      TabIndex        =   57
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Valor residual"
      Height          =   255
      Index           =   14
      Left            =   5040
      TabIndex        =   56
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Amort. acumulad"
      Height          =   255
      Index           =   13
      Left            =   6840
      TabIndex        =   55
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Valor adqusicion"
      Height          =   255
      Index           =   12
      Left            =   5040
      TabIndex        =   54
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta Inmovilizado"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   53
      Top             =   2880
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta. amortizacion acumulada"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   51
      Top             =   3600
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta gastos"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   49
      Top             =   4320
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Porcentaje"
      Height          =   195
      Index           =   7
      Left            =   3480
      TabIndex        =   47
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Centro de coste"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   46
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   43
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Serie"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   42
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fec. Adq"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   41
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Factura"
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   40
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   39
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   38
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Elemento"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   37
      Top             =   600
      Width           =   855
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmEltoInmo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Nuevo As String                      'Nuevo desde el form de facturas
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private Const NO = "No encontrado"
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmCI As frmConceptosInmo
Attribute frmCI.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

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
'6 Modo momentaneo para poder poner los campos


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


'Para pasar de lineas a cabeceras
Dim Linliapu As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar


Dim TotalLin As Currency
Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean
Dim RC As String

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    Dim Limp As Boolean

    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
                If InsertarDesdeForm(Me) Then
                    espera 0.5
                    
                    
                   'Si nuevo <>"" siginifac k veni,mos de facturas proveedores. Les vamos a devolver
                   'el valor de la cta de amortizacion
                   'y nos salimos
                   If Nuevo <> "" Then
                        CadenaDesdeOtroForm = Text1(9).Text & "|" & Text4(1).Text & "|"
                        'Para que pueda salir
                        Modo = 10
                        PulsadoSalir = True
                        Unload Me
                        Exit Sub
                   End If
                    
                   'Ponemos la cadena consulta
                    If SituarData1 Then
                        lblIndicador.Caption = ""
                        'Ahora preguntamos si tiene centros de desea agregar centros de reparto
                        SQL = "¿Desea agregar sub centros de reparto?"
                        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                            PonerModo 5
                            'Haremos como si pulsamo el boton de insertar nuevas lineas
                            cmdCancelar.Caption = "Cabecera"
                            ModificandoLineas = 0
                            AnyadirLinea True
                        Else
                            PonerModo 2
                        End If
                    Else
                        SQL = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                        SQL = SQL & vbCrLf & " CLAVE: FRMiNMOV. cmdAceptar. SituarData1"
                        MsgBox SQL, vbCritical
                        Exit Sub
                    End If
                End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1 Then
                        lblIndicador.Caption = ""
                        PonerModo 2
                    End If
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
                CargaGrid True
                Limp = True
                ModificandoLineas = 0
                AnyadirLinea True
                Else
                    ModificandoLineas = 0
                    CamposAux False, 0, False
                    cmdCancelar.Caption = "Cabecera"
                End If
        End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click(Index As Integer)
    cmdAux(0).Tag = 100
   LLamaCuenta
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
    Case 5
        CamposAux False, 0, False
        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
                 DataGrid1.Refresh
            End If
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
Private Function SituarData1() As Boolean
    Dim SQL As String
    
    On Error GoTo ESituarData1
    Data1.Refresh
    espera 0.2
    Data1.Recordset.Find "codinmov = " & Text1(0).Text
    If Not Data1.Recordset.EOF Then
        SituarData1 = True
        Exit Function
    End If
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
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    Text1(0).Text = ObtenerSigueinteNumeroLinea(True)
    
    'Ponemos otros valores por defecto
    Text1(13).Text = "0"
    Text1(14).Text = "0"
    Text1(4).Text = Format(Now, "dd/mm/yyyy")
    Combo1.ListIndex = 3
    Combo2.ListIndex = 0
    Text1(1).SetFocus
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
        Combo1.ListIndex = -1
        Combo2.ListIndex = -1
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Text1(kCampo).SetFocus
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
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
    
Select Case Index
    Case 0
        If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
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
            
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
End Sub

Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    If Not EliminarDesdeActualizar Then
        '### a mano
        SQL = "INMOVILIZADO." & vbCrLf
        SQL = SQL & "-----------------------------" & vbCrLf & vbCrLf
        SQL = SQL & "Va a eliminar el elemento de inmovilizado:"
        SQL = SQL & vbCrLf & "Cod.         :   " & Data1.Recordset.Fields(0)
        SQL = SQL & vbCrLf & "Descripción    :   " & CStr(Data1.Recordset.Fields(2))
        SQL = SQL & "      ¿Desea continuar ? "
        I = MsgBox(SQL, vbQuestion + vbYesNoCancel)
        'Borramos
        If I <> vbYes Then Exit Sub
        
        SQL = DevuelveDesdeBD("fechainm", "shisin", "codinmov", Data1.Recordset.Fields(0), "N")
        If SQL <> "" Then
            SQL = "Los datos del histórico de inmovilizado del elemento se borrarán también. ¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        'Borro, por si existieran, las lineas
        SQL = "Delete from shisin  WHERE codinmov =" & Data1.Recordset!Codinmov
        Conn.Execute SQL
        
        'Borro el elemento
        SQL = "Delete from sinmov  WHERE codinmov =" & Data1.Recordset!Codinmov
        DataGrid1.Enabled = False
        NumRegElim = Data1.Recordset.AbsolutePosition
        Conn.Execute SQL
        
    End If
    
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid False
        PonerModo 0
        Else
            Data1.Recordset.MoveFirst
            NumRegElim = NumRegElim - 1
            If NumRegElim > 1 Then
                For I = 1 To NumRegElim - 1
                    Data1.Recordset.MoveNext
                Next I
            End If
            PonerCampos
            DataGrid1.Enabled = True
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Elimina Elto."
            Data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    Cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until I = 0
    
    '###a mano
    'Devuelvo tb el estado de elemento
    Cad = Cad & Combo2.ListIndex & "|"
    
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        
        Modo = 0
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codinmov = -1"
        Data1.RecordSource = CadenaConsulta
        Data1.Refresh
        
        
        If Nuevo <> "" Then
            BotonAnyadir
            
            'Ahora de la cadena NUEVO desglosamos los datos
            ' Codprove, nomprove, numfac, cta amort
            Text1(2).Text = RecuperaValor(Nuevo, 1)  'codprove
            Text4(0).Text = RecuperaValor(Nuevo, 2)  'nombre
            Text1(3).Text = RecuperaValor(Nuevo, 3)  'numfac
            Text1(4).Text = RecuperaValor(Nuevo, 4)  'fecha adq
            Text1(12).Text = RecuperaValor(Nuevo, 5)  'importe
            Text1(9).Text = RecuperaValor(Nuevo, 6)  'Cuenta
            Text4(1).Text = RecuperaValor(Nuevo, 7)  'Des. cuenta
            
        Else
        
            'Procedimiento normal
            PonerModo CInt(Modo)
            CargaGrid (Modo = 2)
            If Modo <> 2 Then
                CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
                Data1.RecordSource = CadenaConsulta
            End If
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Set miTag = New CTag
    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
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
    
    'Los campos auxiliares
    CamposAux False, 0, True
    
    
    '## A mano
    NombreTabla = "sinmov"
    Ordenacion = " ORDER BY codinmov"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
'    Data1.UserName = vUsu.Login
'    Data1.password = vUsu.Passwd
'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login
    
    imgCC.Enabled = vParam.autocoste
    Text1(7).Enabled = vParam.autocoste
        
    
    PonerOpcionesMenu
    
    'Maxima longitud cuentas
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(2).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(9).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(10).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(11).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    'CadAncho = False
    cmdRegresar.Visible = DatosADevolverBusqueda <> ""
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
End Sub




Private Sub Form_Unload(Cancel As Integer)
    If Modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Nuevo = ""
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Me.DatosADevolverBusqueda = "" Then Set miTag = Nothing
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
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux

        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas
Select Case cmdAux(0).Tag
Case 100
    'Cuenta normal
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
Case 0, 1, 2, 3
    I = Val(cmdAux(0).Tag)
    Text4(I).Text = RecuperaValor(CadenaSeleccion, 2)
    If I = 0 Then
        I = 2
    Else
        I = I + 8
    End If
    Text1(I).Text = RecuperaValor(CadenaSeleccion, 1)
    
End Select
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCI_DatoSeleccionado(CadenaSeleccion As String)
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
    Text1(8).Text = RecuperaValor(CadenaSeleccion, 3)
    Text1(18).Text = RecuperaValor(CadenaSeleccion, 4)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    If I = 0 Then
        I = 4
    Else
        I = 19
    End If
    Text1(I).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    frmMensajes.Opcion = 3
    frmMensajes.Show vbModal
End Sub

Private Sub imgCC_Click()
    Set frmCC = New frmCCoste
    frmCC.DatosADevolverBusqueda = "0|1|"
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub imgConcep_Click()
    Set frmCI = New frmConceptosInmo
    frmCI.DatosADevolverBusqueda = "0|1|2|3|"
    frmCI.Show vbModal
    Set frmCI = Nothing
End Sub

Private Sub imgcta_Click(Index As Integer)
    cmdAux(0).Tag = Index
   LLamaCuenta
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim F As Date
    
    Set frmF = New frmCal
    F = Now
    If Index = 0 Then
        If Text1(4).Text <> "" Then F = CDate(Text1(4).Text)
    Else
        If Text1(19).Text <> "" Then F = CDate(Text1(19).Text)
    End If
    I = Index
    frmF.Fecha = F
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar False
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    'Condiciones para NO salir
    If Modo = 5 Then Exit Sub
        
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
            PonFoco Text1(Index)
    End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
Dim Im As Currency
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite  '&H80000018
    End If
    
    
    If Modo = 0 Then
        Text1(Index).Text = ""
        Exit Sub
    End If
    
    If Text1(Index).Text = "" Then
        Select Case Index
            Case 2, 9, 10, 11
                    If Index = 2 Then
                        I = 0
                    Else
                        I = (Index - 8)
                    End If
                    Text4(I).Text = ""
                    
            Case 6
                Text2.Text = ""
            Case 7
                Text3.Text = ""
        End Select
        Exit Sub
    End If
    'Si estamos insertando o modificando o buscando
    If Modo >= 3 Then  'Or Modo = 4 Or Modo = 1
    
        Select Case Index
        Case 2, 9, 10, 11
                If Index = 2 Then
                    I = 0
                Else
                    I = (Index - 8)
                End If
                RC = Text1(Index).Text
                If CuentaCorrectaUltimoNivel(RC, SQL) Then
                    Text1(Index).Text = RC
                    Text4(I).Text = SQL
                    RC = ""
                Else
                    If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                        SQL = SQL & " ¿Desea crearla?"
                        If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                            CadenaDesdeOtroForm = RC
                            If Index = 2 Then
                                cmdAux(0).Tag = 0
                            Else
                                cmdAux(0).Tag = Index - 8
                            End If
                            Set frmC = New frmColCtas
                            frmC.DatosADevolverBusqueda = "0|1|"
                            frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                            frmC.Show vbModal
                            Set frmC = Nothing
                            If Text1(Index).Text = RC Then SQL = "" 'Para k no los borre
                        End If
                     Else
                        MsgBox SQL, vbExclamation
                     End If
                        
                        
                    If SQL <> "" Then
                        Text1(Index).Text = ""
                        Text4(I).Text = ""
                        RC = "NO"
                    End If
                End If
                If RC <> "" Then Text1(Index).SetFocus
                
        Case 6
                'Concepto de imovilizado
                If Not IsNumeric(Text1(6).Text) Then
                    MsgBox "Concepto debe ser numérico:", vbExclamation
                    Text1(6).SetFocus
                    Exit Sub
                End If
                SQL = DevuelveDesdeBD("nomconam", "sconam", "codconam", Text1(6).Text, "N")
                If SQL = "" Then
                    MsgBox "Concepto NO encontrado: " & Text1(6).Text, vbExclamation
                    Text1(6).Text = ""
                    Text1(6).SetFocus
                    Exit Sub
                End If
                Text2.Text = SQL
                SQL = "Select * from sconam where codconam =" & Text1(6).Text
                Set miRsAux = New ADODB.Recordset
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                If miRsAux.EOF Then
                    MsgBox "Centro de coste NO encontrado: " & Text1(6).Text, vbExclamation
                    Text1(6).Text = ""
                    Text2.Text = ""
                    Text1(6).SetFocus
                Else
                    Text1(6).Text = miRsAux.Fields(0)
                    Text2.Text = miRsAux.Fields(1)
                    If Modo < 6 Then
                        'Solo insertando y modificando
                        Text1(8).Text = miRsAux.Fields(2)
                        Text1(18).Text = miRsAux.Fields(3)
                    End If
                End If
                miRsAux.Close
                Set miRsAux = Nothing
                
        Case 7
                'Centro de coste
                Text1(7).Text = UCase(Text1(7).Text)
                SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", Text1(7).Text, "T")
                If SQL = "" Then
                    MsgBox "Centro de coste NO encontrado: " & Text1(7).Text, vbExclamation
                    Text1(7).Text = ""
                    Text1(7).SetFocus
                End If
                Text3.Text = SQL

        Case 4, 19
                If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    Text1(Index).SetFocus
                End If
        Case Else
                If Index = 8 Or Index = 12 Or Index = 13 Or Index = 14 Or Index = 15 Then
                    If InStr(1, Text1(Index).Text, ",") Then
                        'El importe ya esta formateado
                        If CadenaCurrency(Text1(Index).Text, Im) Then
                            Text1(Index).Text = Format(Im, FormatoImporte)
                        Else
                            MsgBox "importe incorrecto", vbExclamation
                            Text1(Index).Text = ""
                        End If
                    Else
                        Text1(Index).Text = TransformaPuntosComas(Text1(Index).Text)
                    End If
                End If
                If miTag.Cargar(Text1(Index)) Then
                    If miTag.Comprobar(Text1(Index)) Then
                            If miTag.Formato <> "" Then
                                miTag.DarFormato Text1(Index)
                            End If
                    Else
                        Text1(Index).Text = ""
                        Text1(Index).SetFocus
                    End If
                End If
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
                PonerCadenaBusqueda
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 10)
        Cad = Cad & ParaGrid(Text1(1), 35, "Descripcion")
        Cad = Cad & ParaGrid(Text1(4), 15, "Fecha")
        Cad = Cad & ParaGrid(Text1(12), 15, "Valor adqui.")
        Cad = Cad & ParaGrid(Text1(13), 15, "Amort. acum.")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.VCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|"
            frmB.vTitulo = "Eltos. inmovilizado"
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

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla de elementos de inmovilizado", vbInformation
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
Dim Antmodo As Byte

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True

    'Vamos a poner los valores de los campos referencales
    Antmodo = Modo
    Modo = 6
    Text1_LostFocus (6)
    Text1_LostFocus (2)
    Text1_LostFocus (7)
    Text1_LostFocus (9)
    Text1_LostFocus (10)
    Text1_LostFocus (11)

    Modo = Antmodo
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
        Text1(0).BackColor = &H80000018
    End If
    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nuevo elemento de inmovilizado"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar elemento de inmovilizado"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar elemento de inmovilizado"
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nuevo c. coste reparto"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar c. coste reparto"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar c. coste reparto"
    End If
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B And vUsu.Nivel < 3
    Toolbar1.Buttons(11).Enabled = B
        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    
    
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 3
    Me.mnNuevo.Enabled = Toolbar1.Buttons(6).Enabled
    B = B Or Modo = 1  'o buscar
    cmdAceptar.Visible = B
    For I = 1 To Text1.Count - 1
        Text1(I).Locked = Not B
    Next I
    For I = 0 To 3
        Me.imgCta(I).Enabled = B
    Next I
    Combo1.Enabled = B
    Combo2.Enabled = B
    Me.imgFecha(0).Enabled = B
    Me.imgFecha(1).Enabled = B
    Me.imgCC.Enabled = B And vParam.autocoste
    Me.imgConcep.Enabled = B
    
    
    B = B Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    mnOpcionesAsiPre.Enabled = Not B
   
   
    'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5

    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.Visible = (Modo = 2)
'    Else
'        cmdRegresar.Visible = False
'    End If
    
    '
    Text1(0).Enabled = Modo = 1 Or Modo = 3
    'El text   'Con permisios
    B = ((Modo = 2) Or (Modo = 5)) And vUsu.Nivel < 2
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B

   
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
    
    
    'Ahora , si es superusuario podra modificar el valor del combo de situacion
    'Si no se lo forazremos nosotros
    If Combo2.Enabled Then
        If Modo = 4 Then
            If vUsu.Nivel > 1 Then Combo2.Enabled = False
        End If
    End If
    If Combo2.Enabled Then
        Combo2.BackColor = vbWhite
    Else
        Combo2.BackColor = &H80000018
    End If
End Sub


Private Function DatosOk() As Boolean
    Dim RS As ADODB.Recordset
    Dim B As Boolean
    Dim Im As Currency
    'Comprobamos si tiene lineas
    'Para guardarlo en la BD esta el campo oculto
    If Adodc1.Recordset.EOF Then
        Text1(20).Text = 0
    Else
        Text1(20).Text = 1
    End If
    
    
    
    
    
    
    
    'Si el usuario es administrador entonces el valor del objeto lo tiene
    'k poner a mano
    If vUsu.Nivel > 1 Then   'Usuarios generales
        If Text1(19).Text <> "" Then
            Combo2.ListIndex = 2 'De baja
        Else
            If CCur(Text1(12).Text) <= CCur(Text1(13).Text) Then
                Combo2.ListIndex = 3
            Else
                Combo2.ListIndex = 1
            End If
        End If
    End If
    

    
    B = CompForm(Me)
    
    
    If B Then
        If Text1(15).Text <> "" Then
            Im = ImporteFormateado(Text1(15).Text)
            If Im = 0 Then Text1(15).Text = ""
        End If
        
        If Modo = 4 Then
            SQL = ""
            If IsNull(Data1.Recordset!fecventa) Then
                If Text1(19).Text <> "" Then SQL = "No deberia ponerle fecha de baja. El proceso de dar de baja esta en otro punto de la aplicación"
            Else
                If Text1(19).Text = "" Then SQL = "No deberia quitarle la fecha de baja al elemento."
            End If
            If SQL <> "" Then
                SQL = SQL & vbCrLf & "¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then B = False
                SQL = ""
            End If
        End If
    End If
    DatosOk = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
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
   
        'Nuevo Modo
        PonerModo 5
        'Fuerzo que se vean las lineas
        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
    Case 11
'''        'ACtualizar asiento
'''        If Data1.Recordset.EOF Then
'''            MsgBox "Ningún asiento para actualizar.", vbExclamation
'''            Exit Sub
'''        End If
'''        If Adodc1 Is Nothing Then Exit Sub
'''        If Adodc1.Recordset.EOF Then
'''            MsgBox "No hay lineas insertadas para este asiento", vbExclamation
'''            Exit Sub
'''        End If
        

    Case 13
        'Imprimir asientos ''Esto esa mal

            Screen.MousePointer = vbHourglass
            SQL = ""
            If Not Data1.Recordset Is Nothing Then
                If Not Data1.Recordset.EOF Then SQL = Data1.RecordSource
            End If
            If ImpirmirListadoInmovilizados(SQL) Then
                    frmImprimir.Opcion = 64
                    frmImprimir.NumeroParametros = 0
                    frmImprimir.OtrosParametros = ""
                    frmImprimir.FormulaSeleccion = "{ado.codusu} = " & vUsu.Codigo
                    frmImprimir.Show vbModal
            End If
            Screen.MousePointer = vbDefault
    Case 14
        mnSalir_Click
    Case 16 To 19
        Desplazamiento (Button.Index - 16)
    Case Else
    
    End Select
End Sub





Private Sub DespalzamientoVisible(bol As Boolean)
    For I = 16 To 19
        Toolbar1.Buttons(I).Enabled = bol
        Toolbar1.Buttons(I).Visible = bol
    Next I
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
    DataGrid1.Columns(3).Width = 2900

    If vParam.autocoste Then
        DataGrid1.Columns(4).Caption = "C.C."
        DataGrid1.Columns(4).Width = 800

        DataGrid1.Columns(5).Caption = "Centro de coste"
        DataGrid1.Columns(5).Width = 2100
    Else
        DataGrid1.Columns(4).Visible = False
        DataGrid1.Columns(5).Visible = False
    End If
    
    DataGrid1.Columns(6).Caption = "% Porce."
    DataGrid1.Columns(6).Width = 900
    DataGrid1.Columns(6).NumberFormat = "0.00"
        

    
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

        
        
    
       
        CadAncho = True
    End If
        
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
    
    DataGrid1.Tag = "Calculando"

    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
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
    SQL = "SELECT sbasin.codinmov, sbasin.numlinea, sbasin.codmacta2, cuentas.nommacta,"
    SQL = SQL & " cabccost.codccost, cabccost.nomccost, sbasin.porcenta"
    SQL = SQL & " FROM (sbasin INNER JOIN cuentas ON sbasin.codmacta2 = cuentas.codmacta) LEFT"
    SQL = SQL & " JOIN cabccost ON sbasin.codccost = cabccost.codccost"
    If Enlaza Then
        SQL = SQL & " WHERE codinmov = " & Data1.Recordset!Codinmov
        Else
        SQL = SQL & " WHERE codinmov = -1"
    End If
    SQL = SQL & " ORDER BY numlinea"
    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea(Limpiar As Boolean)
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
        
    CalculaTotalLineas
    If TotalLin = 0 Then
        MsgBox "No se pueden insertar nuevas lineas." & TotalLin, vbExclamation
        Exit Sub
    End If
    'Obtenemos la siguiente numero de factura
    Linliapu = ObtenerSigueinteNumeroLinea(False)
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
    LLamaLineas anc, 1, Limpiar
    txtaux(4).Text = TotalLin
    'Ponemos el foco
    txtaux(0).SetFocus
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    Linliapu = Adodc1.Recordset!NumLinea
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
    txtaux(0).Text = Adodc1.Recordset.Fields(2)
    txtaux(1).Text = Adodc1.Recordset.Fields!nommacta
    txtaux(2).Text = DataGrid1.Columns(4).Text
    txtaux(3).Text = DataGrid1.Columns(5).Text
    txtaux(4).Text = DataGrid1.Columns(6).Text
    LLamaLineas anc, 2, False
End Sub

Private Sub EliminarLineaFactura()
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Lineas de reparto." & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar la linea: "
    SQL = SQL & Adodc1.Recordset.Fields(3) & " - " & DataGrid1.Columns(5).Text & " " & DataGrid1.Columns(6).Text & "%"
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = "Delete from sbasin"
        SQL = SQL & " WHERE numlinea = " & Adodc1.Recordset!NumLinea
        SQL = SQL & " AND codinmov =" & Adodc1.Recordset!Codinmov
        Conn.Execute SQL
        DataGrid1.Enabled = False
        CargaGrid (Not Data1.Recordset.EOF)
        DataGrid1.Enabled = True
        ActualizaRepartos
    End If
End Sub



Private Function ObtenerSigueinteNumeroLinea(Cabecera As Boolean) As Long
    Dim RS As ADODB.Recordset
    Dim I As Long
    
    Set RS = New ADODB.Recordset
    If Cabecera Then
        SQL = "SELECT Max(codinmov) FROM sinmov"
    Else
        SQL = "Select max(numlinea) from sbasin where codinmov=" & Data1.Recordset!Codinmov
    End If
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

Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim I As Integer
    Dim B As Boolean
    DataGrid1.Enabled = Not Visible
    For I = 0 To 4
        If I = 2 Or I = 3 Then
            B = vParam.autocoste
        Else
            B = True
        End If
        txtaux(I).Visible = Visible And B
        txtaux(I).Top = Altura
    Next I
    

    cmdAux(0).Visible = Visible
    cmdAux(0).Top = Altura
    If Limpiar Then
        For I = 0 To 4
            txtaux(I).Text = ""
        Next I
    End If
    
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
With txtaux(Index)
    If Index <> 5 Then
         .SelStart = 0
        .SelLength = Len(.Text)
    Else
        .SelStart = Len(.Text)
    End If
End With

End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
    Dim Sng As Double
        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtaux(Index).Text = Trim(txtaux(Index).Text)
    
        'Comun a todos
        If txtaux(Index).Text = "" Then
            Select Case Index
            Case 0
                txtaux(1).Text = ""
            Case 2
                txtaux(3).Text = ""
            
       
            End Select
            Exit Sub
        End If
        
        Select Case Index
        Case 0
            'Cta
            
            RC = txtaux(0).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(0).Text = RC
                txtaux(1).Text = SQL
                RC = ""
            Else
                MsgBox SQL, vbExclamation
                txtaux(0).Text = ""
                txtaux(1).Text = ""
                RC = "NO"
            End If
            If RC <> "" Then txtaux(0).SetFocus
            
        Case 2
            'C coste
                txtaux(2).Text = UCase(txtaux(2).Text)
                SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtaux(2).Text, "T")
                If SQL = "" Then
                    MsgBox "Centro de coste NO encontrado: " & txtaux(2).Text, vbExclamation
                    txtaux(2).Text = ""
                    txtaux(2).SetFocus
                End If
                txtaux(3).Text = SQL
        Case 4
                If Not IsNumeric(txtaux(4).Text) Then
                    MsgBox "Porcentaje debe ser numérico", vbExclamation
                    txtaux(4).Text = ""
                    txtaux(4).SetFocus
                End If
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
    
            
    'Porcentaje
    If txtaux(4).Text = "" Then
        AuxOK = "Porcentaje en blanco"
        Exit Function
    End If
   
    If Not IsNumeric(txtaux(4).Text) Then
        AuxOK = "El porcentaje DEBE debe ser numérico"
        Exit Function
    End If
    I = Val(txtaux(4).Text)
    If I < 0 Or I > 100 Then
        AuxOK = "Porcentajes incorrecto"
    End If
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS
        SQL = "INSERT INTO sbasin (codinmov, numlinea, codmacta2, codccost, porcenta) VALUES (" & Data1.Recordset!Codinmov & ","
        SQL = SQL & Linliapu & ",'"
        SQL = SQL & txtaux(0).Text & "',"
        If txtaux(2).Text = "" Then
            SQL = SQL & "NULL"
        Else
            SQL = SQL & "'" & txtaux(2).Text & "'"
        End If
        SQL = SQL & "," & TransformaComasPuntos(txtaux(4).Text) & ")"
        
        
    Else
    
        'MODIFICAR
        'UPDATE linasipre SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
        '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
        SQL = "UPDATE sbasin SET "
        SQL = SQL & " codmacta2 = '" & txtaux(0).Text & "',"
        SQL = SQL & " codccost = "
        If txtaux(2).Text = "" Then
            SQL = SQL & "NULL"
        Else
            SQL = SQL & "'" & txtaux(2).Text & "'"
        End If
        SQL = SQL & ", porcenta = " & TransformaComasPuntos(txtaux(4).Text)
        SQL = SQL & " WHERE sbasin.numlinea = " & Linliapu
        SQL = SQL & " AND sbasin.codinmov =" & Data1.Recordset!Codinmov
        
    End If
    Conn.Execute SQL
    
    'Ahora actualizamos la BD para ver si tiene centro de repartos
    ActualizaRepartos
    InsertarModificar = True
    Exit Function
EInsertarModificar:
        MuestraError Err.Number, "InsertarModificar linea asiento.", Err.Description
End Function
 
Private Sub ActualizaRepartos()
    SQL = "UPDATE sinmov SET Repartos="
    If ModificandoLineas = 1 Or ModificandoLineas = 2 Then
        RC = "1"
    Else
        If Adodc1.Recordset.EOF Then
            RC = "0"
        Else
            RC = "1"
        End If
    End If
    Text1(20).Text = RC
    SQL = SQL & RC & " WHERE codinmov =" & Data1.Recordset!Codinmov
    Conn.Execute SQL
End Sub

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub






Private Sub CargaGrid(Enlaza As Boolean)
Dim B As Boolean
    B = DataGrid1.Enabled
    CargaGrid2 Enlaza
    DataGrid1.Enabled = B
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub LLamaCuenta()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0"
    frmC.ConfigurarBalances = 3  'nuevo
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub CalculaTotalLineas()
    TotalLin = 0
    SQL = "Select Sum(porcenta) from sbasin where codinmov=" & Data1.Recordset!Codinmov
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then _
            TotalLin = miRsAux.Fields(0)
    End If
    miRsAux.Close
    TotalLin = 100 - TotalLin
    Set miRsAux = Nothing
End Sub
