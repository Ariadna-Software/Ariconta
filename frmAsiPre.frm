VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmAsiPre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de asientos predefinidos"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAsiPre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11910
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
      Left            =   10560
      TabIndex        =   11
      Top             =   7200
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   60
      TabIndex        =   2
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
      MaxLength       =   10
      TabIndex        =   3
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   4560
      TabIndex        =   4
      Top             =   6240
      Width           =   885
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   5
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   5
      Left            =   6240
      MaxLength       =   30
      TabIndex        =   6
      Top             =   6240
      Width           =   2070
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
      TabIndex        =   7
      Top             =   6240
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
      TabIndex        =   8
      Top             =   6240
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
      TabIndex        =   9
      Top             =   6240
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   795
      Left            =   7440
      TabIndex        =   18
      Top             =   480
      Width           =   4335
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2940
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   4
         Left            =   2940
         TabIndex        =   24
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "HABER"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   23
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "DEBE"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   22
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   3540
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
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre asiento predefinido|T|N|||cabasipre|nomaspre|||"
      Text            =   "commor"
      Top             =   840
      Width           =   4155
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10560
      TabIndex        =   15
      Top             =   7200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "Nº asiento predefinido|N|N|||cabasipre|numaspre|0000|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   7200
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   0
      Top             =   3060
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAsiPre.frx":030A
      Height          =   4815
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   8493
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
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Lineas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8280
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame framelineas 
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   6240
      Width           =   11715
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   7800
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   29
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
         TabIndex        =   27
         Text            =   "Text3"
         Top             =   420
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   10920
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   8580
         Picture         =   "frmAsiPre.frx":031F
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   5160
         Picture         =   "frmAsiPre.frx":0D21
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1980
         Picture         =   "frmAsiPre.frx":1723
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
         TabIndex        =   31
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
         TabIndex        =   28
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
         TabIndex        =   26
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame frameextras 
      Height          =   855
      Left            =   120
      TabIndex        =   36
      Top             =   6240
      Width           =   11715
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nomctapar"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   16
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Num:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   630
      Width           =   1215
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
Attribute VB_Name = "frmAsiPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private Const NO = "No encontrado"
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1

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


'Cuando la cuenta lleva contrapartida
Private LlevaContraPartida As Boolean
'Para pasar de lineas a cabeceras
Dim NumLin As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar








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
                If SituarData1 Then
                    PonerModo 5
                    'Haremos como si pulsamo el boton de insertar nuevas lineas
                    cmdCancelar.Caption = "Cabecera"
                    AnyadirLinea True
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
                    If SituarData1 Then PonerModo 2
                    'lblIndicador.Caption = "Insertado"
                    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
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
                CargaGrid Data1.Recordset!numaspre
                Limp = True
                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    'Si ha puesto contrapartida borramos
                    If txtaux(3).Text <> "" Then
                        If LlevaContraPartida Then
                            'Ya lleva la contra partida, luego no hacemos na
                            LlevaContraPartida = False
                        Else
                            FijarContraPartida
                            Limp = False
                            LlevaContraPartida = True
                        End If
                    Else
                        LlevaContraPartida = False
                    End If
                    txtaux(8).Text = ""
                    Text3(2).Text = ""
                    If Limp Then
                        For I = 0 To 7
                            txtaux(I).Text = ""
                        Next I
                        Text3(0).Text = ""
                        Text3(1).Text = ""
                    End If
                    ModificandoLineas = 0
                    cmdAceptar.Visible = True
                    AnyadirLinea False
                    If Limp Then
                        txtaux(0).SetFocus
                    Else
                        txtaux(2).SetFocus
                    End If
                Else
                    ModificandoLineas = 0
                    CamposAux False, 0, False
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

Private Sub cmdAux_Click(Index As Integer)
    cmdAux(0).Tag = 0
    LlamaContraPar
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1, 3
    LimpiarCampos
    PonerModo 0
Case 4
    PonerModo 2
    PonerCampos
Case 5
    CamposAux False, 0, False
    LlevaContraPartida = False
    'Si esta insertando/modificando lineas haremos unas cosas u otras
    DataGrid1.Enabled = True
    If ModificandoLineas = 0 Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        PonerModo 2
    Else
        If ModificandoLineas = 1 Then
             DataGrid1.AllowAddNew = False
             If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
             DataGrid1.Refresh
        End If
        frameextras.Visible = True
        framelineas.Visible = False
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
            'Actualizamos el recordset
            CadenaConsulta = "Select * from " & NombreTabla
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = " numaspre = " & Val(Text1(0).Text)
            Data1.Recordset.Find SQL
            If Data1.Recordset.EOF Then GoTo ESituarData1
            SituarData1 = True
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid -1
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    SugerirCodigoSiguiente
    '###A mano
    Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid -1
        
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                'Fallo . No poner el foco directamente.
                'Text1(kCampo).SetFocus
                PonleFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid -1
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
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
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    
    
    
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim I As Integer
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    '### a mano
    Cad = "Seguro que desea eliminar de la BD el asiento predefinido:"
    Cad = Cad & vbCrLf & "Nº Asiento: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Descrpcion: " & Data1.Recordset.Fields(1)
    I = MsgBox(Cad, vbQuestion + vbYesNo)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        'FALTA###
        'Habra que BEGIN-TRANS
        'Eliminar cabeceras
        Cad = "Delete from linasipre where numaspre = " & Data1.Recordset!numaspre
        Conn.Execute Cad
        
        'Borramos sus lineas
        Cad = "Delete from cabasipre where numaspre = " & Data1.Recordset!numaspre
        NumRegElim = Data1.Recordset.AbsolutePosition
        Conn.Execute Cad

        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
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
        End If

    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdRegresar_Click()

If Data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

RaiseEvent DatoSeleccionado(Data1.Recordset.Fields(0) & "|" & Data1.Recordset.Fields(1) & "|")
Unload Me
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    LimpiarCampos
    
        ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
    End If
     Me.Height = 8640
    'Los campos auxiliares
    CamposAux False, 0, True
    
    'Si no es analitica no mostramos el label
    Text3(2).Visible = vParam.autocoste
    Label2(2).Visible = vParam.autocoste
    
    '## A mano
    NombreTabla = "cabasipre"
    Ordenacion = " ORDER BY numaspre"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
'    Data1.UserName = vUsu.Login
'    Data1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    
    PonerOpcionesMenu
    
    'Maxima longitud
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    txtaux(3).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE numaspre = -1"
    Data1.Refresh
    CargaGrid -1
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        Text1(0).BackColor = vbYellow
    End If

    
    CadAncho = False

End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
End Sub


'Private Sub Form_Resize()
'If Me.WindowState <> 0 Then
'
'    Exit Sub
'End If
'If Me.Width < 11610 Then Me.Width = 11610
'End Sub

Private Sub Form_Unload(Cancel As Integer)
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
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        'If CadB <> "" Then CadB = CadB & " AND "
        'CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas
If cmdAux(0).Tag = 0 Then
    'Cuenta normal
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
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
'Concepto
txtaux(4).Text = RecuperaValor(CadenaSeleccion, 1)
Text3(1).Text = RecuperaValor(CadenaSeleccion, 2) & " "
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    'Cta contrapartida
    cmdAux(0).Tag = 1
    LlamaContraPar
    txtaux(4).SetFocus
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
End Select
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
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
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
    If Modo <> 1 Then _
        FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
    

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
        Cad = Cad & ParaGrid(Text1(0), 20, "Asiento:")
        Cad = Cad & ParaGrid(Text1(1), 80, "Denominación")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1"
            frmB.vTitulo = "Asientos predefinidos"
            frmB.vSelElem = 1
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                Text1(kCampo).SetFocus
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
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
    CargaGrid Data1.Recordset!numaspre
    DataGrid1.Enabled = True
    
    frameextras.Visible = Not Adodc1.Recordset.EOF
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim b As Boolean
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For I = 0 To Text1.Count - 1
            'Text1(i).BackColor = vbWhite
            Text1(0).BackColor = &H80000018
        Next I
        'chkVistaPrevia.Visible = False
    End If
    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nuevo asiento predefinido"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar asiento predefinido"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar asiento predefinido"
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea asiento predefinido"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea asiento predefinido"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea asiento predefinido"
    End If
    b = (Modo < 5)
    chkVistaPrevia.Visible = b
    frameextras.Visible = b
    If b Then framelineas.Visible = False
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    Toolbar1.Buttons(10).Enabled = b  'Lineas factur
    If Not b Then frameextras.Visible = False
        
    
    DataGrid1.Enabled = b Or (Modo = 5)
    'Modo insertar o modificar
    b = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not b
    cmdAceptar.Visible = b Or Modo = 1

    '
    b = b Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    mnOpcionesAsiPre.Enabled = Not b
   
   
        'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5
    'Modificar
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    Else
        cmdRegresar.Visible = False
    End If
    b = b Or (Modo = 5)
    Toolbar1.Buttons(7).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b
    mnEliminar.Enabled = b

   
   
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    b = b Or Modo = 0   'En B tenemos modo=2 o a 5
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = b
        If b Then
            Text1(I).BackColor = &H80000018
        ElseIf Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    
    b = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = b
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    PonerOpcionesMenuGeneral Me
End Sub


Private Function DatosOk() As Boolean
    Dim Rs As ADODB.Recordset
    Dim b As Boolean
    b = CompForm(Me)
    DatosOk = b
End Function


'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    
    SQL = "Select Max(numaspre) from " & NombreTabla
    Text1(0).Text = 1
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            Text1(0).Text = Rs.Fields(0) + 1
        End If
    End If
    Rs.Close
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Modo <> 5 Then Adodc1.Recordset.Cancel
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
        BotonEliminar
    Else
        'ELIMINAR linea factura
        EliminarLineaFactura
    End If
Case 10
    'Nuevo Modo
    PonerModo 5
    'Fuerzo que se vean las lineas
    frameextras.Visible = True
    cmdCancelar.Caption = "Cabecera"
    lblIndicador.Caption = "Lineas detalle"
Case 11
    'imprimir
    frmListado.Opcion = 3
    frmListado.Show vbModal
Case 12
    Unload Me
Case 14 To 17
    Desplazamiento (Button.Index - 14)
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = bol
    Next I
End Sub

'--- A mano // control de devoluciones de prismáticos
Private Sub FrmB1_DatoSeleccionado(CadenaSeleccion As String) '-- Proveedores

End Sub

Private Sub CargaGrid(NumFac As Long)
Dim b As Boolean
b = DataGrid1.Enabled
CargaGrid2 NumFac
DataGrid1.Enabled = b
End Sub


Private Sub CargaGrid2(NumFac As Long)
    Dim anc As Single
    
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = MontaSQLCarga(NumFac)
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockPessimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
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
    DataGrid1.Columns(10).NumberFormat = "#,##0.00"
    DataGrid1.Columns(10).Width = 1154 + ancho
    DataGrid1.Columns(10).Alignment = dbgRight
            
    DataGrid1.Columns(11).Caption = "Haber"
    DataGrid1.Columns(11).NumberFormat = "#,##0.00"
    DataGrid1.Columns(11).Width = 1154 + ancho
    DataGrid1.Columns(11).Alignment = dbgRight
            
            
    If vParam.autocoste Then
        DataGrid1.Columns(12).Caption = "C.C."
        DataGrid1.Columns(12).Width = 510
    Else
        DataGrid1.Columns(12).Visible = False
    End If
    DataGrid1.Columns(13).Visible = False
    'Fiajamos el cadancho
    If Not CadAncho Then
        anc = 323
        txtaux(0).Left = DataGrid1.Left + anc
        txtaux(0).Width = DataGrid1.Columns(2).Width
        
        
        anc = 150
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(3).Left
        
        txtaux(1).Left = DataGrid1.Columns(3).Left + anc
        txtaux(1).Width = DataGrid1.Columns(3).Width - 30
        
    
        txtaux(2).Left = DataGrid1.Columns(4).Left + anc
        txtaux(2).Width = DataGrid1.Columns(4).Width - 30
        
    
        txtaux(3).Left = DataGrid1.Columns(5).Left + anc
        txtaux(3).Width = DataGrid1.Columns(5).Width - 30
        
        
        'Concepto
        txtaux(4).Left = DataGrid1.Columns(6).Left + anc
        txtaux(4).Width = DataGrid1.Columns(6).Width - 30
        
        
        txtaux(5).Left = DataGrid1.Columns(8).Left + anc
        txtaux(5).Width = DataGrid1.Columns(8).Width - 30
        
        
        
        txtaux(6).Left = DataGrid1.Columns(10).Left + anc
        txtaux(6).Width = DataGrid1.Columns(10).Width - 30
        
        
        txtaux(7).Left = DataGrid1.Columns(11).Left + anc
        txtaux(7).Width = DataGrid1.Columns(11).Width - 30
        
        txtaux(8).Left = DataGrid1.Columns(12).Left + anc
        txtaux(8).Width = DataGrid1.Columns(12).Width - 30
      
        CadAncho = True
    End If
    
    
    
    
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
    
    
    
    'Obtenemos las sumas
    ObtenerSumas
    
End Sub

Private Sub ObtenerSumas()
Dim Deb As Currency
Dim hab As Currency
Dim Rs As ADODB.Recordset
If Data1.Recordset.EOF Then
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    Exit Sub
End If

If Adodc1.Recordset.EOF Then
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    Exit Sub
End If



Set Rs = New ADODB.Recordset
SQL = "SELECT Sum(linasipre.timporteD) AS SumaDetimporteD, Sum(linasipre.timporteH) AS SumaDetimporteH,linasipre.numaspre"
SQL = SQL & " From linasipre"
SQL = SQL & " GROUP BY linasipre.numaspre"
SQL = SQL & " HAVING (((linasipre.numaspre)=" & Data1.Recordset!numaspre & "));"
Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
Deb = 0
hab = 0
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields(0)) Then Deb = Rs.Fields(0)
    If Not IsNull(Rs.Fields(1)) Then hab = Rs.Fields(1)
End If

Text2(0).Text = Format(Deb, FormatoImporte): Text2(1).Text = Format(hab, FormatoImporte)
'Metemos en DEB el total
Deb = Deb - hab
If Deb < 0 Then
    Text2(2).ForeColor = vbRed
    Else
    Text2(2).ForeColor = vbBlack
End If
Text2(2).Text = Format(Deb, FormatoImporte)

End Sub


Private Function MontaSQLCarga(vNumFac As Long) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    '--------------------------------------------------------------------
    Dim SQL As String

    SQL = "SELECT linasipre.numaspre,linasipre.linlapre, linasipre.codmacta, cuentas.nommacta, linasipre.numdocum,"
    SQL = SQL & " linasipre.ctacontr, linasipre.codconce, conceptos.nomconce as nombreconcepto, linasipre.ampconce,"
    SQL = SQL & " cuentas_1.nommacta as nomctapar, linasipre.timporteD, linasipre.timporteH, linasipre.codccost, cabccost.nomccost as centrocoste"
    SQL = SQL & " FROM (((linasipre INNER JOIN conceptos ON linasipre.codconce = conceptos.codconce)"
    SQL = SQL & " INNER JOIN cuentas ON linasipre.codmacta = cuentas.codmacta)"
    SQL = SQL & " LEFT JOIN cuentas AS cuentas_1 ON linasipre.ctacontr = cuentas_1.codmacta)"
    SQL = SQL & " LEFT JOIN cabccost ON linasipre.codccost = cabccost.codccost"
    SQL = SQL & " WHERE numaspre = " & vNumFac
    SQL = SQL & " ORDER BY linasipre.linlapre"

    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea(Limpiar As Boolean)
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    NumLin = ObtenerSigueinteNumeroLinea
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
    HabilitarImportes 0
    'Ponemos el foco
    txtaux(0).SetFocus
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    NumLin = Adodc1.Recordset!linlapre
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
    txtaux(0).Text = Adodc1.Recordset.Fields!Codmacta
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
 
End Sub

Private Sub EliminarLineaFactura()
If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
If ModificandoLineas <> 0 Then Exit Sub
SQL = "Seguro que desea eliminar la linea: " & Adodc1.Recordset.Fields(3) & " "
If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
    SQL = "Delete from linasipre WHERE numaspre =" & Data1.Recordset!numaspre
    SQL = SQL & " AND linlapre = " & Adodc1.Recordset!linlapre
    Conn.Execute SQL
    CargaGrid Data1.Recordset!numaspre
End If
End Sub



Private Function ObtenerSigueinteNumeroLinea() As Long
Dim Rs As ADODB.Recordset
Dim I As Long

Set Rs = New ADODB.Recordset
Rs.Open "SELECT Max(linlapre) FROM linasipre where numaspre =" & Text1(0).Text, Conn, adOpenDynamic, adLockOptimistic, adCmdText
I = 0
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields(0)) Then I = Rs.Fields(0)
End If
Rs.Close
ObtenerSigueinteNumeroLinea = I + 1
End Function



'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------

Private Sub HabilitarCentroCoste()
Dim hab As Boolean
Dim Ch As String
    If Not vParam.autocoste Then Exit Sub
    hab = False
    If txtaux(0).Text <> "" Then
            Ch = Mid(txtaux(0).Text, 1, 1)
            If Ch = vParam.grupogto Or Ch = vParam.grupovta Or Ch = vParam.grupoord Then hab = True
    Else
        txtaux(8).Text = ""
    End If
    If hab Then
        txtaux(8).BackColor = &H80000005
        Else
        txtaux(8).BackColor = &H80000018
    End If
    txtaux(8).Enabled = hab
End Sub
Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
Dim b As Boolean
DeseleccionaGrid DataGrid1
cmdCancelar.Caption = "Cancelar"
ModificandoLineas = xModo
b = (xModo = 0)
framelineas.Visible = Not b
'frameextras.Visible = b
cmdAceptar.Visible = Not b
cmdCancelar.Visible = Not b
frameextras.Visible = Not b
CamposAux Not b, alto, Limpiar
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

Private Sub txtAux_LostFocus(Index As Integer)
Dim RC As String
Dim Sng As Single

    'Si no estamos modificando o insertando lineas no hacemos na de na
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
         HabilitarCentroCoste
         If RC <> "" Then txtaux(0).SetFocus
         
     Case 3
         RC = txtaux(3).Text
         If CuentaCorrectaUltimoNivel(RC, SQL) Then
             txtaux(3).Text = RC
             Text3(0).Text = SQL
         Else
             MsgBox SQL, vbExclamation
             txtaux(3).Text = ""
             Text3(0).Text = ""
             txtaux(3).SetFocus
         End If
            
    Case 4
            If Not IsNumeric(txtaux(4).Text) Then
                MsgBox "El concepto debe de ser numérico", vbExclamation
                Exit Sub
            End If
            RC = "tipoconce"
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtaux(4).Text, "N", RC)
            If SQL = "" Then
                MsgBox "Concepto NO encontrado: " & txtaux(4).Text, vbExclamation
                txtaux(4).Text = ""
                RC = "0"
            Else
                SQL = SQL & " "
            End If
            HabilitarImportes CByte(Val(RC))
            Text3(1).Text = SQL
            txtaux(5).Text = SQL
    Case 6, 7
            'LOS IMPORTES
            
            
            If Not IsNumeric(txtaux(Index).Text) Then
                MsgBox "Importes deben ser numéricos.", vbExclamation
                On Error Resume Next
                txtaux(Index).Text = ""
                txtaux(Index).SetFocus
                Exit Sub
            End If
            
            
            'Es numerico
            SQL = TransformaPuntosComas(txtaux(Index).Text)
            Sng = Round(CSng(SQL), 2)
            txtaux(Index).Text = Format(Sng, "0.00")
            
            'Ponemos el otro campo a ""
            If Index = 6 Then
                txtaux(7).Text = ""
            Else
                txtaux(6).Text = ""
            End If
    Case 8
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
        AuxOK = "La cta. contrapartida no esta dad de alta en el sistema."
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
    If Not IsNumeric(txtaux(6).Text) Then
        AuxOK = "El importe DEBE debe ser numérico"
        Exit Function
    End If
End If

If txtaux(7).Text <> "" Then
    If Not IsNumeric(txtaux(7).Text) Then
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

AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean

On Error GoTo EInsertarModificar
InsertarModificar = False

If ModificandoLineas = 1 Then
    'INSERTAR LINEAS
    'INSERT INTO linasipre (numaspre, linlapre, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab) VALUES (1, 1, '4730001', '1', 1, NULL, NULL, NULL, NULL, NULL, NULL)
    SQL = "INSERT INTO linasipre (numaspre, linlapre, codmacta, numdocum, codconce,"
    SQL = SQL & "ampconce, timporteD, timporteH, codccost, ctacontr, idcontab) VALUES ("
    SQL = SQL & Data1.Recordset.Fields(0) & ","
    SQL = SQL & NumLin & ",'"
    SQL = SQL & txtaux(0).Text & "','"
    SQL = SQL & txtaux(2).Text & "',"
    SQL = SQL & txtaux(4).Text & ",'"
    SQL = SQL & txtaux(5).Text & "',"
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
    
    If txtaux(3).Text = "" Then
      SQL = SQL & ValorNulo & ","
      Else
      SQL = SQL & txtaux(3).Text & ","
    End If
    'Marca de entrada manual de datos
    SQL = SQL & "'contab')"
    
    
Else

    'MODIFICAR
    'UPDATE linasipre SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
    '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
    SQL = "UPDATE linasipre SET "
    
    SQL = SQL & " codmacta = '" & txtaux(0).Text & "',"
    SQL = SQL & " numdocum = '" & txtaux(2).Text & "',"
    SQL = SQL & " codconce = " & txtaux(4).Text & ","
    SQL = SQL & " ampconce = '" & txtaux(5).Text & "',"
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
    SQL = SQL & " Where numaspre = " & Data1.Recordset.Fields(0)
    SQL = SQL & " And linlapre = " & NumLin


End If
Conn.Execute SQL
InsertarModificar = True
Exit Function
EInsertarModificar:
    MuestraError Err.Number, "InsertarModificar linea asiento predefinido.", Err.Description
End Function
 

Private Sub LlamaContraPar()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.ConfigurarBalances = 3
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




Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub

