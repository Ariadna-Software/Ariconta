VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConExtr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta extractos cuentas"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmConExtr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwicth 
      Height          =   375
      Left            =   4800
      Picture         =   "frmConExtr.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Saldo ACTUAL / PERIODO"
      Top             =   7680
      Width           =   375
   End
   Begin VB.Frame FramePeriodo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      TabIndex        =   36
      Top             =   8040
      Width           =   6375
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   2200
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text6"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   3520
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text6"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   4840
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text6"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "S. PERIODO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   37
         Top             =   120
         Width           =   1575
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   360
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame framePregunta 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      Top             =   60
      Width           =   6255
      Begin VB.OptionButton optDucum 
         Caption         =   "Contrapartida"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   34
         Top             =   4320
         Width           =   1455
      End
      Begin VB.OptionButton optDucum 
         Caption         =   "Documento"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   1560
         Width           =   3915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1575
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Left            =   1440
         Picture         =   "frmConExtr.frx":6B5C
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha inicio"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   29
         Top             =   2685
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmConExtr.frx":755E
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha fin"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   28
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "frmConExtr.frx":75E9
         Top             =   2640
         Width           =   240
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   600
         Top             =   2400
         Width           =   5415
      End
      Begin VB.Shape Shape2 
         Height          =   1335
         Left            =   600
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label9 
         Caption         =   "Consulta de extractos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   600
         TabIndex        =   27
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Otra consulta"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cta. anterior"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cta siguiente"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver asiento"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   9780
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   7740
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text6"
      Top             =   7740
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   7140
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text6"
      Top             =   7740
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text6"
      Top             =   890
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   890
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   890
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   1320
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmConExtr.frx":7674
      Height          =   5895
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.TextBox Text5 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   840
      Width           =   3315
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label100 
      Caption         =   "Leyendo BD ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1440
      TabIndex        =   42
      Top             =   7800
      Width           =   3255
   End
   Begin VB.Label Label11 
      Caption         =   "Cargando datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   4000
      TabIndex        =   32
      Top             =   4000
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Shape Shape3 
      Height          =   375
      Left            =   4740
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label3 
      Caption         =   "DEBE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7440
      TabIndex        =   21
      Top             =   600
      Width           =   510
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "S .ACTUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   5400
      TabIndex        =   24
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9960
      TabIndex        =   23
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "HABER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   8520
      TabIndex        =   22
      Top             =   600
      Width           =   645
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5340
      Top             =   7770
      Width           =   1500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "HASTA PERIODO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   17
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label8 
      Caption         =   "Acumulado anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   13
      Top             =   1320
      Width           =   1650
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label101 
      Caption         =   "1990 de 1000"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   7800
      Width           =   1095
   End
End
Attribute VB_Name = "frmConExtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cuenta As String   'Si es con cuenta
Public EjerciciosCerrados As Boolean


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1

Dim SQL As String
Dim vSQL As String
Dim RC As String
Dim Mostrar As Boolean
Dim anc As Integer


Dim RT As Recordset
Private VieneDeIntroduccion As Boolean
Dim AnyoInicioEjercicio As String

Dim QuedanLineasDespuesModificar As Boolean

Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Label10.Caption = DBLet(Adodc1.Recordset!nommacta, "T")
    If Err.Number <> 0 Then
        Err.Clear
        Label10.Caption = ""
         
    End If
    Label10.Refresh
End Sub

Private Sub cmdAceptar_Click()
'CONSULTA EXTRACTOS
Dim F As Date
    If Text1.Text = "" Then
        MsgBox "Introduzca una cuenta", vbExclamation
        PonerFoco Text1
        Exit Sub
    End If
    If Text3(0).Text = "" Or Text3(1).Text = "" Then
        MsgBox "Introduce las fechas de consulta de extractos", vbExclamation
        Exit Sub
    End If
    
    If Text3(0).Text <> "" And Text3(1).Text <> "" Then
        If CDate(Text3(0).Text) > CDate(Text3(1).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Sub
        End If
    End If

    lblFecha.Caption = ""
    SQL = ""
    'Llegados aqui. Vemos la fecha y demas
    If Text3(0).Text <> "" Then
        SQL = " fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
        lblFecha.Caption = "Desde " & Text3(0).Text
    End If
    
    If Text3(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
        If lblFecha.Caption <> "" Then lblFecha.Caption = lblFecha.Caption
        lblFecha.Caption = lblFecha.Caption & " hasta " & Text3(1).Text
    End If
    Text3(0).Tag = SQL  'Para las fechas
    

    'Para ver si la cuenta tiene movimientos o no
    vSQL = "Select count(*) from hlinapu"
    If EjerciciosCerrados Then vSQL = vSQL & "1"
    vSQL = vSQL & " WHERE  fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    vSQL = vSQL & " AND fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    vSQL = vSQL & " AND codmacta ='"


    'Fijamos el año de incio de jercicio si es CERRADO
    F = CDate(Text3(0).Text)

    If Month(F) >= Month(vParam.fechaini) Then
        AnyoInicioEjercicio = Year(F)
    Else
        AnyoInicioEjercicio = Year(F) - 1
    End If







    'Ponemos la cuenta
    Text4.Text = Text1.Text
    Text5.Text = Text2.Text
    
    Screen.MousePointer = vbHourglass
    If Not TieneMovimientos(Text1.Text) Then
        MsgBox "La cuenta " & Text2.Text & " NO tiene movimientos en las fechas", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Me.framePregunta.Visible = False
    PonerTamaños True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    CargarDatos False
    If Cuenta = "" Then PonerFoco DataGrid1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Imprimir()
Dim MostrarAnterior As Byte
If Me.Adodc1.Recordset.RecordCount < 1 Then Exit Sub


'Resto parametros
Screen.MousePointer = vbHourglass

'borramos, todos menos la que esta mirando
SQL = "DELETE FROM tmpconextcab WHERE codusu =" & vUsu.Codigo & " AND cta<>'" & Text4.Text & "'"
Conn.Execute SQL
SQL = "DELETE FROM tmpconext WHERE codusu =" & vUsu.Codigo & " AND cta<>'" & Text4.Text & "'"
Conn.Execute SQL



SQL = ""


SQL = SQL & "Titulo= ""Extractos de cuentas""|"
'Fechas intervalor
SQL = SQL & "Fechas= ""Fechas:  desde " & Text3(0).Text & " hasta " & Text3(1).Text & """|"
'Cuentas
SQL = SQL & "Cuenta= """"|"
SQL = SQL & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
SQL = SQL & "NumPag= 0|"
SQL = SQL & "Salto= 2|"
MostrarAnterior = FechaInicioIGUALinicioEjerecicio(CDate(Text3(0).Text), EjerciciosCerrados)
SQL = SQL & "MostrarAnterior= " & MostrarAnterior & "|"

If Not GeneraraExtractos Then Exit Sub
Screen.MousePointer = vbDefault
With frmImprimir
    .OtrosParametros = SQL
    .NumeroParametros = 6
    .FormulaSeleccion = "{ado_lineas.codusu}=" & vUsu.Codigo
    '.SoloImprimir = True
    'Opcion dependera del combo
    .Opcion = 3
    .Show vbModal
End With

End Sub

Private Sub OtraCuenta(Index As Integer)
Dim I As Integer

If Cuenta <> "" Then Exit Sub

    'Obtener la cuenta
    If Not ObtenerCuenta(Index = 0) Then Exit Sub

    'Poner datos
    Screen.MousePointer = vbHourglass
    
    'Ponemos los text a blanco
        For I = 0 To 8
            Text6(I).Text = ""
        Next I
    Label100.Visible = True
    Label101.Caption = ""
    Label10.Caption = ""
    Me.DataGrid1.Enabled = False
    Me.DataGrid1.Visible = False
    Label11.Visible = True
    Me.Refresh
    DoEvents
    Screen.MousePointer = vbHourglass
    CargarDatos False
    Label100.Visible = False
    Me.DataGrid1.Visible = True
    Me.DataGrid1.Enabled = True
    DataGrid1.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSwicth_Click()
    Me.FramePeriodo.Visible = Not FramePeriodo.Visible
    If Me.FramePeriodo.Visible Then
        Me.cmdSwicth.ToolTipText = "Ver saldo ACTUAL"
         
    Else
        Me.cmdSwicth.ToolTipText = "Ver saldo PERIODO"
    End If
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub DataGrid1_DblClick()
Dim Numasien As Long
If Not VieneDeIntroduccion Then
    If Not Me.Adodc1.Recordset.EOF Then
        Screen.MousePointer = vbHourglass
        AsientoConExtModificado = 0
        SQL = Adodc1.Recordset!NumDiari & "|" & Adodc1.Recordset!fechaent & "|" & Adodc1.Recordset!Numasien & "|"
        frmHcoApuntes.EjerciciosCerrados = EjerciciosCerrados
        frmHcoApuntes.ASIENTO = SQL
        frmHcoApuntes.LINASI = Adodc1.Recordset!Linliapu
        frmHcoApuntes.Show vbModal
        espera 0.1
        If AsientoConExtModificado = 1 Then
            QuedanLineasDespuesModificar = True
            Numasien = Adodc1.Recordset!Numasien
            'Volvemos a recargar datos
            Screen.MousePointer = vbHourglass
            DataGrid1.Visible = False
            Me.Refresh
            Screen.MousePointer = vbHourglass
            CargarDatos True
            DataGrid1.Visible = True
            If QuedanLineasDespuesModificar Then
                'Intentamos buscar el asiento
                Adodc1.Recordset.Find "numasien = " & Numasien
                If Cuenta = "" Then PonerFoco DataGrid1
            Else
                'NO QUEDAN LINEAS
                HacerToolBar 1
            End If
            Screen.MousePointer = vbDefault
            Me.Refresh
        Else
'''''            Stop
        End If
    End If
Else
    MsgBox "Esta en la introduccion de apuntes.", vbExclamation
End If
End Sub

Private Sub PonerTamaños(Grande As Boolean)
    If Grande Then
        AjustarPantalla Me
        If Screen.Height >= 9200 Then
            Me.Height = 9105
        Else
            Me.Height = 8610
        End If
    Else
        Me.Top = 1200
        Me.Left = 1200
        Me.Width = Me.framePregunta.Width + 30
        Me.Height = Me.framePregunta.Height + 220
    End If
End Sub



Private Sub Form_Activate()
   ' If Cuenta <> "" Then
   '     Cuenta = ""
   '     cmdAceptar.SetFocus
   ' End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim J As Integer
Dim I As Integer

'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login
    Limpiar Me
    
    If EjerciciosCerrados Then
        SQL = "-1"
    Else
        AnyoInicioEjercicio = ""
        SQL = "0"
    End If
    
    ParametroPorDefecto True
    
    'Situaremos los textos
    '--------------------------------
    If vParam.autocoste Then
        anc = 1120
    Else
        anc = 1300
    End If
    
    
    If vParam.autocoste Then
        For I = 0 To 2
                J = I * 3
                Text6(0 + J).Left = 6315 + 270
                Text6(0 + J).Width = anc - 15
                Text6(1 + J).Left = Text6(0 + J).Left + anc + 15
                Text6(1 + J).Width = anc - 15
                Text6(2 + J).Left = Text6(1 + J).Left + anc + 15
                Text6(2 + J).Width = anc - 15 + 100
            
        Next I
        Label3.Left = Label3.Left - 300
        Label4.Left = Label4.Left - 450
        Label7.Left = Label7.Left - 550
    Else
        For I = 0 To 2
                J = I * 3
                Text6(0 + J).Left = 6840
                Text6(0 + J).Width = 1335
                Text6(1 + J).Left = 8160
                Text6(1 + J).Width = 1335
                Text6(2 + J).Left = 9480
                Text6(2 + J).Width = 1335
        Next I
        Label3.Left = 7400
        Label4.Left = 8520
        Label7.Left = 9960
    End If
    
    'Los del periodo
    For I = 6 To 8
        Text6(3 + I).Width = Text6(I).Width
        Text6(3 + I).Left = Text6(I).Left - 4940
    Next I
    
    
    
    If Screen.Height >= 9200 Then
        Me.FramePeriodo.Top = 8040
        cmdSwicth.Visible = False
        
        'Los labels
        'El de la nommacta
        Label10.FontSize = 10
        Label10.FontBold = True
        Label10.Left = 240
        Label10.Top = 8160
        Label10.Width = 4695
        Label10.Height = 255
        'El de 1 de 1
        Label101.FontSize = 10
        Label101.Width = 2295
        
    Else
        Me.FramePeriodo.Top = 7680
        cmdSwicth.Visible = True
        CargarFramePeridodVisible True
        
        
        'Los labels
        'El de la nommacta
        Label10.FontSize = 8
        Label10.FontBold = False
        Label10.Left = 1140
        Label10.Top = 7800
        Label10.Width = 3255
        Label10.Height = 220
        'El de 1 de 1
        Label101.FontSize = 8
        Label101.Width = 1095
        
    End If
    
    
    If EjerciciosCerrados Then
        I = -1
    Else
        I = 0
    End If
    
    
    Text3(0).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
    Text3(1).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
    VieneDeIntroduccion = False
    If Cuenta <> "" Then
        VieneDeIntroduccion = True
        Text1.Text = Cuenta
        SQL = ""
        CuentaCorrectaUltimoNivel Cuenta, SQL
        Text2.Text = SQL
'        cmdAceptar_Click
'        VaEnGrande = True
    End If
    Me.framePregunta.Visible = True
    'La toolbar
    With toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 18
        '.Buttons(1).Enabled = B
        .Buttons(2).Image = 16
        .Buttons(4).Image = 7
        
        .Buttons(5).Image = 8
        
        .Buttons(7).Image = 11
        '.Buttons(7).Enabled = Not EjerciciosCerrados
        .Buttons(9).Image = 15
    End With
    
    Caption = "Consulta de extractos"
    If EjerciciosCerrados Then Caption = Caption & " EJERCICIOS CERRADOS"
    
    
    PonerTamaños False
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ParametroPorDefecto False
    If Me.cmdSwicth.Visible Then CargarFramePeridodVisible False
End Sub

Private Sub frmC_Selec(vFecha As Date)
Text3(CByte(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text1.Text = RecuperaValor(CadenaSeleccion, 1)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    Image1(0).Tag = Index
    If Text3(Index).Text <> "" Then
        frmC.Fecha = CDate(Text3(Index).Text)
    Else
        frmC.Fecha = Now
    End If
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub imgCuentas_Click()
    Set frmCta = New frmColCtas
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub


Private Sub optDucum_KeyPress(Index As Integer, KeyAscii As Integer)
    Text1_KeyPress KeyAscii
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus()

    
    RC = Trim(Text1.Text)
    If RC = "" Then
        Text2.Text = ""
        Exit Sub
    End If
    If CuentaCorrectaUltimoNivel(RC, SQL) Then
        Text1.Text = RC
        Text2.Text = SQL
    Else
        MsgBox SQL, vbExclamation
        Text1.Text = ""
        Text2.Text = ""
        PonerFoco Text1
    End If
End Sub


Private Sub Text3_GotFocus(Index As Integer)
With Text3(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index).Text, vbExclamation
        Text3(Index).Text = ""
        PonerFoco Text3(Index)
        Exit Sub
    End If
End Sub


Private Function CargarDatos(DesdeModificarLinea As Boolean)
Dim N As Long
On Error GoTo ECargaDatos


    Label101.Caption = ""
    Label100.Visible = True
    Label100.Refresh
    '-------------  ANTES DE 19/11/04
    'SQL = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
    
    
    
    SQL = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo & " AND Cta = '" & Text4.Text & "'"
    Conn.Execute SQL
        
    'SQL = "DELETE from tmpconext where codusu= " & vUsu.Codigo
    SQL = "DELETE from tmpconext where codusu= " & vUsu.Codigo & " AND Cta = '" & Text4.Text & "'"
    Conn.Execute SQL
    
    If EjerciciosCerrados Then
        CargaDatosConExtCerrados Text4.Text, Text3(0).Text, Text3(1).Text, Text3(0).Tag, Text5.Text, AnyoInicioEjercicio
    Else
        CargaDatosConExt Text4.Text, Text3(0).Text, Text3(1).Text, Text3(0).Tag, Text5.Text
    End If
    
    If DesdeModificarLinea Then
        'Compruebo que haya ALGUN datos, si no explota
        SQL = "cta = '" & Text4.Text & "' AND codusu"
        N = DevuelveDesdeBD("count(*)", "tmpconext", SQL, vUsu.Codigo)
        If N = 0 Then
            QuedanLineasDespuesModificar = False
            Exit Function
        End If
    End If
    CargaGrid
    CargaImportes
    Label100.Visible = False
    Exit Function
ECargaDatos:
        MuestraError Err.Number, "Datos cuenta"
        Label100.Visible = False
End Function





Private Sub CargaGrid()


    Adodc1.ConnectionString = Conn
    'Modificacion 12 Junio 2009
    'ORDEN NORMAL
    '(codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, ampconce, timporteD, timporteH, saldo, Punteada, contra, ccost)
    'If Me.optDucum(0).Value Then
    '    SQL = " codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, ampconce, timporteD, timporteH, saldo,ccost, Punteada, contra"
    'Else
    '    SQL = " codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, contra, ampconce, timporteD, timporteH, saldo,ccost, Punteada, nomdocum"
    'End If
    If Me.optDucum(0).Value Then
        SQL = " codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, ampconce, timporteD, timporteH, saldo,ccost, Punteada, contra"
    Else
        SQL = " codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, contra, ampconce, timporteD, timporteH, saldo,ccost, Punteada, nomdocum"
    End If
    If Text4.Text <> "" Then
        SQL = SQL & ",nommacta"
        SQL = "Select " & SQL & " from tmpConExt left join cuentas on tmpConExt.contra=cuentas.codmacta  WHERE codusu = " & vUsu.Codigo
    Else
        'Si esta a "" pongo otro select para que no de error
        SQL = SQL & ",linliapu"
        SQL = "Select " & SQL & " from tmpConExt where codusu = " & vUsu.Codigo
    End If
    SQL = SQL & " AND cta = '" & Text4.Text & "' ORDER BY POS"
    
    'Si text4.text=""
    
    Adodc1.RecordSource = SQL
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 270
    
    
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Visible = False
    DataGrid1.Columns(2).Visible = False
    DataGrid1.Columns(3).Visible = False

    'Cuenta
    DataGrid1.Columns(4).Caption = "Fecha"
    If vParam.autocoste Then
        'Tiene analitica
        DataGrid1.Columns(4).Width = 900
        DataGrid1.Columns(4).NumberFormat = "dd/mm/yy"
    Else
        DataGrid1.Columns(4).Width = 1200
        DataGrid1.Columns(4).NumberFormat = "dd/mm/yyyy"
    End If
    DataGrid1.Columns(5).Caption = "NºAsie."
    DataGrid1.Columns(5).Width = 900
    
    DataGrid1.Columns(6).Visible = False
    
    If Me.optDucum(0).Value Then
        DataGrid1.Columns(7).Caption = "Documento"
    Else
        DataGrid1.Columns(7).Caption = "Contra."
    End If
    DataGrid1.Columns(7).Width = 1200

    DataGrid1.Columns(8).Caption = "Ampliación"
    DataGrid1.Columns(8).Width = 3000
    
    DataGrid1.Columns(9).Caption = "Debe"
    DataGrid1.Columns(9).Width = anc
    DataGrid1.Columns(9).Alignment = dbgRight
    DataGrid1.Columns(9).NumberFormat = "#,##0.00"
    
    DataGrid1.Columns(10).Caption = "Haber"
    DataGrid1.Columns(10).Width = anc
    DataGrid1.Columns(10).Alignment = dbgRight
    DataGrid1.Columns(10).NumberFormat = "#,##0.00"
    
    DataGrid1.Columns(11).Caption = "Saldo"
    DataGrid1.Columns(11).Width = anc + 100
    DataGrid1.Columns(11).Alignment = dbgRight
    DataGrid1.Columns(11).NumberFormat = "#,##0.00"
    
    'Centro de coste
    DataGrid1.Columns(12).Visible = vParam.autocoste
    If vParam.autocoste Then
        DataGrid1.Columns(12).Width = 700
        DataGrid1.Columns(12).Caption = " C.C."
    End If
    
    DataGrid1.Columns(13).Caption = " P"
    DataGrid1.Columns(13).Width = 500
    DataGrid1.Columns(13).Alignment = dbgCenter
    
    DataGrid1.Columns(14).Visible = False
    DataGrid1.Columns(15).Visible = False
    If Me.cmdSwicth.Visible Then
        Label101.Caption = "Lineas:"
    Else
        Label101.Caption = "Total lineas:   "
    End If
    Label101.Caption = Label101.Caption & Me.Adodc1.Recordset.RecordCount
    
End Sub







Private Function ObtenerCuenta(Siguiente As Boolean) As Boolean
    Label101.Caption = ""
    Label100.Visible = True
    Label100.Refresh
    SQL = "select codmacta from hlinapu"
    If EjerciciosCerrados Then SQL = SQL & "1"
    SQL = SQL & " WHERE codmacta "
    If Siguiente Then
        SQL = SQL & ">"
    Else
        SQL = SQL & "<"
    End If
    SQL = SQL & " '" & Text4.Text & "'"
    SQL = SQL & " AND  fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    SQL = SQL & " group by codmacta ORDER BY codmacta"
    If Siguiente Then
        SQL = SQL & " ASC"
    Else
        SQL = SQL & " DESC"
    End If
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        SQL = "No se ha obtenido la cuenta "
        If Siguiente Then
            SQL = SQL & "siguiente"
        Else
            SQL = SQL & "anterior"
        End If
        SQL = SQL & " con movimientos en el periodo."
        MsgBox SQL, vbExclamation
        ObtenerCuenta = False
    Else
        Text4.Text = RT!codmacta
        Text5.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RT!codmacta, "T")
        ObtenerCuenta = True
    End If
    RT.Close
    Set RT = Nothing
    Label100.Visible = False
End Function



Private Sub CargaImportes()
Dim I As Integer
Dim Im1 As Currency
Dim Im2 As Currency


    SQL = "Select * from tmpconextcab where codusu=" & vUsu.Codigo & " and cta='" & Text4.Text & "'"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        'Limpiaremos
        For I = 0 To 11
            Text6(I).Text = ""
        Next I
    Else
        Im1 = 0: Im2 = 0
        For I = 0 To 8
            Text6(I).Text = Format(RT.Fields(I + 4), FormatoImporte)
        Next I
        
        'Importes calculado del periodo
        Im1 = RT.Fields(7) - RT.Fields(4)
        Im2 = RT.Fields(8) - RT.Fields(5)
        Text6(9).Text = Format(Im1, FormatoImporte)
        Text6(10).Text = Format(Im2, FormatoImporte)
        Im1 = Im1 - Im2
        Text6(11).Text = Format(Im1, FormatoImporte)
        
        
    End If
    RT.Close
End Sub



Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode < 41 Then TeclaPulsada KeyCode
        If KeyCode <> 27 Then PonerFoco Text4
End Sub


Private Sub TeclaPulsada(Codigo As Integer)
    Select Case Codigo
    Case 37 To 40
        If Codigo = 39 Or Codigo = 40 Then
            OtraCuenta 0
        Else
            OtraCuenta 1
        End If
    Case 13
        
    Case 27
        Unload Me
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub
Private Sub HacerToolBar(ButtonIndex As Integer)
Select Case ButtonIndex
Case 1
    Screen.MousePointer = vbHourglass
    Text4.SetFocus
    Text1.Text = Text4.Text
    Text2.Text = Text5.Text
    
    'Pongo a "" para cargar el grid a vacio
    Text4.Text = ""
    CargaGrid
    PonerTamaños False

    Me.framePregunta.Visible = True
    PonerFoco DataGrid1
    Screen.MousePointer = vbDefault
Case 2
    Imprimir
Case 4
    OtraCuenta 1
Case 5
    OtraCuenta 0
Case 7
    DataGrid1_DblClick
    
Case 9
    Unload Me
End Select
End Sub


Private Function TieneMovimientos(Cuenta As String) As Boolean
Dim RT As ADODB.Recordset
    
    Set RT = New ADODB.Recordset
    RT.Open vSQL & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TieneMovimientos = False
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then
            If Val(RT.Fields(0)) > 0 Then TieneMovimientos = True
        End If
    End If
    RT.Close
    Set RT = Nothing
End Function


Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub ParametroPorDefecto(Leer As Boolean)
    On Error GoTo ESQL
    
    SQL = App.path & "\conext.xdf"
    If Leer Then
        If Dir(SQL, vbArchive) = "" Then
            anc = 0
        Else
            anc = 1
        End If
        optDucum(anc).Value = True
    Else
        'Escribir
        If Me.optDucum(0).Value Then
            If Dir(SQL, vbArchive) <> "" Then Kill SQL
        Else
            anc = FreeFile
            Open SQL For Output As #anc
            Print #anc, Now
            Close #anc
            
        End If
    End If
    Exit Sub
ESQL:
    MuestraError Err.Number, "Valores por defecto en el formulario"
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()

WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()

  WheelUnHook
End Sub



Private Sub CargarFramePeridodVisible(Leer As Boolean)
Dim B As Boolean

    On Error GoTo ESQL2
    
    SQL = App.path & "\conext2.xdf"
    If Leer Then
        B = (Dir(SQL, vbArchive) <> "")
        Me.FramePeriodo.Visible = B
        Me.FramePeriodo.Tag = Abs(B)
        If Not B Then
            cmdSwicth.ToolTipText = "Ver saldo PERIODO"
        Else
            cmdSwicth.ToolTipText = "Ver saldo ACTUAL"
        End If
    Else
        'Escribir
        If Abs(Me.FramePeriodo.Visible) <> Val(FramePeriodo.Tag) Then
            If Not Me.FramePeriodo.Visible Then
                If Dir(SQL, vbArchive) <> "" Then Kill SQL
            Else
                If Dir(SQL, vbArchive) = "" Then
                    anc = FreeFile
                    Open SQL For Output As #anc
                    Print #anc, Now
                    Close #anc
                End If
            End If
        End If
    End If
    Exit Sub
ESQL2:
    MuestraError Err.Number, "Valores por defecto cmd"
        
End Sub
