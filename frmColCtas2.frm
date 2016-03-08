VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmColCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   Icon            =   "frmColCtas2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   8160
      TabIndex        =   24
      Top             =   5190
      Width           =   1395
      Begin VB.OptionButton Option2 
         Caption         =   "Cod. Cta"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   25
         Top             =   600
         Width           =   1035
      End
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   4560
      TabIndex        =   23
      Top             =   6240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAccion 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   4
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   3
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   2100
      TabIndex        =   1
      Top             =   5160
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   4545
      Left            =   8160
      TabIndex        =   15
      Top             =   600
      Width           =   1425
      Begin VB.CheckBox Check1 
         Caption         =   "9º nivel"
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   4245
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "8º nivel"
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   3805
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7º nivel"
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   3365
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6º nivel"
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   2925
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5º nivel"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2485
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4º nivel"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   2045
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3º nivel"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2º nivel"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1165
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1er nivel"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   725
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Último:  "
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Value           =   1  'Checked
         Width           =   1125
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColCtas2.frx":000C
      Height          =   5445
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   9604
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6030
      Top             =   30
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Busqueda avanzada"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cuentas libres"
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
            Object.Tag             =   "2"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver observaciones Plan Contable"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Comprobar cuentas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6960
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label lblComprobar 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   22
      Top             =   6300
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblComprobar 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   21
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Menu mnOpciones 
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
Attribute VB_Name = "frmColCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public ConfigurarBalances As Byte
    '0.- Normal
    '1.- Busqueda
    '2.- Agrupacion de cuentas
    '3.- BUSQUEDA NUEVA
    '4.- Nueva cuenta
    '5.- Busquedas de envio de e-mail
    '6.- Exclusion de cuentas en consolidado. Como la agrupacion pero acepta niveles inferiores al penultimo

Public Event DatoSeleccionado(CadenaSeleccion As String)


Private CadenaConsulta As String
Dim CadAncho As Boolean 'Para cuando llamemos al al form de lineas
Dim RS As Recordset
Dim NF As Integer
Dim Errores As Long
Dim PrimeraVez As Boolean
Dim Aux As String

'Dim Clik1 As Boolean

Private Sub BotonAnyadir(Cuenta As String)
    ParaBusqueda False
    frmCuentas.vModo = 1
    frmCuentas.CodCta = Cuenta
    CadenaDesdeOtroForm = ""
    frmCuentas.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        CargaGrid
        'Intentamos situar el grid
        SituaGrid CadenaDesdeOtroForm
        If Me.ConfigurarBalances = 4 Then cmdRegresar_Click
    Else
        If Me.ConfigurarBalances = 4 Then CargaGrid
    End If
End Sub

Private Sub BotonBuscar()
    CadenaConsulta = GeneraSQL("codmacta= 'David'")  'esto es para que no cargue ningun registro
    CargaGrid
    ParaBusqueda True
    txtAux(0).Text = "": txtAux(1).Text = ""
    PonerFoco txtAux(1)
End Sub

Private Sub BotonVerTodos()
    ParaBusqueda False
    CadenaConsulta = GeneraSQL("")
    CargaGrid
End Sub



Private Sub BotonModificar()
    
    ParaBusqueda False
    CadenaDesdeOtroForm = ""
    frmCuentas.vModo = 2
    frmCuentas.CodCta = adodc1.Recordset!codmacta
    frmCuentas.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CargaGrid
        SituaGrid CadenaDesdeOtroForm
    End If
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    ParaBusqueda False
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar la cuenta:"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = adodc1.Recordset.Fields(0)
        Screen.MousePointer = vbHourglass
        If SepuedeEliminarCuenta(SQL) Then
            'Hay que eliminar
            Screen.MousePointer = vbHourglass
            SQL = "Delete from cuentas where codmacta='" & adodc1.Recordset!codmacta & "'"
            Conn.Execute SQL
            Screen.MousePointer = vbHourglass
            espera 0.5
            'Cancelamos el adodc1
            DataGrid1.Enabled = False
            NumRegElim = adodc1.Recordset.AbsolutePosition - 1
            CargaGrid
            DataGrid1.Enabled = True
            If NumRegElim > 0 Then
                If NumRegElim >= adodc1.Recordset.RecordCount Then
                    adodc1.Recordset.MoveLast
                Else
                    adodc1.Recordset.Move NumRegElim
                    'DataGrid1.Bookmark = Adodc1.Recordset.AbsolutePosition
                End If
            End If
            
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            Errores = Err.Number
            SQL = Err.Description
            FijarError SQL
            MsgBox "Error eliminando cuenta. " & vbCrLf & SQL, vbExclamation
        End If
End Sub

Private Sub FijarError(ByRef Cad As String)
    On Error Resume Next
    Cad = Conn.Errors(0).Description
    If Err.Number <> 0 Then
        Err.Clear
        Cad = ""
    End If
End Sub


Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adReason = adRsnMove And adStatus = adStatusOK Then PonLblIndicador Me.lblIndicador, adodc1
        
End Sub



Private Sub Check1_Click(Index As Integer)
    If PrimeraVez Then Exit Sub
    OpcionesCambiadas
End Sub

Private Sub OpcionesCambiadas()
    If txtAux(0).Visible Then Exit Sub

    If Not adodc1.Recordset.EOF Then
        If adodc1.Recordset.RecordCount > 0 Then
            If MsgBox("¿Desea refrescar los datos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    
    CadenaConsulta = GeneraSQL("")
    CargaGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim SQL As String


    If Index = 0 Then
        'Ha pulsado aceptar
        txtAux(0).Text = Trim(txtAux(0).Text)
        txtAux(1).Text = Trim(txtAux(1).Text)
        'Si estan vacios no hacemos nada
        SQL = ""
        Aux = ""
        If txtAux(0).Text <> "" Then
            If SeparaCampoBusqueda("T", "codmacta", txtAux(0).Text, Aux) = 0 Then SQL = Aux
        End If
        If txtAux(1).Text <> "" Then
            Aux = ""
            
            'VEO si ha puesto un *
            If InStr(1, txtAux(1).Text, "*") = 0 Then txtAux(1).Text = "*" & txtAux(1).Text & "*"
            If SeparaCampoBusqueda("T", "nommacta", txtAux(1).Text, Aux) = 0 Then
                If SQL <> "" Then SQL = SQL & " AND "
                SQL = SQL & Aux
            End If
        End If
        
        'Si sql<>"" entonces hay puestos valores
        If SQL = "" Then Exit Sub
        
        'Llamamos a carga grid
        Screen.MousePointer = vbHourglass
        CadenaConsulta = GeneraSQL(SQL)
        CargaGrid
        Screen.MousePointer = vbDefault
        If adodc1.Recordset.EOF Then
            MsgBox "Ningún resultado para la búsqueda.", vbExclamation
            Exit Sub
        Else
            
        End If
        PonerFoco DataGrid1
    End If
    ParaBusqueda False
    'lblIndicador.Caption = ""
End Sub

Private Sub cmdRegresar_Click()
    If adodc1.Recordset Is Nothing Then
        BotonBuscar
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    RaiseEvent DatoSeleccionado(adodc1.Recordset!codmacta & "|" & adodc1.Recordset!nommacta & "|" & adodc1.Recordset!bloqueada & "|")
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then
        cmdRegresar_Click
    Else
    
        If adodc1.Recordset Is Nothing Then Exit Sub
        If adodc1.Recordset.EOF Then Exit Sub
    
        'Vemos todos los valores de la cuenta
        frmCuentas.vModo = 0
        frmCuentas.CodCta = adodc1.Recordset!codmacta
        frmCuentas.Show vbModal
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        
        PonerOptionsVisibles 1
        PrimeraVez = False
        DoEvents
        'Vamos a ver si funciona 30 Sept 2003
        Select Case ConfigurarBalances
        Case 0, 1, 2, 5, 6
            If ConfigurarBalances = 2 Or ConfigurarBalances = 6 Then CadenaConsulta = GeneraSQL("")
            If ConfigurarBalances = 5 Then
                'Estoy buscando los que tienen e-mail
                CadenaConsulta = CadenaConsulta & " AND maidatos <> ''"
            End If
            CargaGrid
        
        Case 3
            BotonBuscar
        Case 4
            
            BotonAnyadir CadenaDesdeOtroForm
            
       
       
        End Select
        CadenaDesdeOtroForm = ""
    End If
    Screen.MousePointer = vbDefault
End Sub


'
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
     
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 24
        .Buttons(3).Image = 2
        .Buttons(4).Image = 21  'Busqueda ctas libres
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 19
        .Buttons(11).Image = 16
        .Buttons(12).Image = 20
        .Buttons(14).Image = 15
        
        'campos que no se pueden ver si vienen de..
        .Buttons(4).Visible = ConfigurarBalances <> 2 And ConfigurarBalances <> 6
        
    End With
     
     
     
     
    PrimeraVez = True
    pb1.Visible = False
    'Poner niveles
    PonerOptionsVisibles 0
    
    'Opciones segun sea su nivel
    PonerOpcionesMenu
    
    'Ocultamos busqueda
    ParaBusqueda False
    
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.Password = vUsu.Passwd
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    Frame2.Enabled = (DatosADevolverBusqueda = "")
    CadAncho = False
    'Cadena consulta
    CadenaConsulta = GeneraSQL("")
    
    lblIndicador.Caption = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    ConfigurarBalances = 0
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
    BotonAnyadir ""
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub



'----------------------------------------------------------------


Private Sub Option2_Click(Index As Integer)
    If PrimeraVez Then Exit Sub
'    Clik1 = Not Clik1
'
'    If Not Clik1 Then Exit Sub
    OpcionesCambiadas
End Sub

Private Sub Option2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String
    WheelUnHook
    Select Case Button.Index
    Case 1
            BotonBuscar
    Case 2
        'Busqueda avanzada
            CadenaDesdeOtroForm = ""
            frmCuentas.vModo = 3
            frmCuentas.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                Me.Refresh
                Screen.MousePointer = vbHourglass
                ParaBusqueda False
                PonerResultadosBusquedaAvanzada
                Screen.MousePointer = vbDefault
            End If
    Case 3
            BotonVerTodos
            
            
    Case 4
            'Busqueda ctas libres
            Screen.MousePointer = vbHourglass
            CadenaDesdeOtroForm = ""
            frmUtilidades.Opcion = 7
            frmUtilidades.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                C = CadenaDesdeOtroForm
                CadenaDesdeOtroForm = ""
                BotonAnyadir C
                If CadenaDesdeOtroForm <> "" Then
                    Me.Refresh
                    DoEvents
                    Screen.MousePointer = vbHourglass
                    CadenaDesdeOtroForm = " codmacta = '" & CadenaDesdeOtroForm & "'"
                    PonerResultadosBusquedaAvanzada
                    Screen.MousePointer = vbDefault
                End If
            End If
    Case 6
            BotonAnyadir ""
    Case 7
            BotonModificar
    Case 8
            BotonEliminar
    
    Case 10
            'Ver observaciones para las cuentas a 3 digitos
            VerObservaciones
    Case 11
            'Imprimimos el listado
                Screen.MousePointer = vbHourglass
                frmListado.Opcion = 2 'Listado de cuentas
                frmListado.Show vbModal
                
                
                
    Case 12
            'Comprobar cuentas
            ComprobarCuentas
    Case 14
            Unload Me
    Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim I
    For I = 16 To 19
        Toolbar1.Buttons(I).Enabled = bol
        Toolbar1.Buttons(I).Visible = bol
    Next I
End Sub


Private Sub CargaGrid()
Dim B As Boolean

    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    espera 0.1
    lblComprobar(0).Visible = True
    lblComprobar(0).Caption = "Leyendo BD"
    lblComprobar(0).Refresh
    CargaGrid2
    PonerFoco Check1(0)
    lblComprobar(0).Visible = False
    DataGrid1.Enabled = B
End Sub

Private Sub CargaGrid2()
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    Dim SQL As String
    Dim B As Boolean
    adodc1.ConnectionString = Conn
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    SQL = CadenaConsulta
    SQL = SQL & " ORDER BY"
    If Option2(0).Value Then
        SQL = SQL & " codmacta"
    Else
        SQL = SQL & " nommacta"
    End If
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    
   
        DataGrid1.Columns(0).Caption = "Cuenta"
        DataGrid1.Columns(0).Width = 1200
    
   
        DataGrid1.Columns(1).Caption = "Denominación"
        DataGrid1.Columns(1).Width = 5300
        TotalAncho = TotalAncho + DataGrid1.Columns(1).Width
    
   
        DataGrid1.Columns(2).Caption = "Direc."
        DataGrid1.Columns(2).Width = 500
        TotalAncho = TotalAncho + DataGrid1.Columns(2).Width
               
        DataGrid1.Columns(3).Caption = "Bloq"
        DataGrid1.Columns(3).Width = 400
        TotalAncho = TotalAncho + DataGrid1.Columns(3).Width
               
               
               
        If Not CadAncho Then
            txtAux(0).Left = DataGrid1.Columns(0).Left + 150
            txtAux(0).Width = DataGrid1.Columns(0).Width - 30
            txtAux(0).Top = DataGrid1.Top + 235
            txtAux(1).Left = DataGrid1.Columns(1).Left + 150
            txtAux(1).Width = DataGrid1.Columns(1).Width - 30
            txtAux(1).Top = txtAux(0).Top
            txtAux(0).Height = DataGrid1.RowHeight - 15
            txtAux(1).Height = txtAux(0).Height
            CadAncho = True
        End If
               
    'Habilitamos modificar y eliminar
    Toolbar1.Buttons(4).Enabled = vUsu.Nivel < 3
    Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    Else
        Toolbar1.Buttons(8).Enabled = False
    End If
   
    'Para k la barra de desplazamiento sea mas alta
    If Not adodc1.Recordset.EOF Then
            DataGrid1.ScrollBars = dbgVertical
    End If
    DataGrid1.Enabled = B
End Sub


' 0 solo textos
'1 Solo enables
'2 todo
Private Sub PonerOptionsVisibles(Opcion As Byte)
Dim I As Integer
Dim J As Integer
Dim Cad As String

    'Utilizo la variable cadancho
If Opcion <> 1 Then
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        If J > 0 Then
            Cad = "Digitos: " & J
            Check1(I).Caption = Cad
            Check1(I).Tag = J
        Else
            Check1(I).Caption = "Error"
        End If
        Check1(I).Value = 0
    Next I
    'Ultimo nivel
    J = DigitosNivel(I)
    If J > 0 Then Check1(0).Caption = Check1(0).Caption & J
    For I = vEmpresa.numnivel To 9
        Check1(I).Visible = False
    Next I
End If
If Opcion <> 0 Then
    Select Case ConfigurarBalances
    Case 1
        For I = 1 To vEmpresa.numnivel - 1
            J = DigitosNivel(I)
            If J < 5 Then  'A balances van ctas de 4 digitos
               Check1(I).Value = 1
            Else
               Check1(I).Value = 0
            End If
        Next I
        Check1(0).Value = 1
    Case 2
        'Agrupar ctas digitos . Realmete agrupamos al nivel de cuentas -1
        J = DevuelveDigitosNivelAnterior
        For I = 0 To 9
            Check1(I).Visible = I = J
            Check1(I).Value = Abs(I = J)
        Next I
        
    Case 6
        'Todos los niveles menos el ultimo
        'Agrupar ctas digitos . Realmete agrupamos al nivel de cuentas -1
        J = DevuelveDigitosNivelAnterior
        For I = 0 To 9
            Check1(I).Visible = I < J And I > 0
            Check1(I).Value = Abs(I <= J) And I > 0
        Next I
    Case Else
        Check1(0).Value = 1
    End Select
        
End If
End Sub



Private Function GeneraSQL(Busqueda As String) As String
Dim I As Integer
Dim SQL As String
Dim nexo As String
Dim J As Integer
Dim wildcar As String
Dim Digitos As Integer

SQL = ""
nexo = ""
If Check1(0).Value Then
    SQL = "( apudirec = 'S')"
    nexo = " OR "
End If
For I = 1 To vEmpresa.numnivel - 1
    If Check1(I).Value = 1 Then
        Digitos = Val(Check1(I).Tag)
        If Digitos = 0 Then Digitos = I
        wildcar = ""
        For J = 1 To Digitos
            wildcar = wildcar & "_"
        Next J
        SQL = SQL & nexo & " ( codmacta like '" & wildcar & "')"
        nexo = " OR "
    End If
Next I
wildcar = "SELECT codmacta, nommacta, apudirec,if(fecbloq is null,"""",""*"") as bloqueada"
wildcar = wildcar & " FROM cuentas "


'Nexo
nexo = " WHERE "
If Busqueda <> "" Then
    wildcar = wildcar & " WHERE (" & Busqueda & ")"
    nexo = " AND "
End If
If SQL <> "" Then wildcar = wildcar & nexo & "(" & SQL & ")"

GeneraSQL = wildcar
End Function



Private Function SepuedeEliminarCuenta(Cuenta As String) As Boolean
Dim NivelCta As Integer
Dim I, J As Integer
Dim Cad As String

    SepuedeEliminarCuenta = False
    If EsCuentaUltimoNivel(Cuenta) Then
        'ATENCION###
        ' Habra que ver casos particulares de eliminacion de una subcuenta de ultimo nivel
        'Si esta en apuntes, en ....
        'NO se puede borrar
        lblComprobar(0).Caption = "Comprobando"
        lblComprobar(0).Visible = True
        Cad = BorrarCuenta(Cuenta, lblComprobar(0))
        lblComprobar(0).Visible = False
        If Cad <> "" Then
            Cad = Cuenta & vbCrLf & Cad
            MsgBox Cad, vbExclamation
            Exit Function
        End If
        
    Else
        'No
        'No
        'no es una cuenta de ultimo nivel
        NivelCta = NivelCuenta(Cuenta)
        If NivelCta < 1 Then
            MsgBox "Error obteniendo nivel de la subcuenta", vbExclamation
            Exit Function
        End If
        
        'Ctas agrupadas
        I = DigitosNivel(NivelCta)
        If I = 3 Then
            Cad = DevuelveDesdeBD("codmacta", "ctaagrupadas", "codmacta", Cuenta, "T")
            If Cad <> "" Then
                MsgBox "El subnivel pertenece a agrupacion de cuentas en balance"
                Exit Function
            End If
        End If
        For J = NivelCta + 1 To vEmpresa.numnivel
            Cad = Cuenta & "__________"
            I = DigitosNivel(J)
            Cad = Mid(Cad, 1, I)
            If TieneEnBD(Cad) Then
                MsgBox "Tiene cuentas en niveles superiores (" & J & ")", vbExclamation
                Exit Function
            End If
        Next J
    End If
    SepuedeEliminarCuenta = True
End Function

Private Function TieneEnBD(Cad As String) As String
    'Dim Cad1 As String
    
    Set RS = New ADODB.Recordset
    'Cad1 = "Select codmacta from cuentas where codmacta like '" & Cad & "'"
    RS.Open "Select codmacta from cuentas where codmacta like '" & Cad & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TieneEnBD = Not RS.EOF
    RS.Close
    Set RS = Nothing
End Function


'Private Function BorrarCuenta(Cuenta As String) As Boolean
'On Error GoTo Salida
'Dim SQL As String
'
'
'pb1.Max = 6
'pb1.Value = 0
'pb1.Visible = True
'
''Con ls tablas declarads sin el ON DELETE , no dejara borrar
'BorrarCuenta = False
'Set RS = New ADODB.Recordset
'
'
''lineas de apuntes, contrapartidads   -->1
'RS.Open "Select * from linasipre where ctacontr ='" & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RS.EOF Then
'    RS.Close
'    GoTo Salida
'End If
'RS.Close
'pb1.Value = 1
'
''-->2
'RS.Open "Select * from linapu where ctacontr ='" & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RS.EOF Then
'    RS.Close
'    GoTo Salida
'End If
'RS.Close
'pb1.Value = 2
'
''-->3
''Otras tablas
''Reparto de gastos para inmovilizado
'SQL = "Select codmacta2 from sbasin where codmacta2='" & Cuenta & "'"
'RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RS.EOF Then
'    RS.Close
'    GoTo Salida
'End If
'RS.Close
'pb1.Value = 3
'
'
'
''-->4
'RS.Open "Select * from presupuestos where codmacta ='" & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RS.EOF Then
'    RS.Close
'    GoTo Salida
'End If
'RS.Close
'pb1.Value = 4
'
'
''-->5    Referencias a ctas desde eltos de inmovilizado
'SQL = "select codinmov from sinmov where codmact1='" & Cuenta & "'"
'SQL = SQL & " or codmact2='" & Cuenta & "'"
'SQL = SQL & " or codmact3='" & Cuenta & "'"
'RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RS.EOF Then
'    RS.Close
'    GoTo Salida
'End If
'RS.Close
'pb1.Value = 5
'
'
'
''-->6    Referencias a ctas desde eltos de inmovilizado
'SQL = "select codiva from samort where codiva='" & Cuenta & "'"
'RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'If Not RS.EOF Then
'    RS.Close
'    GoTo Salida
'End If
'RS.Close
'pb1.Value = 6
'
'
''SI kkega aqui es k ha ido bien
'BorrarCuenta = True
'Salida:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar ctas." & Err.Description
'    Set RS = Nothing
'    pb1.Visible = False
'End Function


Private Sub SituaGrid(CADENA As String)
On Error GoTo ESituaGrid
If adodc1.Recordset.EOF Then Exit Sub

adodc1.Recordset.Find " codmacta =  " & CADENA & ""
If adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst

Exit Sub
ESituaGrid:
    MuestraError Err.Number, "Situando registro activo"
End Sub


Private Sub ParaBusqueda(Ver As Boolean)
txtAux(0).Visible = Ver
txtAux(1).Visible = Ver
cmdAccion(0).Visible = Ver
cmdAccion(1).Visible = Ver
If Ver Then lblIndicador.Caption = "Búsqueda"
End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
    With txtAux(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub ComprobarCuentas()
Dim Cad As String
Dim N As Integer
Dim I As Integer
Dim Col As Collection
Dim C1 As String

On Error GoTo EComprobarCuentas
'NO hay cuentas
If Me.adodc1.Recordset.EOF Then Exit Sub
'Buscando datos
If txtAux(0).Visible Then Exit Sub
'Para cada nivel n comprobaremos si existe la cuenta en un
'nivel n-1
'La comprobacion se hara para cada cta de n sabiendo k
' para las cuentas de nivel 4 digitos  4300 ..4309 tienen
'el mismo subnivel n-1 430
lblComprobar(0).Caption = ""
lblComprobar(1).Caption = ""
lblComprobar(0).Visible = True
lblComprobar(1).Visible = True
Me.lblIndicador.Caption = "Comprobar cuentas"
Me.Refresh
Errores = 0
NF = FreeFile
Open App.path & "\Errorcta.txt" For Output As #NF


'Primero comprobamos las cuentas de mayor longitud que la permitida
CuentasDeMasNivel
lblComprobar(0).Caption = "Cuentas > Ultimo nivel"
lblComprobar(0).Refresh
Set Col = New Collection
'Hasta 2 pq el uno no tiene subniveles
For I = vEmpresa.numnivel To 2 Step -1
    N = DigitosNivel(I)
    lblComprobar(0).Caption = "Nivel: " & I
    lblComprobar(0).Refresh
    Do
        If ObtenerCuenta(Cad, I, N) Then
            'Frame1.Visible = False
            lblComprobar(1).Caption = Cad
            lblComprobar(1).Refresh
            ComprobarCuenta Cad, I, Col
        End If
    Loop Until Cad = ""
Next I


'Otras comprobaciones de las cuentas
Me.lblComprobar(0).Caption = "Comp. cta numerica o con ' '"
Me.lblComprobar(1).Caption = "Leyendo BD"
Set RS = New ADODB.Recordset
OtrasComprobacionesCuentas
Set RS = Nothing

Close #NF
Me.lblComprobar(0).Caption = "Proceso"
Me.lblComprobar(1).Caption = "Finalizado"

If Errores = 0 Then
    Kill App.path & "\Errorcta.txt"
    MsgBox "Comprobación finalizada", vbInformation
    
    Else
        Cad = Dir("C:\WINDOWS\NOTEPAD.exe")
        If Cad = "" Then
            Cad = Dir("C:\WINNT\NOTEPAD.exe")
        End If
        If Cad = "" Then
            MsgBox "Se ha producido errores. Vea el archivo Errorcta.txt"
            Else
            Shell Cad & " " & App.path & "\Errorcta.txt", vbMaximizedFocus
        End If
        espera 2

        If vUsu.Nivel < 2 Then
            If MsgBox("Desea crear los subniveles?", vbQuestion + vbYesNo) = vbYes Then
                    
                    Cad = "insert into `cuentas` (`codmacta`,`nommacta`,`apudirec`,dirdatos) VALUES ('"
                    For NF = 1 To Col.Count
                        N = DigitosNivelAnterior(Col.Item(NF))
                        If N > 0 Then
                            C1 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Mid(Col.Item(NF), 1, N), "T")
                        Else
                            C1 = ""
                        End If
                        If C1 = "" Then C1 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Col.Item(NF), "T")
                        
                        If C1 = "" Then C1 = "AUTOM: " & Col.Item(NF)

                        EjecutaSQL Cad & Col.Item(NF) & "','" & DevNombreSQL(C1) & "','N','AUTOMATICA en la comprobacion')"
                    Next NF
            End If
        End If
End If

EComprobarCuentas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar cuentas: ", Err.Description
        Close #NF
    End If
    Me.lblComprobar(0).Visible = False
    Me.lblComprobar(1).Visible = False
    Me.lblIndicador.Caption = ""
    Me.Refresh
    Set Col = Nothing
End Sub


Private Function ObtenerCuenta(ByRef CADENA As String, Nivel As Integer, ByRef Digitos As Integer) As Boolean
Dim RT As Recordset
Dim SQL As String


If CADENA = "" Then
    SQL = ""
Else
    SQL = DevuelveUltimaCuentaGrupo(CADENA, Nivel, Digitos)
    SQL = " codmacta > '" & SQL & "' AND "
End If
SQL = "Select codmacta from Cuentas WHERE " & SQL
SQL = SQL & " codmacta like '" & Mid("__________", 1, Digitos) & "'"

Set RT = New ADODB.Recordset
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RT.EOF Then
    ObtenerCuenta = False
    CADENA = ""
Else
    ObtenerCuenta = True
    CADENA = RT!codmacta
End If
RT.Close
Set RT = Nothing
End Function


Private Function DigitosNivelAnterior(Cuenta As String) As Integer
Dim I As Integer
    I = Len(Cuenta)
    Select Case I
    Case vEmpresa.numdigi7
        DigitosNivelAnterior = vEmpresa.numdigi6
        
    Case vEmpresa.numdigi5
        DigitosNivelAnterior = vEmpresa.numdigi4
        
    Case vEmpresa.numdigi4
        DigitosNivelAnterior = vEmpresa.numdigi3
        
    
    Case vEmpresa.numdigi3
        DigitosNivelAnterior = vEmpresa.numdigi2
    
    Case vEmpresa.numdigi2
        DigitosNivelAnterior = vEmpresa.numdigi1
    

    Case Else
        DigitosNivelAnterior = 0
    End Select
    
        
End Function


Private Sub ComprobarCuenta(Cuenta As String, Nivel As Integer, ByRef Cole As Collection)
Dim N As Integer
Dim AUX2 As String

N = DigitosNivel(Nivel - 1)
Aux = Mid(Cuenta, 1, N)
AUX2 = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Aux, "T")
If AUX2 = "" Then
    'Error
   Errores = Errores + 1
   Print #NF, "Nivel: " & Nivel
   Print #NF, "Cuenta: " & Cuenta & "  -> " & Aux & " NO encontrada "
   Print #NF, ""
   Print #NF, ""
   Cole.Add Aux
End If

End Sub



Private Function DevuelveUltimaCuentaGrupo(Cta As String, Nivel As Integer, ByRef Digitos As Integer) As String
Dim Cad As String
Dim N As Integer
N = DigitosNivel(Nivel - 1)
Cad = Mid(Cta, 1, N)
Cad = Cad & "9999999999"
DevuelveUltimaCuentaGrupo = Mid(Cad, 1, Digitos)
End Function


Private Sub CuentasDeMasNivel()
'###MYSQL
Set RS = New ADODB.Recordset
RS.Open "SELECT codmacta FROM cuentas WHERE ((Length(cuentas.codmacta)>" & vEmpresa.DigitosUltimoNivel & "))", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    Print #NF, "Cuentas de longitud mayor a la permitida"
    Print #NF, "Digitos ultimo nivel: " & vEmpresa.DigitosUltimoNivel
    While Not RS.EOF
        Errores = Errores + 1
        Print #NF, "      .- " & RS!codmacta
        RS.MoveNext
    Wend
    Print #NF, ""
    Print #NF, ""
End If
RS.Close
Set RS = Nothing
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    'Caption = KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then Unload Me
    End If
End Sub


Private Sub PonerFoco(ByRef Obje As Object)
On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerResultadosBusquedaAvanzada()

    On Error GoTo EC
        CadenaConsulta = GeneraSQL(CadenaDesdeOtroForm)
        CargaGrid
    Exit Sub
EC:
    MuestraError Err.Number, "Poner resultados busqueda avanzada"
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub



Private Sub VerObservaciones()
    On Error GoTo EVerObservaciones
        
    If adodc1.Recordset Is Nothing Then Exit Sub
    
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Len(adodc1.Recordset!codmacta) < 3 Then Exit Sub
    
    frmMensajes.Opcion = 21
    frmMensajes.Parametros = Mid(adodc1.Recordset!codmacta, 1, 3)
    frmMensajes.Show vbModal
    
    Exit Sub
EVerObservaciones:
    MuestraError Err.Number, "Ver Observaciones"
    
End Sub


Private Sub OtrasComprobacionesCuentas()

    'Busco cuentas que no sean numericas
    Me.Refresh
    Aux = "Select codmacta from cuentas"
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = "Cuentas NO numericas o con espacios en blanco"
    While Not RS.EOF
        Me.lblComprobar(1).Caption = RS!codmacta
        Me.lblComprobar(1).Refresh
         If Not IsNumeric(RS!codmacta) Then
            EscribeEnErrores Aux, RS!codmacta
         Else
            If InStr(1, RS!codmacta, " ") > 0 Then EscribeEnErrores Aux, RS!codmacta
        End If
        RS.MoveNext
    Wend
    RS.Close
End Sub

Private Sub EscribeEnErrores(Titulito As String, Cuenta As String)
    'Error
   Errores = Errores + 1
   If Titulito <> "" Then
        Print #NF, " *****  " & Titulito
        Print #NF,: Print #NF,: Print #NF,
        Titulito = ""
   End If
   Print #NF, " - " & Cuenta


End Sub

Private Sub txtaux_LostFocus(Index As Integer)

    If Index = 0 Then
        Aux = txtAux(0).Text
        If CuentaCorrectaUltimoNivel(Aux, "") Then txtAux(0).Text = Aux
    End If
End Sub
