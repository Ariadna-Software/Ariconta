VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmUtilidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameCLI 
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   4800
      Width           =   5055
      Begin VB.Frame FrameProgresoFac 
         Height          =   735
         Left            =   1080
         TabIndex        =   42
         Top             =   120
         Width           =   4815
         Begin VB.Label Label7 
            Caption         =   "Label7"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   44
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkContabilizacion 
         Caption         =   "Contab."
         Height          =   195
         Left            =   3840
         TabIndex        =   41
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1560
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCLI 
         Height          =   315
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   30
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   32
         Top             =   240
         Width           =   210
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   3
         Left            =   720
         Picture         =   "frmUtilidades.frx":030A
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   2
         Left            =   1920
         Picture         =   "frmUtilidades.frx":0395
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Serie"
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame framBusCta 
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   4920
      Width           =   5055
      Begin VB.CommandButton cmdCrearCuenta 
         Caption         =   "Crear"
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtHuecoCta 
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblHuecoCta 
         Caption         =   "Label5"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame frameBus2 
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   5055
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4035
      End
   End
   Begin VB.Frame frameBusASiento 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Frame FrameAccionesCtas 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar ctas"
         Height          =   435
         Left            =   3060
         TabIndex        =   7
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Quitar selec."
         Height          =   435
         Left            =   1620
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Selec. todo"
         Height          =   435
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdBus 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5280
      TabIndex        =   36
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Can"
      Height          =   375
      Left            =   5280
      TabIndex        =   37
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame frameAgrupar 
      Height          =   1035
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   4875
      Begin VB.CommandButton cmdEliminarAgrup 
         Caption         =   "Eliminar"
         Height          =   435
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevaAgrup 
         Caption         =   "Insertar"
         Height          =   435
         Left            =   1020
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FrameExclusion 
      Height          =   1035
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Width           =   4875
      Begin VB.CommandButton cmdNuevaExclusion 
         Caption         =   "Insertar"
         Height          =   435
         Left            =   1020
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminaExclusion 
         Caption         =   "Eliminar"
         Height          =   435
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FrameDescuadre 
      Caption         =   "Intervalo busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   4995
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   1
         Left            =   3360
         Picture         =   "frmUtilidades.frx":0420
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   0
         Left            =   900
         Picture         =   "frmUtilidades.frx":04AB
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   16
         Top             =   480
         Width           =   210
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame FrameDHCuenta 
      Height          =   975
      Left            =   120
      TabIndex        =   38
      Top             =   4920
      Width           =   4935
      Begin VB.TextBox Text2 
         Height          =   320
         Index           =   1
         Left            =   2880
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   320
         Index           =   0
         Left            =   600
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   40
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   39
         Top             =   240
         Width           =   1155
      End
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   5350
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Cuentas sin movimientos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   4980
      Width           =   4515
   End
End
Attribute VB_Name = "frmUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////
'/*
'/*         Este formulario es para algunos puntos de utilidades.
'/*         Esta a parte pq vamos a poner el boton de parar busqueda
'/*         ya que es una simple busqueda



Public Opcion As Byte
    '0.- Cuentas sin movimiento
    '1.- ASientos descuadrados
    '2.- Agrupacion de cuentas en balance
    '3.- Cuentas as excluir en el consolidado
    '4.- Busqueda huecos cuentas libres
    '5.- Facturas Clientes
    '6.- Facturas proveedores
    '7.- Cuentas libres. Igual que el 4 pero cuando pulse crear, devolvera la cta libre
    
    
    
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
    
Private Estado As Byte
    '0.- Antes de empezar a buscar
    '1.- Buscando
    '2.- Han parado la busqueda
    '3.- Ha terminado la busqueda y hay datos
Dim SQL As String
Dim RS As Recordset
Dim NumCuentas As Long
Dim I As Long
Dim ItmX As ListItem
Dim HanPulsadoCancelar As Boolean
Dim PrimeraVez As Boolean


Dim SePuedeCErrar As Boolean

Private Sub cmdBus_Click()
    SePuedeCErrar = False
    HacerBusqueda
    SePuedeCErrar = True
End Sub

Private Sub HacerBusqueda()

    Select Case Estado
    Case 0
        ListView1.ListItems.Clear
        Select Case Opcion
        Case 0
            NumCuentas = 1
            MontarBusqueda
            If NumCuentas = 0 Then Exit Sub
            'Solo para esto.
            frameBus2.Visible = True
            cmdCancel.Enabled = False
''            QuitarHlinApu 0   ' Hco apuntes
''            QuitarHlinApu 1   ' Hco Apuntes cerrados
''            QuitarHlinApu 2   ' Intoduccion de pauntes
''            QuitarHlinApu 3   ' Facturas clientes
''            QuitarHlinApu 4   ' Facturas proveedores

            For I = 0 To 11
                QuitarHlinApu CByte(I)
            Next I
            
            
            'Si tiene tesoreria comprobar que no esta en tablas de tesoreria
            If vEmpresa.TieneTesoreria Then
                For I = 21 To 25
                    QuitarHlinApu CByte(I)
                Next I
            End If
            
            
            QuitarOtrasCuentas
            
            
            
            RecordsetRestantes
            frameBus2.Visible = False
            PonerCampos 1
            cmdCancel.Enabled = True
            HanPulsadoCancelar = False
            RecorriendoRecordset
            
        Case 4, 7
            If Len(Me.txtHuecoCta.Text) < Me.txtHuecoCta.MaxLength Then
                MsgBox "Escriba el subgrupo completo", vbExclamation
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            cmdCrearCuenta.Visible = False
            CargarRecordSetCtasLibres
            If ListView1.ListItems.Count > 0 Then
                ListView1.ListItems(1).EnsureVisible
                cmdCrearCuenta.Visible = True
            Else
                MsgBox "Ninguna cuenta libre para el subgrupo: " & Me.txtHuecoCta.Text, vbInformation
            End If
            Screen.MousePointer = vbDefault
            PonerCampos 0
            
        Case 5, 6
            Screen.MousePointer = vbHourglass
            
            If Me.chkContabilizacion.Value = 1 Then
                CargaEncabezado 100   '
                Set miRsAux = New ADODB.Recordset
                BuscarContabilizacionFacturas
                Set miRsAux = Nothing
            Else
                
                    
                CargaEncabezado Opcion
                BuscarFacturasSaltos
            End If
            
            Screen.MousePointer = vbDefault
        Case Else
            'Buscar asiento descuadrado
            Me.FrameDescuadre.Visible = False
            MontaSQLBuscaAsien
            PonerCampos 1
            HanPulsadoCancelar = False
            RecorriendoRecordsetDescuadres
        End Select
    
    Case 2
        'Volvemos donde nos habiamos quedado
        PonerCampos 1
        HanPulsadoCancelar = False
        If Opcion = 0 Then
            RecorriendoRecordset
        Else
            RecorriendoRecordsetDescuadres
        End If
    Case 3
        ListView1.ListItems.Clear
        PonerCampos 0
        
        
    Case 4
        'Busqueda cta libre
        
    End Select
End Sub

Private Sub cmdCancel_Click()
    Select Case Estado
    Case 0
        SePuedeCErrar = True
        If Opcion = 7 Then CadenaDesdeOtroForm = ""
        Unload Me
        
    Case 1
        HanPulsadoCancelar = True
        PonerCampos 0
        
    Case 2
        'Volvemos a poner una nueva busqueda
        IntentaCErrar
        PonerCampos 0
        If Opcion = 1 Then Me.FrameDescuadre.Visible = True
        
    Case 3
        SePuedeCErrar = True
        If Opcion = 7 Then CadenaDesdeOtroForm = ""
        Unload Me
        
    End Select
End Sub


Private Sub IntentaCErrar()
On Error Resume Next
    RS.Close
    Err.Clear
    Set RS = Nothing
End Sub


Private Sub cmdCrearCuenta_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione una cuenta", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = ""
    If Opcion = 7 Then
        SePuedeCErrar = True
        CadenaDesdeOtroForm = ListView1.SelectedItem.Text
        Unload Me
    Else
        frmCuentas.CodCta = ListView1.SelectedItem.Text
        frmCuentas.vModo = 1
        frmCuentas.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            'Eliminamos el nodo
            If ListView1.SelectedItem.Text = CadenaDesdeOtroForm Then ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    End If
End Sub

Private Sub cmdEliminaExclusion_Click()
    EliminarCta
End Sub

Private Sub cmdEliminarAgrup_Click()
    EliminarCta
End Sub


Private Sub EliminarCta()
On Error Resume Next
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    
    SQL = "Va a eliminar de "
    If Opcion = 2 Then
        SQL = SQL & "  la agrupacion de cuentas en balance la cuenta: " & vbCrLf
    Else
        SQL = SQL & "  la exclusion de cuentas en consolidado: " & vbCrLf
    End If
    SQL = SQL & ListView1.SelectedItem.Text & " - " & ListView1.SelectedItem.SubItems(1) & vbCrLf
    SQL = SQL & "Desea continuar ?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    Screen.MousePointer = vbHourglass
    SQL = "DELETE FROM "
    If Opcion = 3 Then
        SQL = SQL & "ctaexclusion"
    Else
        SQL = SQL & "ctaagrupadas"
    End If
    SQL = SQL & " WHERE codmacta = '" & ListView1.SelectedItem.Text & "';"
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar cuenta de agrupacion"
    Else
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdNuevaAgrup_Click()
    Set frmC = New frmColCtas
    frmC.ConfigurarBalances = 2
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.Show vbModal
End Sub

Private Sub cmdNuevaExclusion_Click()
    Set frmC = New frmColCtas
    frmC.ConfigurarBalances = 6
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.Show vbModal
End Sub

Private Sub Command1_Click()
    Checkear True
End Sub

Private Sub Checkear(SiNO As Boolean)
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = SiNO
    Next I
End Sub


Private Sub Command2_Click()
    Checkear False
End Sub

Private Sub Command3_Click()
    SQL = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            SQL = "SI"
            Exit For
        End If
    Next I
    If SQL = "" Then
        MsgBox "Seleccione alguna cuenta a eliminar", vbExclamation
        Exit Sub
    End If
    SQL = "Va a eliminar las cuentas seleccionadas. ¿ Esta seguro ?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        Screen.MousePointer = vbHourglass
        Eliminar
        Screen.MousePointer = vbDefault
    End If
            
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            If Not BloqueoManual(True, "Busquedas", "1") Then
                MsgBox "Se esta realizando la busqueda desde otro PC", vbExclamation
                PrimeraVez = True
                SePuedeCErrar = True
                Unload Me
            End If
            PonFocus Text2(0)
        Case 1
            Text1(0).SetFocus
        Case 4, 7
            txtHuecoCta.SetFocus
        Case 5, 6
            Text1(3).SetFocus
            
        End Select
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    SePuedeCErrar = False
    Me.ListView1.Icons = frmPpal.ImageList1
    Me.ListView1.SmallIcons = frmPpal.ImageList1
    CargaEncabezado Opcion
    PonerCampos 0
    Limpiar Me
    Me.FrameDescuadre.Visible = False
    Me.FrameAccionesCtas.Visible = False
    frameBusASiento.Visible = False
    frameBus2.Visible = False
    FrameExclusion.Visible = False
    Me.frameAgrupar.Visible = False
    Me.framBusCta.Visible = False
    Me.frameCLI.Visible = False
    Me.FrameDHCuenta.Visible = False
    Select Case Opcion
    Case 0
        FrameDHCuenta.Visible = True
        Label1.Caption = "Cuentas sin movimientos"
    Case 1
        Label1.Caption = "Asientos descuadrados"
        Me.FrameDescuadre.Visible = True
        'ofertamos fechas
        Text1(0).Text = vParam.fechaini
        Text1(1).Text = vParam.fechafin
    Case 2
        Label1.Caption = "Agrupación de cuentas en balance"
        frameAgrupar.Visible = True
        cargaAgrupacion "ctaagrupadas"
    Case 3
        FrameExclusion.Visible = True
        Label1.Caption = "Exclusion ctas. consolidado empresas"
        cargaAgrupacion "ctaexclusion"
    Case 4, 7
        framBusCta.Visible = True
        Label1.Caption = "Búsquedas número de cuentas libres"
        txtHuecoCta.Text = ""
        cmdCrearCuenta.Visible = False
        PonerDigitosPenultimoNivel
        Me.cmdCrearCuenta.Enabled = vUsu.Nivel < 3
    Case 5, 6
        'Facturas clienes  y Facturas proveedores
        FrameProgresoFac.Left = 120
        FrameProgresoFac.Visible = False
        Me.frameCLI.Visible = True
        Label5.Visible = Opcion = 5
        txtCLI.Visible = Opcion = 5
        Text1(3).Text = vParam.fechaini
        Text1(2).Text = vParam.fechafin
        If Opcion = 5 Then
            Label1.Caption = "Nº facturas CLIENTE incorrectos"
        Else
            Label1.Caption = "Nº facturas PROVEEDORES incorrectos"
        End If
    End Select
    
    'No puede eliminar cuentas
    Command3.Enabled = vUsu.Nivel < 2
    Me.cmdEliminarAgrup.Enabled = vUsu.Nivel < 2
    Me.cmdNuevaAgrup.Enabled = Me.cmdEliminarAgrup.Enabled
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos(NuevoEstado As Byte)


    Select Case NuevoEstado
    Case 0
        Me.Label2.Caption = ""
        Me.pb1.Visible = False
        Me.cmdCancel.Caption = "Salir"
        Me.cmdBus.Caption = "Iniciar"
    Case 1
        Me.cmdCancel.Caption = "Parar"
        
    Case 2
        Me.cmdCancel.Caption = "Cancelar"
        Me.cmdBus.Caption = "Reanudar"
    Case 3
        Me.cmdBus.Caption = "Nueva busq."
        Me.cmdCancel.Caption = "Salir"
    End Select
    If Opcion = 0 Then
        Me.FrameAccionesCtas.Visible = NuevoEstado = 3
    Else
        Me.frameBusASiento.Visible = NuevoEstado = 3
    End If
    Me.cmdBus.Enabled = (NuevoEstado <> 1)
    cmdBus.Visible = (Opcion < 2) Or Opcion >= 4 'Cuando es agrupacion no mostramos el inciar
    Estado = NuevoEstado
End Sub


Private Sub CargaEncabezado(LaOpcion As Byte)
Dim clmX As ColumnHeader

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    
    Select Case LaOpcion
    Case 0, 2, 3
        Me.ListView1.Checkboxes = LaOpcion = 0
        '* Estamos en cuentas sin movimiento
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Cuenta"
        clmX.Width = 1500
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Título"
        clmX.Width = 4500
    Case 1
        Me.ListView1.Checkboxes = False
        '* Estamos en cuentas sin movimiento
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Asiento"
        clmX.Width = 2000
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Fecha"
        clmX.Width = 1300
        'Diario
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Diario"
        clmX.Width = 800
    
        'Debe
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Descuadre"
        clmX.Width = 2000
        clmX.Alignment = lvwColumnRight
    Case 4, 7
        Me.ListView1.Checkboxes = False
        '* Estamos en buscando huecos cuentas
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Cuenta"
        clmX.Width = 1500
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Comentario"
        clmX.Width = 4500
    Case 5, 6
        
        Me.ListView1.Checkboxes = False
        '* Facturas
        'Cuenta
        If LaOpcion = 5 Then
            Set clmX = ListView1.ColumnHeaders.Add()
            clmX.Text = "Serie"
            clmX.Width = 600
            I = 3900
            SQL = "Codigo"
        Else
            I = 4500
            SQL = "Registro"
        End If
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = SQL
        clmX.Width = 1500
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Año"
        clmX.Width = 800
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Comentario"
        clmX.Width = I
       
       
    Case 100
        'Esta opcion es para las facturas, la busqueda de las contbilizaciones
        Me.ListView1.Checkboxes = False
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Factura"
        clmX.Width = 2500
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Asiento"
        clmX.Width = 2500
        
        
    End Select
End Sub


Private Sub MontarBusqueda()
    SQL = "DELETE FROM tmpbussinmov"
    Conn.Execute SQL
    SQL = "INSERT INTO tmpbussinmov SELECT codmacta,nommacta from cuentas where apudirec='S'"
    If Text2(0).Text <> "" Then SQL = SQL & " AND codmacta >= '" & Text2(0).Text & "'"
    If Text2(1).Text <> "" Then SQL = SQL & " AND codmacta <= '" & Text2(1).Text & "'"
    Conn.Execute SQL
    
    
    If Text2(0).Text <> "" Or Text2(1).Text <> "" Then
        RecordsetRestantes
        If NumCuentas > 0 Then
            RS.Close
            Set RS = Nothing
        Else
            MsgBox "Ningun dato seleccionado", vbExclamation
        End If
    End If
    
End Sub



Private Sub QuitarHlinApu(vOpcion As Byte)
Dim T As Long
Dim t2 As Long
Dim codmacta1 As String
    'Opcion
    ' 0 .- Halinapu
    ' 1 .- hlinapu1
    ' 2 .- linapu
   
    codmacta1 = "codmacta"
    Select Case vOpcion
    Case 0, 1
        SQL = "hlinapu"
        If vOpcion = 1 Then SQL = SQL & "1"
    Case 2
        SQL = "linapu"
        
    Case 3, 4
        SQL = "cabfact"
        If vOpcion = 4 Then SQL = SQL & "prov"
    Case 5, 6
        SQL = "hsaldosanal"
        If vOpcion = 6 Then SQL = SQL & "1"
    Case 7, 8
        'hsaldos
        SQL = "hsaldos"
        If vOpcion = 8 Then SQL = SQL & "1"
        
    Case 9, 10
        'Contrapartida en hlinapu, y hlinapu1
        codmacta1 = "ctacontr"
        SQL = "hlinapu"
        If vOpcion = 10 Then SQL = SQL & "1"
        
        
        
    Case 11
        'PResupuestaria
        SQL = "presupuestos"
        
        
    '-----------------------------
    'TESORERIA
    Case 21
        SQL = "slicaja"
    
    Case 22
        codmacta1 = "ctacaja"
        SQL = "susucaja"
        
    Case 23
        SQL = "Departamentos"
        
    Case 24
        SQL = "scobro"
        
    Case 25
        SQL = "spagop"
        codmacta1 = "ctaprove"
    
    End Select
    Label4.Tag = SQL
    Label4.Caption = "buscando datos " & SQL
    pb2.Value = 0
    Me.Refresh
    
    SQL = "Select " & codmacta1 & " from " & SQL
    
    'Si es de hsaldos entonces tenemos k buscar solo en las k sean de ultmo nivel
    If vOpcion = 8 Or vOpcion = 7 Then _
        SQL = SQL & " WHERE codmacta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
    
    SQL = SQL & " group by " & codmacta1
    
    'having
    SQL = SQL & " HAVING NOT (" & codmacta1 & " IS NULL)"
    
    Set RS = New ADODB.Recordset
    'Primro el contador
    RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    T = 0
    While Not RS.EOF
        T = T + 1
        RS.MoveNext
    Wend
    
    If T > 0 Then
        RS.MoveFirst
        Label4.Caption = Label4.Tag
        Label4.Refresh
        t2 = 0
        T = T + 1
        While Not RS.EOF
            t2 = t2 + 1
            pb2.Value = ((t2 / T) * 1000)
            SQL = "Delete from tmpbussinmov where codmacta ='" & RS.Fields(0) & "';"
            Conn.Execute SQL
            RS.MoveNext
        Wend
    End If
    RS.Close
    Label4.Caption = ""
    Set RS = Nothing
End Sub


Private Sub RecordsetRestantes()
    Set RS = New ADODB.Recordset
    SQL = "Select count(*) from tmpbussinmov"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumCuentas = 0
    I = 0
    If Not RS.EOF Then
        NumCuentas = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    If NumCuentas = 0 Then Exit Sub
    pb1.Visible = True
    Label2.Caption = ""
    pb1.Value = 0
    Me.Refresh
    SQL = "Select * from tmpbussinmov order by codmacta"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Sub



Private Sub RecorriendoRecordset()
Dim ExisteReferencia As Boolean

    If NumCuentas = 0 Then Exit Sub
    While Not RS.EOF
        ExisteReferencia = False
        
        Label2.Caption = RS.Fields(0) & " - " & RS.Fields(1)
        Label2.Refresh
        I = I + 1
        
        pb1.Value = Int(((I / NumCuentas)) * 1000)
        
        'Comprobamos en Facturas
        SQL = DevuelveDesdeBD("codmacta", "cabfact", "codmacta", RS.Fields(0), "T")
        If SQL <> "" Then ExisteReferencia = True
        
        'Proveedores
        If Not ExisteReferencia Then
            SQL = DevuelveDesdeBD("codmacta", "cabfactprov", "codmacta", RS.Fields(0), "T")
            If SQL <> "" Then ExisteReferencia = True
        End If
        
        'Lineas de facturas
        If Not ExisteReferencia Then
            SQL = DevuelveDesdeBD("codtbase", "linfact", "codtbase", RS.Fields(0), "T")
            If SQL <> "" Then ExisteReferencia = True
        End If
        
        'Lineas de facturas proveedores
        If Not ExisteReferencia Then
            SQL = DevuelveDesdeBD("codtbase", "linfactprov", "codtbase", RS.Fields(0), "T")
            If SQL <> "" Then ExisteReferencia = True
        End If
        
        
        'AMortizacion
        '-------------
        'Proveedor
        If Not ExisteReferencia Then
            SQL = DevuelveDesdeBD("codprove", "sinmov", "codprove", RS.Fields(0), "T")
            If SQL <> "" Then ExisteReferencia = True
        End If
        If Not ExisteReferencia Then
            SQL = DevuelveDesdeBD("codmact1", "sinmov", "codmact1", RS.Fields(0), "T")
            If SQL <> "" Then ExisteReferencia = True
        End If
        If Not ExisteReferencia Then
            SQL = DevuelveDesdeBD("codmact2", "sinmov", "codmact2", RS.Fields(0), "T")
            If SQL <> "" Then ExisteReferencia = True
        End If
        If Not ExisteReferencia Then
            SQL = DevuelveDesdeBD("codmact3", "sinmov", "codmact3", RS.Fields(0), "T")
            If SQL <> "" Then ExisteReferencia = True
        End If
        
        
        
        If Not ExisteReferencia Then
           Set ItmX = ListView1.ListItems.Add(, , RS.Fields(0))
           ItmX.SmallIcon = 1
           ItmX.SubItems(1) = RS.Fields(1)
           ItmX.EnsureVisible
        Else
            Conn.Execute "Delete from tmpbussinmov where codmacta='" & RS.Fields(0) & "'"
        End If
        
        'Siguiente
        RS.MoveNext
        'Miramos si hay algo por hacer
        DoEvents
        
        'Si han pulsado parar
        If HanPulsadoCancelar Then
            PonerCampos 2
            Exit Sub
        End If
    Wend
    RS.Close
    
    If ListView1.ListItems.Count > 0 Then
        PonerCampos 3
    Else
        Label2.Caption = ""
        Label2.Refresh
        MsgBox "Ninguna cuenta sin movimientos", vbExclamation
        PonerCampos 0
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SePuedeCErrar Then
        Cancel = 1
    Else
        If Not PrimeraVez Then
            If Opcion = 0 Then BloqueoManual False, "Busquedas", ""
            IntentaCErrar
        End If
    End If
            
End Sub


Private Sub Eliminar()
Dim Cad As String
    SQL = "DELETE FROM cuentas where codmacta = '"
    For I = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(I).Checked Then
            Cad = BorrarCuenta(ListView1.ListItems(I).Text, Me.Label2)
            If Cad = "" Then
                If EliminaCuenta(ListView1.ListItems(I).Text) Then ListView1.ListItems.Remove I
            Else
                Cad = ListView1.ListItems(I).Text & " - " & ListView1.ListItems(I).SubItems(1) & vbCrLf & Cad & vbCrLf
                MsgBox Cad, vbExclamation
            End If
        End If
    Next I
End Sub


Private Function EliminaCuenta(ByRef Cuenta As String) As Boolean
    On Error Resume Next
    Conn.Execute SQL & Cuenta & "'"
    If Err.Number <> 0 Then
        MuestraError Err.Number, Cuenta
        EliminaCuenta = False
    Else
        EliminaCuenta = True
    End If
End Function

Private Sub MontaSQLBuscaAsien()
    Set RS = New ADODB.Recordset
    
    SQL = ""
    'Fecha inicio
    If Text1(0).Text <> "" Then SQL = " fechaent >= '" & Format(Text1(0).Text, FormatoFecha) & "'"
    'Fecha fin
    If Text1(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " fechaent <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    End If
    If SQL <> "" Then SQL = " WHERE " & SQL
    
    SQL = "Select numasien,numdiari,fechaent from hlinapu " & SQL
    SQL = SQL & " group by numasien,numdiari,fechaent"
    RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    NumCuentas = 0
    I = 0
    While Not RS.EOF
        I = I + 1
        RS.MoveNext
    Wend
    NumCuentas = I
    I = 0
    If NumCuentas = 0 Then
        RS.Close
        Exit Sub
    End If
    RS.MoveFirst
    pb1.Visible = True
    Label2.Caption = ""
    pb1.Value = 0
    Me.Refresh

End Sub

Private Sub RecorriendoRecordsetDescuadres()


    If NumCuentas = 0 Then Exit Sub
    While Not RS.EOF
        
        Label2.Caption = RS.Fields(0) & " - " & RS.Fields(2)
        Label2.Refresh
        
        pb1.Value = Int(((I / NumCuentas)) * 1000)
        I = I + 1
        
        ObtenerSumas
        
        'Siguiente
        RS.MoveNext
        'Miramos si hay algo por hacer
        DoEvents
        
        'Si han pulsado parar
        If HanPulsadoCancelar Then
            PonerCampos 2
            Exit Sub
        End If
    Wend
    RS.Close
    
    If ListView1.ListItems.Count > 0 Then
        PonerCampos 3
    Else
        MsgBox "Ningun asiento descuadrado.", vbExclamation
        PonerCampos 0
    End If

End Sub


Private Function ObtenerSumas() As Boolean
    Dim Deb As Currency
    Dim hab As Currency
    Dim RsA As ADODB.Recordset

    Set RsA = New ADODB.Recordset
    'Abril 2004. objetivo Quitar GROUP BY
'    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
'    SQL = SQL & " ,numdiari,fechaent,numasien"
'    SQL = SQL & " From hlinapu GROUP BY numdiari, fechaent, numasien "
'    SQL = SQL & " HAVING (((numdiari)=" & RS!NumDiari
'    SQL = SQL & ") AND ((fechaent)='" & Format(RS!fechaent, FormatoFecha)
'    SQL = SQL & "') AND ((numasien)=" & RS!NumAsien
'    SQL = SQL & "));"
    
    
    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    SQL = SQL & " From hlinapu "
    SQL = SQL & " WHERE (((numdiari)=" & RS!NumDiari
    SQL = SQL & ") AND ((fechaent)='" & Format(RS!fechaent, FormatoFecha)
    SQL = SQL & "') AND ((numasien)=" & RS!Numasien
    SQL = SQL & "));"
    RsA.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RsA.EOF Then
        If IsNull(RsA.Fields(0)) Then
            Deb = 0
        Else
            Deb = RsA.Fields(0)
        End If
        
        'Deb = Round(Deb, 2)
        If IsNull(RsA.Fields(1)) Then
            hab = 0
        Else
            hab = RsA.Fields(1)
        End If
        
        
        
    Else
        Deb = 0
        hab = 0
    End If
    RsA.Close
    
    'Metemos en DEB el total
    Deb = Deb - hab
    If Deb <> 0 Then
            SQL = Format(RS!Numasien, "0000000")
            SQL = "    " & SQL
            Set ItmX = ListView1.ListItems.Add(, , SQL)
            ItmX.SmallIcon = 2
            ItmX.Icon = 2
            ItmX.SubItems(1) = Format(RS!fechaent, "dd/mm/yyyy")
            ItmX.SubItems(2) = RS!NumDiari
            ItmX.SubItems(3) = Format(Deb, FormatoImporte)
    End If
End Function


Private Sub cargaAgrupacion(Tabla As String)
    On Error GoTo E1
    Set RS = New ADODB.Recordset
    SQL = "Select " & Tabla & ".codmacta, nommacta from " & Tabla & ",cuentas where "
    SQL = SQL & Tabla & ".codmacta=cuentas.codmacta order by " & Tabla & ".codmacta"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
            SQL = RS!codmacta
            Set ItmX = ListView1.ListItems.Add(, , SQL)
            ItmX.SmallIcon = 2
            ItmX.Icon = 2
            ItmX.SubItems(1) = RS!nommacta
            RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Exit Sub
E1:
    MuestraError Err.Number, Tabla
    Set RS = Nothing
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
On Error Resume Next
    If Opcion = 3 Then
        SQL = "ctaexclusion"
    Else
        SQL = "ctaagrupadas"
    End If
    SQL = "INSERT INTO " & SQL & "(codmacta) VALUES ('" & RecuperaValor(CadenaSeleccion, 1) & "')"
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando la cuenta"
    Else
        ListView1.ListItems.Clear
        If Opcion = 3 Then
            cargaAgrupacion "ctaexclusion"
        Else
            cargaAgrupacion "ctaagrupadas"
        End If
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(I).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imfech_Click(Index As Integer)
    I = Index
    Set frmF = New frmCal
    SQL = Now
    If Text1(I).Text <> "" Then
        If IsDate(Text1(I).Text) Then SQL = Text1(I).Text
    End If
    frmF.Fecha = CDate(SQL)
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
        Text1(Index).Text = ""
    End If
End Sub

Private Sub PonerDigitosPenultimoNivel()
    'Veremos cual es el ultimo nivel
    I = vEmpresa.numnivel
    If I < 2 Then
        MsgBox "Empresa mal configurada", vbExclamation
        Exit Sub
    End If
    NumCuentas = I - 1
    I = DigitosNivel(CInt(NumCuentas))
    lblHuecoCta.Caption = "Digitos del nivel " & NumCuentas & ":    " & I
    lblHuecoCta.Tag = I
    Me.txtHuecoCta.MaxLength = I
End Sub






Private Sub Text2_GotFocus(Index As Integer)
    PonFoco Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    SQL = Text2(Index).Text
    If CuentaCorrectaUltimoNivelSIN(SQL, CadenaDesdeOtroForm) < 1 Then
        
        MsgBox CadenaDesdeOtroForm, vbExclamation
        Text2(Index).Text = ""
        PonFocus Text2(Index)
    Else
        Text2(Index).Text = SQL
    End If
    CadenaDesdeOtroForm = ""
End Sub

Private Sub txtCLI_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtHuecoCta_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtHuecoCta_LostFocus()
    txtHuecoCta.Text = Trim(txtHuecoCta.Text)
    If txtHuecoCta.Text <> "" Then
        If Not IsNumeric(txtHuecoCta.Text) Then
            MsgBox "Campo numérico", vbExclamation
            Exit Sub
        End If
        txtHuecoCta.Text = Val(txtHuecoCta.Text)
    End If
End Sub


Private Sub CargarRecordSetCtasLibres()
Dim Cad As String
Dim J As Long
Dim Multiplicador As Long
Dim vFormato As String

    I = vEmpresa.DigitosUltimoNivel - lblHuecoCta.Tag
    vFormato = Mid("00000000000", 1, I)
    Multiplicador = I
    Cad = Me.txtHuecoCta.Text & Mid("0000000000", 1, I)
    I = 1   'Primer Numero de cuenta
    
    Set RS = New ADODB.Recordset
    SQL = "DELETE FROM tmpbussinmov"
    Conn.Execute SQL
    
    
    
    SQL = "Select codmacta from cuentas where codmacta like '" & Me.txtHuecoCta.Text & "%' AND Apudirec='S'"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "INSERT INTO tmpbussinmov VALUES ('"
    If RS.EOF Then
        'Estan todas libres
        
        InsertaCtasLibres Format(I, vFormato), "TODAS LIBRES"
        RS.Close
    Else
        
        While Not RS.EOF
            NumCuentas = CLng(Right(CStr(RS.Fields(0)), Multiplicador))
            If NumCuentas > I Then
                For J = I To NumCuentas - 1
                    InsertaCtasLibres Format(J, vFormato), "SALTO"
                Next J
            End If
            I = NumCuentas + 1
            RS.MoveNext
        Wend
        RS.Close
        'Cojemos desde la ultima
        I = vEmpresa.DigitosUltimoNivel - lblHuecoCta.Tag
        Cad = Mid("999999999", 1, I)
        I = Val(Cad) 'Utlima cta del subgrupo
        
        If NumCuentas < I Then
            NumCuentas = NumCuentas + 1
            InsertaCtasLibres Format(NumCuentas, vFormato), "Desde aqui LIBRES"
        End If
        
        
    End If
    
        SQL = "Select * from tmpbussinmov ORDER BY codmacta"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add(, , RS.Fields(0))
           ItmX.SmallIcon = 1
           ItmX.SubItems(1) = RS.Fields(1)
           ItmX.EnsureVisible
            RS.MoveNext
        Wend
        RS.Close
End Sub


Private Sub InsertaCtasLibres(Cta As String, Descripcion As String)
Dim Cad As String
        Cad = Me.txtHuecoCta.Text & Cta
        Cad = Cad & "','" & Descripcion & "')"
        Conn.Execute SQL & Cad
End Sub



Private Sub BuscarContabilizacionFacturas()
    
    cmdCancel.Enabled = False
    cmdBus.Enabled = False
    
    ListView1.ListItems.Clear
    NumCuentas = 0
    Set RS = New ADODB.Recordset
    
    Label7(0).Caption = ""
    Label7(1).Caption = ""
    Me.FrameProgresoFac.Visible = True
    
    'Comprobamos facturas que estando contabilizadas no tienen apuntes
    FacturasContabilizadas
    
    'Apuntes que siendo de factura, no esta la factura
    ApuntesSinFactura
    
    If NumCuentas = 0 Then MsgBox "Proceso finalizado", vbInformation
    
    CadenaDesdeOtroForm = ""
    
EBuscarContabilizacionFacturas:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
    Set miRsAux = Nothing
    Me.FrameProgresoFac.Visible = False
    cmdCancel.Enabled = True
    cmdBus.Enabled = True
    
End Sub


Private Sub BuscarFacturasSaltos()
Dim Cad As String
Dim Aux As String
Dim Anyo As Integer
Dim J As Integer


    On Error GoTo EBuscarFacturas
        
    
    If Opcion = 5 Then
        SQL = "numserie,anofaccl as ano,codfaccl as codigo"
        Cad = "fecfaccl"
    Else
        SQL = "anofacpr as ano,numregis as codigo"
        Cad = "fecrecpr"
    End If
    
    SQL = SQL & " FROM cabfact"
    If Opcion = 6 Then SQL = SQL & "prov"
    
    
    'Si hay fecha inicio
    If Text1(3).Text = "" Or Text1(2).Text = "" Then
        MsgBox "Debe escribir las fechas de inicio y fin", vbExclamation
        Exit Sub
    End If
    Aux = ""
    Aux = Cad & " >= '" & Format(Text1(3).Text, FormatoFecha) & "'"
    
    
    Aux = Aux & " AND "
    Aux = Aux & Cad & " <= '" & Format(Text1(2).Text, FormatoFecha) & "'"
    
    
    
    If txtCLI.Text <> "" Then
        If Opcion = 5 Then
            If Aux <> "" Then Aux = Aux & " AND "
            Aux = Aux & " numserie = '" & txtCLI.Text & "'"
        End If
    End If
    If Aux <> "" Then SQL = SQL & " WHERE " & Aux
    SQL = SQL & " ORDER BY "
    If Opcion = 5 Then
        SQL = SQL & "numserie,anofaccl ,codfaccl "
    Else
        SQL = SQL & "anofacpr,numregis"
    End If
    Set RS = New ADODB.Recordset
    
    
    '#FALTA revisar esto
    
    'Obtenego el minimo
    
    Set miRsAux = New ADODB.Recordset
    
    'Fale. Ya tenemos montado el SQL
    
    RS.Open "SELECT " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Serie
    Aux = ""
    Anyo = 0
    While Not RS.EOF
        If Opcion = 5 Then
            If RS!NUmSerie <> Aux Then
                'Nueva SERIE
                Aux = RS!NUmSerie
                Anyo = RS!Ano
                I = FacturaMinimo(Aux, CDate(Text1(3).Text), CDate(Text1(2).Text), Anyo)
            End If
        End If
        If Anyo <> RS!Ano Then
            'AÑO DISTINTO
            Anyo = RS!Ano
            I = FacturaMinimo(Aux, CDate(Text1(3).Text), CDate(Text1(2).Text), Anyo)
        End If
        
        'Para cada numero de factura
        If I = RS!Codigo Then
            I = I + 1
            'no hacemos nada mas
        Else
            'Si si que es mayor. Hay salto o hueco
            

            
            
            If RS!Codigo - I >= 2 Then
                'SALTO
                Cad = Format(RS!Codigo - 1, "000000000")
                If Opcion = 5 Then
                    Set ItmX = ListView1.ListItems.Add(, , RS!NUmSerie)
                    ItmX.SubItems(1) = Cad
                    J = 2
                Else
                    Set ItmX = ListView1.ListItems.Add(, , Cad)
                    J = 1
                End If
                ItmX.SubItems(J) = Anyo
                ItmX.SubItems(J + 1) = "Salto desde codigo: " & Format(I, "00000000")
                
                    
                
            Else
                'HUECO
                Cad = Format(I, "000000000")
                If Opcion = 5 Then
                    Set ItmX = ListView1.ListItems.Add(, , RS!NUmSerie)
                    ItmX.SubItems(1) = Cad
                    J = 2
                Else
                    Set ItmX = ListView1.ListItems.Add(, , Cad)
                    J = 1
                End If
                ItmX.SubItems(J) = Anyo
                ItmX.SubItems(J + 1) = "Falta"
                'i = RS!Codigo + 1
            End If
            ItmX.SmallIcon = 1
             I = RS!Codigo + 1
        End If
        'Movemos siguiente
        RS.MoveNext
        
    Wend
    RS.Close
    
    If ListView1.ListItems.Count = 0 Then MsgBox "Proceso finalizado", vbInformation
    
    
EBuscarFacturas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing
End Sub


Private Function FacturaMinimo(Serie As String, FIni As Date, fFin As Date, Anyo As Integer) As Long
Dim C As String
Dim Campo As String
Dim F1 As Date

    If Opcion = 5 Then
        C = "Select min(codfaccl) FROM cabfact WHERE "
        Campo = "fecfaccl"
    Else
        C = "Select min(numregis) FROM cabfactprov WHERE "
        Campo = "fecrecpr"
    End If
    
    'FEHAS   INICO
    If Anyo = Year(FIni) Then
        F1 = FIni
    Else
        F1 = CDate("01/01/" & Anyo)
    End If
    C = C & Campo & " >= '" & Format(F1, FormatoFecha) & "'"
    
    If Anyo = Year(fFin) Then
        F1 = fFin
    Else
        F1 = CDate("31/12/" & Anyo)
    End If
    C = C & " AND " & Campo & " <= '" & Format(F1, FormatoFecha) & "'"
    
    If Opcion = 5 Then C = C & " AND numserie = '" & Serie & "'"
    'Debug.Print C
    FacturaMinimo = 0

    miRsAux.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then FacturaMinimo = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
End Function

Private Sub PonFocus(ByRef Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub FacturasContabilizadas()
    Label7(0).Caption = "Facturas"
    Label7(1).Caption = "Obteniendo registros"
    Me.Refresh
    'Cogemos las facturas en RS
    SQL = "Select numasien,numdiari,fechaent, "
    
    If Opcion = 5 Then
        SQL = SQL & " numserie,codfaccl c, fecfaccl f"
    Else
        SQL = SQL & " numregis c, fecrecpr f"
    End If
    SQL = SQL & " FROM cabfact"
    If Opcion = 6 Then SQL = SQL & "prov"
    SQL = SQL & " WHERE numasien>0 "
    If Opcion = 5 Then
        CadenaDesdeOtroForm = "fecfaccl"
        If txtCLI.Text <> "" Then SQL = SQL & " AND numserie ='" & txtCLI.Text & "'"
    Else
        CadenaDesdeOtroForm = "fecrecpr"
    End If
    
    If Text1(3).Text <> "" Then SQL = SQL & " AND " & CadenaDesdeOtroForm & " >='" & Format(Text1(3).Text, FormatoFecha) & "'"
    If Text1(2).Text <> "" Then SQL = SQL & " AND " & CadenaDesdeOtroForm & " <='" & Format(Text1(2).Text, FormatoFecha) & "'"
    
    SQL = SQL & " ORDER BY numdiari,numasien,fechaent"
    
    
    'Cuento el recordset
    NumRegElim = 0
    RS.Open "SELECT count(*) " & Mid(SQL, InStr(1, SQL, " FROM ")), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then NumRegElim = DBLet(RS.Fields(0), "N")
    RS.Close
    espera 0.2
    If NumRegElim = 0 Then Exit Sub
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'Hay facturas. Ahora en rsaux cargare los apuntes
        SQL = "Select numasien,numdiari,fechaent from hlinapu WHERE idcontab = 'FRA"
        If Opcion = 5 Then
            SQL = SQL & "CLI'"
        Else
            SQL = SQL & "PRO'"
        End If
        
        If Text1(3).Text <> "" Then SQL = SQL & " AND fechaent >='" & Format(Text1(3).Text, FormatoFecha) & "'"
        If Text1(2).Text <> "" Then SQL = SQL & " AND fechaent <='" & Format(Text1(2).Text, FormatoFecha) & "'"
        SQL = SQL & " GROUP BY numasien,numdiari,fechaent"
        
        miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        

    
        'Recorremos las facturas
        I = 0
        While Not RS.EOF
            'Label7(1).Caption = RS!C & " - " & RS!F
            Label7(1).Caption = I & " de " & NumRegElim
            Label7(1).Refresh
            If Not EstaEnMirsaux Then
                InsertaItemsFacturasContabilizadas True
            End If
            RS.MoveNext
            I = I + 1
            If (I Mod 50) = 0 Then
                Me.Refresh
                DoEvents
            End If
        Wend
        miRsAux.Close
    End If 'Rs.eof
    RS.Close
    
End Sub




Private Function EstaEnMirsaux() As Boolean
Dim Fin As Boolean
    EstaEnMirsaux = False
    miRsAux.MoveFirst
    If miRsAux.EOF Then Exit Function
    Fin = False
    While Not Fin
        If miRsAux!NumDiari = RS!NumDiari Then
              If miRsAux!Numasien = RS!Numasien Then
                If miRsAux!fechaent = RS!fechaent Then
                    Fin = True
                    EstaEnMirsaux = True
                End If
            End If
        End If
        miRsAux.MoveNext
        If miRsAux.EOF Then Fin = True
    Wend
    
End Function


Private Sub InsertaItemsFacturasContabilizadas(RegistroFactura As Boolean)
    NumCuentas = NumCuentas + 1
    Set ItmX = ListView1.ListItems.Add(, "C" & NumCuentas)
    If RegistroFactura Then
        SQL = "  " & Format(RS!C, "00000000") & "   " & Format(RS!f, "dd/mm/yyyy")
        If Opcion = 5 Then SQL = RS!NUmSerie & SQL
        ItmX.Text = SQL
        ItmX.SubItems(1) = " **** "
    Else
        ItmX.Text = " **** "
        ItmX.SubItems(1) = RS!NumDiari & "  " & Format(RS!Numasien, "0000000") & " " & Format(RS!fechaent, "dd/mm/yyyy")
    End If
End Sub



Private Sub ApuntesSinFactura()
    Label7(0).Caption = "Asientos"
    Label7(1).Caption = "Obteniendo registros"
    Me.Refresh


    SQL = "Select numasien,numdiari,fechaent FROM hlinapu WHERE idcontab='FRA"
    
    If Opcion = 5 Then
        SQL = SQL & "CLI'"
    Else
        SQL = SQL & "PRO'"
    End If
    If Text1(3).Text <> "" Then SQL = SQL & " AND fechaent >='" & Format(Text1(3).Text, FormatoFecha) & "'"
    If Text1(2).Text <> "" Then SQL = SQL & " AND fechaent <='" & Format(Text1(2).Text, FormatoFecha) & "'"
        
    SQL = SQL & " GROUP BY numasien,numdiari,fechaent"
    SQL = SQL & " ORDER BY numdiari,numasien,fechaent"
    
    
    'Cuento el recordset
    NumRegElim = 0
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        RS.MoveNext
    Wend
    RS.Close
    espera 0.2
    If NumRegElim = 0 Then Exit Sub
    
    

    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'Hay apuntes. Busco sus facturaas
        SQL = "Select numasien,numdiari,fechaent "
        SQL = SQL & " FROM cabfact"
        If Opcion = 6 Then SQL = SQL & "prov"
        SQL = SQL & " WHERE numasien>0 "
        If Opcion = 5 Then
            CadenaDesdeOtroForm = "fecfaccl"
            If txtCLI.Text <> "" Then SQL = SQL & " AND numserie ='" & txtCLI.Text & "'"
        Else
            CadenaDesdeOtroForm = "fecrecpr"
        End If
        
        If Text1(3).Text <> "" Then SQL = SQL & " AND " & CadenaDesdeOtroForm & " >='" & Format(Text1(3).Text, FormatoFecha) & "'"
        If Text1(2).Text <> "" Then SQL = SQL & " AND " & CadenaDesdeOtroForm & " <='" & Format(Text1(2).Text, FormatoFecha) & "'"
        SQL = SQL & " GROUP BY numasien,numdiari,fechaent"
        
        SQL = SQL & " ORDER BY numdiari,numasien,fechaent"
    
        
        
        
        miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        
        'Recorremos las facturas
        I = 0
        While Not RS.EOF
            'Label7(1).Caption = RS!Numasien & " " & RS!fechaent
            Label7(1).Caption = I & " de " & NumRegElim
            Label7(1).Refresh
            If Not EstaEnMirsaux Then
            
                InsertaItemsFacturasContabilizadas False
            End If
            RS.MoveNext
            I = I + 1
            If (I Mod 50) = 0 Then
                Me.Refresh
                DoEvents
            End If
        Wend
        miRsAux.Close
    End If 'Rs.eof
    RS.Close
    
End Sub

Private Sub QuitarOtrasCuentas()
Dim I As Integer
    Set RS = New ADODB.Recordset
    
    'pRIMERO DE LAS CUENTAS BANCARIAS
    'codmacta ctagastos ctaingreso ctagastostarj    ctabancaria
    Label4.Caption = "Cta bancaria "
    Label4.Refresh
    SQL = "Select codmacta , ctagastos , ctaingreso , ctagastostarj   FROM ctabancaria"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "|"
    While Not RS.EOF
        For I = 0 To 3
            If Not IsNull(RS.Fields(I)) Then
                If RS.Fields(I) <> "" Then
                    If InStr(1, SQL, "|" & RS.Fields(I) & "|") = 0 Then SQL = SQL & RS.Fields(I) & "|"
                End If
            End If
        Next
        RS.MoveNext
    Wend
    RS.Close
    
    SQL = Mid(SQL, 2)
    While SQL <> ""
        I = InStr(1, SQL, "|")
        Conn.Execute "Delete from tmpbussinmov where codmacta ='" & RecuperaValor(SQL, 1) & "';"
        SQL = Mid(SQL, I + 1)
    Wend
    
    'PARAMETROS
    SQL = "Delete from tmpbussinmov where codmacta ='" & vParam.ctaperga & "';"
    Conn.Execute SQL
    espera 0.2
    
    If vEmpresa.TieneTesoreria Then
        SQL = "SELECT ctabenbanc  ,par_pen_apli,RemesaCancelacion,RemesaConfirmacion,taloncta,pagarectaPRO,talonctaPRO,ctaefectcomerciales from paramtesor"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            For I = 0 To RS.Fields.Count - 1
                SQL = DBLet(RS.Fields(I), "T")
                If SQL <> "" Then
                    If Len(SQL) = vEmpresa.DigitosUltimoNivel Then
                        SQL = "Delete from tmpbussinmov where codmacta ='" & SQL & "';"
                        Conn.Execute SQL
                    End If
                End If
            Next I
        End If
        RS.Close
    End If
    'IVAS
    Label4.Caption = "IVAS "
    Label4.Refresh
    SQL = "SELECT cuentare ,cuentarr ,cuentaso ,cuentasr ,cuentasn from tiposiva "
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "|"
    While Not RS.EOF
        For I = 0 To 4
            If Not IsNull(RS.Fields(I)) Then
                If RS.Fields(I) <> "" Then
                    If InStr(1, SQL, "|" & RS.Fields(I) & "|") = 0 Then SQL = SQL & RS.Fields(I) & "|"
                End If
            End If
        Next
        RS.MoveNext
    Wend
    RS.Close
    
    SQL = Mid(SQL, 2)
    While SQL <> ""
        I = InStr(1, SQL, "|")
        Conn.Execute "Delete from tmpbussinmov where codmacta ='" & RecuperaValor(SQL, 1) & "';"
        SQL = Mid(SQL, I + 1)
    Wend
    Set RS = Nothing
    
    
    
    
End Sub
    

