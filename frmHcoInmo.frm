VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmHcoInmo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hist�rico inmovilizado"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmHcoInmo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   320
      Left            =   840
      TabIndex        =   15
      Top             =   5640
      Width           =   135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   4680
      TabIndex        =   4
      Tag             =   "Importe|N|N|0||shisin|imporinm|#,###,##0.00||"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3840
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "Porcentaje|N|N|0||shisin|porcinm|##0.00||"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||shisin|fechainm|dd/mm/yyyy|S|"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5340
      TabIndex        =   5
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6540
      TabIndex        =   6
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "Elto inmovilizado|N|N|0||shisin|codinmov||S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6540
      TabIndex        =   9
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   7
      Top             =   5895
      Width           =   1905
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
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
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmHcoInmo.frx":000C
      Height          =   5265
      Left            =   60
      TabIndex        =   12
      Top             =   600
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9287
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   5970
      Top             =   0
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total �"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   6135
      Width           =   615
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
Attribute VB_Name = "frmHcoInmo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmEI As frmEltoInmo
Attribute frmEI.VB_VarHelpID = -1

Private CadenaConsulta As String
Private TextoBusqueda As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim jj As Integer
Dim SQL As String

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
    Modo = vModo
    
    B = (Modo = 0)
    
    For jj = 0 To 4
        txtAux(jj).Visible = Not B
    Next jj
    cmdAux.Visible = Not B
    mnOpciones.Enabled = B
    Toolbar1.Buttons(1).Enabled = B
    Toolbar1.Buttons(2).Enabled = B
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    'DataGrid1.Enabled = b
    
    'Si es regresar
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.Visible = b
'    End If
    'Si estamo mod or insert
    If Modo = 2 Then
       txtAux(0).BackColor = &H80000018
       Else
        txtAux(0).BackColor = &H80000005
    End If
    txtAux(0).Enabled = (Modo <> 2)
    txtAux(2).Enabled = txtAux(0).Enabled
    txtAux(2).BackColor = txtAux(0).BackColor
    cmdAux.Enabled = txtAux(0).Enabled
End Sub

Private Sub BotonAnyadir()
    Dim anc As Single

    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Not adodc1.Recordset.EOF Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.Top + 210
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If
    txtAux(0).Text = ""
    For jj = 1 To 4
        txtAux(jj).Text = ""
    Next jj
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtAux(0).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub



Private Sub BotonVerTodos()
    DataGrid1.Enabled = False
    espera 0.1
    TextoBusqueda = ""
    CargaGrid ""
    DataGrid1.Enabled = True
End Sub

Private Sub BotonBuscar()
    DataGrid1.Enabled = False
    CargaGrid " shisin.codinmov = -1"
    DataGrid1.Enabled = True
    'Buscar
    For jj = 0 To 4
        txtAux(jj).Text = ""
    Next jj
    LLamaLineas DataGrid1.Top + 206, 2
    txtAux(0).SetFocus
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    'If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 600
    End If

    'Llamamos al form
    For jj = 0 To 2
        txtAux(jj).Text = DataGrid1.Columns(jj).Text
    Next jj
    'El porcentaje
    SQL = adodc1.Recordset!porcinm
    txtAux(3).Text = TransformaComasPuntos(SQL)
        'El porcentaje
    SQL = adodc1.Recordset!imporinm
    txtAux(4).Text = TransformaComasPuntos(SQL)
    
    LLamaLineas anc, 1
   
    'Como es modificar
    txtAux(3).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
DeseleccionaGrid
PonerModo xModo + 1
'Fijamos el ancho
For jj = 0 To 4
    txtAux(jj).Top = alto
Next jj
cmdAux.Top = alto
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    '### a mano
    SQL = "Seguro que desea eliminar la linea de hist�rico:" & vbCrLf
    SQL = SQL & vbCrLf & "Inmovilizado: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominaci�n: " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Fecha       : " & adodc1.Recordset.Fields(2)
    SQL = SQL & vbCrLf & "Importe(�)  : " & adodc1.Recordset.Fields(3)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from shisin where codinmov=" & adodc1.Recordset!Codinmov
        SQL = SQL & " AND fechainm ='" & Format(adodc1.Recordset!fechainm, FormatoFecha) & "';"
        Conn.Execute SQL
        CargaGrid ""
        adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                Conn.Execute "commit"
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    Conn.Execute "commit"
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        
        'Para el texto
        TextoBusqueda = ""
        If txtAux(0).Text <> "" Then TextoBusqueda = TextoBusqueda & "Cod. Inmov " & txtAux(0).Text
        If txtAux(2).Text <> "" Then TextoBusqueda = TextoBusqueda & "Fecha " & txtAux(2).Text
        If txtAux(3).Text <> "" Then TextoBusqueda = TextoBusqueda & "Porcentaje " & txtAux(3).Text
        If txtAux(4).Text <> "" Then TextoBusqueda = TextoBusqueda & "Importe " & txtAux(4).Text
        
            
        
        If CadB <> "" Then
            PonerModo 0
            DataGrid1.Enabled = False
            CargaGrid CadB
            DataGrid1.Enabled = True
        End If
    End Select


End Sub

Private Sub cmdAux_Click()
    Screen.MousePointer = vbHourglass
    Set frmEI = New frmEltoInmo
    frmEI.DatosADevolverBusqueda = "0|1|"
    frmEI.Show vbModal
    Set frmEI = Nothing
    If txtAux(0).Text <> "" Then txtAux(2).SetFocus
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
TextoBusqueda = ""
DataGrid1.SetFocus
End Sub

'Private Sub cmdRegresar_Click()
'Dim cad As String
'
'If adodc1.Recordset.EOF Then
'    MsgBox "Ning�n registro a devolver.", vbExclamation
'    Exit Sub
'End If
'
'cad = adodc1.Recordset.Fields(0) & "|"
'cad = cad & adodc1.Recordset.Fields(1) & "|"
'cad = cad & adodc1.Recordset.Fields(2) & "|"
'RaiseEvent DatoSeleccionado(cad)
'Unload Me
'End Sub

Private Sub DataGrid1_DblClick()
'If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        '.Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With

    Set miTag = New CTag
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
'    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    
    'cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    CadenaConsulta = "SELECT shisin.codinmov,sinmov.nominmov,fechainm,porcinm,imporinm "
    CadenaConsulta = CadenaConsulta & " FROM shisin,sinmov WHERE "
    CadenaConsulta = CadenaConsulta & " shisin.codinmov=sinmov.codinmov"
    
    CargaGrid "shisin.codinmov = -1 "  'Para k lo carge vacio
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Set miTag = Nothing
End Sub


Private Sub frmEI_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
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





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
        BotonBuscar
Case 2
        BotonVerTodos
Case 6
        BotonAnyadir
Case 7
        BotonModificar
Case 8
        BotonEliminar
Case 11
        'Ha ido bien
        If GeneraDatosHcoInmov(adodc1.RecordSource) Then
            frmImprimir.Opcion = 53
            frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            If TextoBusqueda <> "" Then
                frmImprimir.OtrosParametros = "Camposeleccion= """ & TextoBusqueda & """|"
                frmImprimir.NumeroParametros = 1
            Else
                frmImprimir.OtrosParametros = ""
                frmImprimir.NumeroParametros = 0
            End If
            frmImprimir.Show vbModal
        End If
Case 12
        Unload Me
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = bol
    Next I
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    Text1.Text = ""
    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codinmov,fechainm"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
        
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Cod."
        DataGrid1.Columns(I).Width = 600

    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Descripci�n"
        DataGrid1.Columns(I).Width = 2800
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'El importe es campo calculado
    I = 2
        DataGrid1.Columns(I).Caption = "Fecha"
        DataGrid1.Columns(I).Width = 1200
        DataGrid1.Columns(I).NumberFormat = "dd/mm/yyyy"
        
    I = 3
        DataGrid1.Columns(I).Caption = "Porc."
        DataGrid1.Columns(I).Width = 800
        DataGrid1.Columns(I).NumberFormat = "##0.00"
        DataGrid1.Columns(I).Alignment = dbgRight
        
    I = 4
        DataGrid1.Columns(I).Caption = "Importe"
        DataGrid1.Columns(I).Width = 1300
        DataGrid1.Columns(I).NumberFormat = "#,###,###0.00"
        DataGrid1.Columns(I).Alignment = dbgRight
    
    For I = 0 To 4
        DataGrid1.Columns(I).AllowSizing = False
    Next I
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Left = DataGrid1.Left + 340
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        cmdAux.Left = DataGrid1.Columns(1).Left
        txtAux(1).Left = cmdAux.Left + cmdAux.Width
        txtAux(1).Width = DataGrid1.Columns(1).Width - cmdAux.Width + 15
        txtAux(2).Left = DataGrid1.Columns(2).Left + 60
        txtAux(2).Width = DataGrid1.Columns(2).Width - 15
        txtAux(3).Left = DataGrid1.Columns(3).Left + 60
        txtAux(3).Width = DataGrid1.Columns(3).Width - 15
        txtAux(4).Left = DataGrid1.Columns(4).Left + 60
        txtAux(4).Width = DataGrid1.Columns(4).Width - 15
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
    CargarSumas vSQL
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
With txtAux(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim Valor As Currency
txtAux(Index).Text = Trim(txtAux(Index).Text)
If Modo = 3 Then Exit Sub 'Busquedas

If txtAux(Index).Text = "" Then
    If Index = 0 Then txtAux(1).Text = ""
    Exit Sub
End If

Select Case Index
Case 0
    txtAux(1).Text = ""
    If Not IsNumeric(txtAux(0).Text) Then
        MsgBox "C�digo de elemento de inmovilizado tiene que ser num�rico", vbExclamation
        txtAux(0).SetFocus
        Exit Sub
    End If
    SQL = DevuelveDesdeBD("nominmov", "sinmov", "codinmov", txtAux(0).Text, "N")
    If SQL = "" Then
        MsgBox "Ningun elemento de inmovilizado para : " & txtAux(0).Text, vbExclamation
        txtAux(0).Text = ""
        txtAux(0).SetFocus
    End If
    txtAux(0).Text = Format(txtAux(0).Text, "00000")
    txtAux(1).Text = SQL
    
Case 2
    If Not EsFechaOK(txtAux(2)) Then
        MsgBox "Fecha incorrecta: " & txtAux(2), vbExclamation
        txtAux(2).Text = ""
        txtAux(2).SetFocus
    End If
Case Else

    miTag.Cargar txtAux(Index)
    If miTag.Comprobar(txtAux(Index)) Then
        If Index > 2 Then
            'Son los numeros
            If InStr(1, txtAux(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(txtAux(Index).Text)
                SQL = CStr(Valor)
            Else
                 SQL = TransformaPuntosComas(txtAux(Index).Text)
            End If
            txtAux(Index).Text = Format(SQL, FormatoImporte)
        End If
        
    Else
        'Error con los datos
        txtAux(Index).Text = ""
        If Modo <> 0 Then txtAux(Index).SetFocus
    End If
    
End Select
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function

If Modo = 1 Then
    'Estamos insertando
    
End If
DatosOk = B
End Function

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function



Private Sub CargarSumas(ByRef vS As String)
On Error GoTo ECargarSumas

    Set miRsAux = New ADODB.Recordset
    SQL = "Select sum(imporinm) from shisin"
    If vS <> "" Then SQL = SQL & " WHERE  " & vS
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then _
            Text1.Text = Format(miRsAux.Fields(0), FormatoImporte)
    End If
    miRsAux.Close
ECargarSumas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar sumas"
    Set miRsAux = Nothing
End Sub
