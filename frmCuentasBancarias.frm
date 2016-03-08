VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmCuentasBancarias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas BANCARIAS"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "frmCuentasBancarias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCta 
      Caption         =   "+"
      Height          =   290
      Left            =   840
      TabIndex        =   14
      Top             =   5640
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   6300
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Cuenta del banco|T|N|0||ctabancaria|ctabanco|0000000000||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   4800
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Control|N|S|0|99|ctabancaria|control|00||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Oficina|N|N|0|9999|ctabancaria|oficina|0000||"
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
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Entidad|N|N|0|9999|ctabancaria|entidad|0000||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6420
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7620
      TabIndex        =   7
      Top             =   6000
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
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Cta contable|N|N|||ctabancaria|codmacta||S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7620
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   8
      Top             =   5895
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
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
         Left            =   4560
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCuentasBancarias.frx":000C
      Height          =   5145
      Left            =   60
      TabIndex        =   13
      Top             =   720
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   9075
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
Attribute VB_Name = "frmCuentasBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////7
'//
'//
'// Cuenta BANCARIA - Cta contable

'Tag: Nombre concepto|T|N|||sconam|nomconam|||
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private CadenaConsulta As String
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


cmdCta.Visible = Modo = 1

For jj = 0 To 5
    txtaux(jj).Visible = Not B
Next jj
mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
'DataGrid1.Enabled = b

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If
'Si estamo mod or insert
If Modo = 2 Then
   txtaux(0).BackColor = &H80000018
   Else
    txtaux(0).BackColor = &H80000005
End If
txtaux(0).Enabled = (Modo <> 2)

End Sub

Private Sub BotonAnyadir()
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Not Adodc1.Recordset.EOF Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
    anc = DataGrid1.Top
    If DataGrid1.Row >= 0 Then
        anc = DataGrid1.RowTop(DataGrid1.Row) + anc
    Else
        anc = anc + 210
    End If
    
    For jj = 0 To 5
        txtaux(jj).Text = ""
    Next jj
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtaux(0).SetFocus
    
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
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "entidad = -1"
    'Buscar
    For jj = 0 To txtaux.Count - 1
        txtaux(jj).Text = ""
    Next jj
    LLamaLineas DataGrid1.Top + 206, 2
    txtaux(0).SetFocus
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If Adodc1.Recordset.EOF Then Exit Sub
    'If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    
    anc = DataGrid1.Top
    If DataGrid1.Row >= 0 Then anc = DataGrid1.RowTop(DataGrid1.Row) + anc
    
    'Llamamos al form
    For jj = 0 To 5
        txtaux(jj).Text = DataGrid1.Columns(jj).Text
    Next jj
    LLamaLineas anc, 1
   
   'Como es modificar
   txtaux(2).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid DataGrid1
    PonerModo xModo + 1
    cmdCta.Top = alto
    'Fijamos el ancho
    For jj = 0 To 5
        txtaux(jj).Top = alto
    Next jj
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
 
    If Not SepuedeBorrar Then Exit Sub
    
    
    '### a mano
    SQL = "Seguro que desea eliminar la linea :"
    SQL = SQL & vbCrLf & "Cuenta: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & Adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Cta bancaria: " & Adodc1.Recordset.Fields(2) & " - " & Adodc1.Recordset.Fields(3) & " - ** - " & Adodc1.Recordset.Fields(5)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from ctabancaria where codmacta = '" & Adodc1.Recordset!Codmacta & "'"
        Conn.Execute SQL
        espera 0.5
        CargaGrid ""
        Adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub


Private Function SepuedeBorrar() As Boolean
    If Not ComprobandoEliminar("norma43", "codmacta") Then Exit Function

    If vEmpresa.TieneTesoreria Then
        SepuedeBorrar = False
        'Falta ver algunas cosas
        'Como es cuenta bancaria
        If Not ComprobandoEliminar("scaja", "codmacta") Then Exit Function
        
        If Not ComprobandoEliminar("usucaja", "ctacaja") Then Exit Function
        

        If Not ComprobandoEliminar("spagop", "ctabanc1") Then Exit Function
        If Not ComprobandoEliminar("spagop", "ctabanc2") Then Exit Function
        
        If Not ComprobandoEliminar("scobro", "ctabanc1") Then Exit Function
        If Not ComprobandoEliminar("scobro", "ctabanc2") Then Exit Function
        If Not ComprobandoEliminar("sgastfij", "ctaprevista") Then Exit Function
        
        
        SepuedeBorrar = True
    Else
        SepuedeBorrar = True
    End If
End Function


Private Function ComprobandoEliminar(Tabla As String, desca As String) As Boolean

    SQL = DevuelveDesdeBD(desca, Tabla, desca, Adodc1.Recordset.Fields(0), "T")
    If SQL = "" Then
        ComprobandoEliminar = True
    Else
        MsgBox "Existe referencia en: " & Tabla & "." & desca, vbExclamation
        ComprobandoEliminar = False
    End If
End Function


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
                    I = Adodc1.Recordset.AbsolutePosition
                    PonerModo 0
                    CargaGrid
                    Adodc1.Recordset.Move I - 1
                    lblIndicador.Caption = ""
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    End Select


End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub

Private Sub cmdCta_Click()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

If Adodc1.Recordset.EOF Then
    MsgBox "Ningún registro a devolver.", vbExclamation
    Exit Sub
End If

Cad = Adodc1.Recordset.Fields(0) & "|"
Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub




Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible Then cmdRegresar_Click
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

    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
   
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
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
'        If GeneraDatosConceptosInmov Then
'            frmImprimir.Opcion = 54
'            frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
'            frmImprimir.NumeroParametros = 0
'            frmImprimir.Show vbModal
'        End If
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
    
    Adodc1.ConnectionString = Conn
    PonerSQL
    If vSQL <> "" Then SQL = SQL & " AND " & vSQL
    SQL = SQL & " ORDER BY ctabancaria.codmacta"
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    'Cuenta contable
    I = 0
        DataGrid1.Columns(I).Caption = "Cuenta"
        DataGrid1.Columns(I).Width = 1100
    
    'Descripcion NOMMACTA
    I = 1
        DataGrid1.Columns(I).Caption = "Descripción"
        DataGrid1.Columns(I).Width = 3200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'Entidad
    I = 2
        DataGrid1.Columns(I).Caption = "Entidad"
        DataGrid1.Columns(I).Width = 800
        DataGrid1.Columns(I).NumberFormat = "0000"
        
    I = 3
        DataGrid1.Columns(I).Caption = "Oficina"
        DataGrid1.Columns(I).Width = 900
        DataGrid1.Columns(I).NumberFormat = "0000"
        
    'Codigo control
    I = 4
        DataGrid1.Columns(I).Caption = "Ctr"
        DataGrid1.Columns(I).Width = 400
        
    'Cueta bancaria
    I = 5
        DataGrid1.Columns(I).Caption = "Cuenta bancaria"
        DataGrid1.Columns(I).Width = 1500
        DataGrid1.Columns(I).NumberFormat = "0000000000"
    
        
    For I = 0 To 3
        DataGrid1.Columns(I).AllowSizing = False
    Next I
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtaux(0).Left = DataGrid1.Left + 340
        txtaux(0).Width = DataGrid1.Columns(0).Width - 60
        
        
        For jj = 1 To 5
            txtaux(jj).Width = DataGrid1.Columns(jj).Width - 60
            txtaux(jj).Left = txtaux(jj - 1).Left + txtaux(jj - 1).Width + 60
        Next jj
        txtaux(5).Left = txtaux(5).Left + 15
        CadAncho = True
        
        'El botoncito para la cuenta
        cmdCta.Left = txtaux(1).Left - 180
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    PonFoco txtaux(Index)

End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim RC As String
    txtaux(Index).Text = Trim(txtaux(Index).Text)
    If txtaux(Index).Text = "" Then Exit Sub
    If Modo = 3 Then Exit Sub 'Busquedas
    Select Case Index
    Case 0
        RC = txtaux(0).Text
        If CuentaCorrectaUltimoNivel(RC, SQL) Then
            txtaux(0).Text = RC
            txtaux(1).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            txtaux(0).Text = ""
            txtaux(1).Text = ""
            txtaux(0).SetFocus
        End If
    
    Case Else
        If Not IsNumeric(txtaux(Index).Text) Then
            MsgBox "El campo debe ser numérico: " & txtaux(Index).Text, vbExclamation
            txtaux(Index).Text = ""
            txtaux(Index).SetFocus
        End If
    End Select

End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Como los campos son obilgatorios, no compurebo si tein valor
    SQL = txtaux(2).Text & txtaux(3).Text & txtaux(5).Text
    If Len(SQL) <> 18 Then
        MsgBox "Longitud cuenta bancaria incorrecta", vbExclamation
        Exit Function
    Else
    
        SQL = CodigoDeControl(SQL)
        If SQL <> txtaux(4).Text Then
            SQL = "Codigo de control para la cuenta bancaria: " & SQL & vbCrLf
            SQL = SQL & "Codigo de control introducido: " & txtaux(4).Text & vbCrLf & vbCrLf
            SQL = SQL & "Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
    End If
    
    DatosOk = B
End Function

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


Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Sub PonerSQL()
    SQL = "Select ctabancaria.codmacta,cuentas.nommacta,ctabancaria.entidad,ctabancaria.oficina,ctabancaria.control,ctabancaria.ctabanco from ctabancaria,cuentas WHERE ctabancaria.codmacta = cuentas.codmacta"
End Sub

