VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmContadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contadores"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frmContadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmContadores.frx":000C
      Left            =   6240
      List            =   "frmContadores.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "Fra cta ajena|N|N|||contadores|FacliAjena|||"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Tag             =   "Siguiente|N|N|0||contadores|contado2|||"
      Text            =   "Dato4"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Tag             =   "Cont. actual|N|N|0||contadores|contado1|||"
      Text            =   "Dato3"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3780
      TabIndex        =   4
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "Denominaci�n|T|N|||contadores|nomregis||N|"
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
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Tipo registro|T|N|||contadores|tiporegi||S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmContadores.frx":0022
      Height          =   5205
      Left            =   60
      TabIndex        =   9
      Top             =   660
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   9181
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   5895
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6000
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Verificar contador"
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
         Left            =   4680
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3240
      Picture         =   "frmContadores.frx":0037
      Top             =   6120
      Width           =   240
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
Attribute VB_Name = "frmContadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte

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

txtaux(0).Visible = Not B
txtaux(1).Visible = Not B
txtaux(2).Visible = Not B
txtaux(3).Visible = Not B
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid1.Enabled = B
Me.mnOpciones.Enabled = B
Combo1.Visible = Not B


Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B

Toolbar1.Buttons(6).Enabled = B And vUsu.Nivel < 2
Me.mnNuevo.Enabled = Toolbar1.Buttons(6).Enabled
Toolbar1.Buttons(7).Enabled = B And vUsu.Nivel < 2
Me.mnModificar.Enabled = Toolbar1.Buttons(7).Enabled
Toolbar1.Buttons(8).Enabled = B And vUsu.Nivel < 2
Me.mnEliminar.Enabled = Toolbar1.Buttons(8).Enabled


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
    

    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    txtaux(1).Enabled = True
    
   
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.Top + 240
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If
    txtaux(0).Text = ""
    txtaux(1).Text = ""
    txtaux(2).Text = ""
    txtaux(3).Text = ""
    
    LLamaLineas anc, 0
    Combo1.ListIndex = 0
    
    'Ponemos el foco
    txtaux(0).SetFocus

End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    cmdCancelar.Visible = True
    cmdCancelar.SetFocus
    CargaGrid "tiporegi = ' '"
    'Buscar
    txtaux(0).Text = "":    txtaux(1).Text = "": txtaux(2).Text = "": txtaux(3).Text = ""
    Combo1.ListIndex = -1
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
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    'Peculiar
    

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
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If

    'Llamamos al form
    For I = 0 To 3
        txtaux(I).Text = DataGrid1.Columns(I).Text
    Next I
    If DataGrid1.Columns(4).Text = "" Then
        Combo1.ListIndex = 0
    Else
        Combo1.ListIndex = 1
    End If
    LLamaLineas anc, 1
   
   'a mano###
    If Adodc1.Recordset!tiporegi = "0" Or Adodc1.Recordset!tiporegi = "1" Then
        txtaux(1).Enabled = False
    Else
        'Como es modificar
        txtaux(1).Enabled = True
        txtaux(1).SetFocus
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
PonerModo xModo + 1
'Fijamos el ancho
txtaux(0).Top = alto
txtaux(1).Top = alto
txtaux(2).Top = alto
txtaux(3).Top = alto
Combo1.Top = alto
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset!tiporegi = "0" Or Adodc1.Recordset!tiporegi = "1" Or Adodc1.Recordset!tiporegi = "Z" Then
        MsgBox "Este contador no se puede eliminar", vbExclamation
        Exit Sub
    End If
    '### a mano
    SQL = "Seguro que desea eliminar el contador:"
    SQL = SQL & vbCrLf & "C�digo: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominaci�n: " & Adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from contadores where tiporegi='" & Adodc1.Recordset!tiporegi & "'"
        Conn.Execute SQL
        CargaGrid ""
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
        End If
End Sub





Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
On Error GoTo EAceptar
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
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
                    CadB = Adodc1.Recordset.Fields(0)
                    DataGrid1.Enabled = False
                    CargaGrid
                    Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " = '" & CadB & "'")
                    PonerModo 0
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            DataGrid1.Enabled = False
            CargaGrid CadB
            DataGrid1.Enabled = True
        End If
    End Select

Exit Sub
EAceptar:
    Err.Clear
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

Private Sub cmdRegresar_Click()
Dim Cad As String

If Adodc1.Recordset.EOF Then
    MsgBox "Ning�n registro a devolver.", vbExclamation
    Exit Sub
End If

If Asc(Adodc1.Recordset.Fields(0)) <= 57 Then
    MsgBox "No es contador de tipo factura.", vbExclamation
    Exit Sub
End If


Cad = Adodc1.Recordset.Fields(0) & "|"
Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
Cad = Cad & Adodc1.Recordset.Fields(2) & "|"
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
        .Buttons(11).Image = 20 'Verificar contador
        .Buttons(12).Image = 15
'        .Buttons(14).Image = 6
'        .Buttons(15).Image = 7
'        .Buttons(16).Image = 8
'        .Buttons(17).Image = 9
    End With


    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    'Cadena consulta
    CadenaConsulta = "Select tiporegi,nomregis,contado1,contado2,if(FacliAjena=0,"""",""Si"") from contadores "
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub Image1_Click()
    Image1.Tag = "Facturaci�n Ajena" & vbCrLf & String(60, "-") & vbCrLf & vbCrLf
    Image1.Tag = Image1.Tag & "Cuando marque esta opcion a la serie de facturas, en el 340 se declarar�" & vbCrLf
    Image1.Tag = Image1.Tag & "en lugar del codigo factura lo que haya en el campo de observaciones factura"
    MsgBox Image1.Tag, vbInformation
    Image1.Tag = ""
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



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
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
        Screen.MousePointer = vbHourglass
        ComprobarContadores
        Screen.MousePointer = vbDefault
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

Private Sub CargaGrid(Optional SQL As String)
Dim B As Boolean
    B = DataGrid1.Enabled
    CargaGrid2 SQL
    DataGrid1.Enabled = B
End Sub



Private Sub CargaGrid2(Optional SQL As String)
    Dim I As Integer
    Dim anc As Integer
    Adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY tiporegi"
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 295
    
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Tipo"
        DataGrid1.Columns(I).Width = 500
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Denominaci�n"
        DataGrid1.Columns(I).Width = 2900
    
    'El importe es campo calculado
    I = 2
        DataGrid1.Columns(I).Caption = "Actual"
        DataGrid1.Columns(I).Width = 900
        DataGrid1.Columns(I).Alignment = dbgRight
        
    I = 3
        DataGrid1.Columns(I).Caption = "Siguiente"
        DataGrid1.Columns(I).Width = 900
        DataGrid1.Columns(I).Alignment = dbgRight
    
    I = 4
        DataGrid1.Columns(I).Caption = "Fac.ajena"
        DataGrid1.Columns(I).Width = 900
        DataGrid1.Columns(I).Alignment = dbgRight
    
    
    
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        anc = 60
        For I = 0 To 3
            txtaux(I).Left = DataGrid1.Columns(I).Left + anc
            txtaux(I).Width = DataGrid1.Columns(I).Width - 15
        Next I
        Combo1.Left = DataGrid1.Columns(I).Left + anc
        Combo1.Width = DataGrid1.Columns(I).Width - 15
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    With txtaux(Index)
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

txtaux(Index).Text = Trim(txtaux(Index).Text)
If txtaux(Index).Text = "" Then Exit Sub
If Modo = 3 Then Exit Sub 'Busquedas
If Index > 1 Then
    If Not IsNumeric(txtaux(Index).Text) Then
        MsgBox "Los contadores tiene que ser num�ricos", vbExclamation
        Exit Sub
    End If
    Else
        If Index = 0 Then txtaux(0).Text = UCase(txtaux(0).Text)
End If
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function

If InStr(1, txtaux(0).Text, " ") > 0 Then
    MsgBox "No se permiten blancos", vbExclamation
    Exit Function
End If


'Del 0 al 9 nos los reservamos.
'   0-Asientos 1- Proveedores 2.- Contador confirming
'
If Modo = 1 Then
    If Len(txtaux(0).Text) = 1 Then
        If IsNumeric(txtaux(0).Text) Then
            MsgBox "Reservados por la aplicaci�n", vbExclamation
            If vUsu.Nivel > 0 Then Exit Function
        End If
    End If
End If


If IsNumeric(txtaux(0).Text) Then
    If Combo1.ListIndex = 1 Then MsgBox "Facturaci�n por cuenta ajena valido solo para SERIES DE FACTURAS.", vbExclamation
End If
If Modo = 1 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD("tiporegi", "contadores", "tiporegi", txtaux(0).Text, "T")
     If Datos <> "" Then
        MsgBox "Ya existe el contador : " & txtaux(0).Text, vbExclamation
        B = False
    End If
End If
DatosOk = B
End Function


Private Function ComprobarContadores()
Dim I As Long
Dim F As Date
Dim SQL As String
Dim AUX2 As String
Dim Aux As String
Dim MaxA As Long
Dim CadenaError As String
    Set miRsAux = New ADODB.Recordset


    '-------------------------------------------------------
    '-------------------------------------------------------
    'Probamos los contadores
    '----------------------------------------------------------
    '-------------------------------------------------------
    '
    ' Asientos
    'actual
    CadenaError = ""
    NumRegElim = 1
    I = 2
    F = vParam.fechaini
    SQL = "Select max(numasien) from hlinapu where fechaent>='" & Format(F, FormatoFecha) & "'"
    F = vParam.fechafin
    SQL = SQL & " AND fechaent<='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        NumRegElim = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    '     introduccion de asientos
    MaxA = 0
    F = vParam.fechaini
    SQL = "Select max(numasien) from cabapu where fechaent>='" & Format(F, FormatoFecha) & "'"
    F = vParam.fechafin
    SQL = SQL & " AND fechaent<='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        MaxA = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    Aux = " YA contabilizado"
    If MaxA > NumRegElim Then
        NumRegElim = MaxA
        Aux = " en introduccion de apuntes"
    End If
    
    '--------------------------------------
    'siguiente   --------------------------
    F = DateAdd("yyyy", 1, vParam.fechaini)
    SQL = "Select max(numasien) from hlinapu where fechaent>='" & Format(F, FormatoFecha) & "'"
    F = DateAdd("yyyy", 1, vParam.fechafin)
    SQL = SQL & " AND fechaent<='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        I = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    ' en introduccion de asientos
    MaxA = 0
    F = DateAdd("yyyy", 1, vParam.fechaini)
    SQL = "Select max(numasien) from cabapu where fechaent>='" & Format(F, FormatoFecha) & "'"
    F = DateAdd("yyyy", 1, vParam.fechafin)
    SQL = SQL & " AND fechaent<='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        MaxA = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    AUX2 = " YA contabilizado"
    If MaxA > I Then
        I = MaxA
        AUX2 = " en introduccion de apuntes"
    End If
    
    'Ahora vemos si son correctos los contadores
    CadenaDesdeOtroForm = "contado2"
    SQL = DevuelveDesdeBD("contado1", "contadores", "tiporegi", "0", "T", CadenaDesdeOtroForm)
    If SQL <> "" Then
        If Val(SQL) <> NumRegElim Then
'            MsgBox "El numero asiento actual en contadores es distinto al mayor numero de asiento" & Aux & _
                vbCrLf & " Contadores: " & SQL & "    --    Asiento: " & NumRegElim, vbExclamation
            Aux = "          Actual.     " & Aux & " : " & _
                 Format(NumRegElim, "000000") & "    --    Contadores: " & Format(SQL, "000000") & vbCrLf
            CadenaError = CadenaError & Aux
        End If
        If Val(CadenaDesdeOtroForm) <> I Then
            'Si es distinto de uno, pq el Nos reservamos el 1 para el asiento de apertura
            If I > 1 Then
                
                AUX2 = "          Siguiente. " & AUX2 & ": " & _
                     Format(I, "000000") & "    --    Contadores: " & Format(CadenaDesdeOtroForm, "000000") & vbCrLf
                CadenaError = CadenaError & AUX2
            End If
        End If
        If CadenaError <> "" Then CadenaError = "ASIENTOS." & vbCrLf & CadenaError & vbCrLf & vbCrLf
    End If
    
    
    '-------------------------------------------------------
    '-------------------------------------------------------
    'Comprobaremos si querenos las facturas proveedores
    '-------------------------------------------------------
    '-------------------------------------------------------
    NumRegElim = 1
    I = 2
    F = vParam.fechaini
    SQL = "Select max(numregis) from cabfactprov where fecrecpr >='" & Format(F, FormatoFecha) & "'"
    F = vParam.fechafin
    SQL = SQL & " AND fecrecpr <='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        NumRegElim = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    'siguiente
    F = DateAdd("yyyy", 1, vParam.fechaini)
    SQL = "Select max(numregis) from cabfactprov where fecrecpr >='" & Format(F, FormatoFecha) & "'"
    F = DateAdd("yyyy", 1, vParam.fechafin)
    SQL = SQL & " AND fecrecpr <='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        I = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    
    'Ahora vemos si son correcto los contadores
    CadenaDesdeOtroForm = "contado2"
    SQL = DevuelveDesdeBD("contado1", "contadores", "tiporegi", "1", "T", CadenaDesdeOtroForm)
    If SQL <> "" Then
        AUX2 = ""
        If Val(SQL) <> NumRegElim Then
            AUX2 = AUX2 & "          Actual:          " & Format(SQL, "000000") & "    --    Contadores: " & Format(NumRegElim, "000000") & vbCrLf
        End If
        If Val(CadenaDesdeOtroForm) <> I Then
            AUX2 = AUX2 & "          Siguiente:   " & Format(CadenaDesdeOtroForm, "000000") & "    --    Contadores: " & Format(I, "000000") & vbCrLf
        End If
        If AUX2 <> "" Then CadenaError = CadenaError & "Facturas proveedor: " & vbCrLf & AUX2
            
    End If
    
    
    
    
    'Podemos comprobar las facturas tb
    'Por eso recorreremos el adodc1
    
    If Adodc1.Recordset.RecordCount > 0 Then
        'Para devlverlo a su posicion
        NumRegElim = Adodc1.Recordset.AbsolutePosition
        
        
        
        Adodc1.Recordset.MoveFirst
        Me.Tag = ""
        While Not Adodc1.Recordset.EOF
            CadenaDesdeOtroForm = ""
            If Adodc1.Recordset.Fields(0) <> "0" And Adodc1.Recordset.Fields(0) <> "1" Then
                'Son FACTURAS DE CLIENTES
                'Actual
                I = DBLet(Adodc1.Recordset!Contado1, "N")
                AUX2 = MontaSQLFacCli(True)
                Aux = "Select max(codfaccl) from cabfact where numserie = '" & Adodc1.Recordset!tiporegi
                SQL = Aux & "' AND " & AUX2
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                MaxA = 0
                If Not miRsAux.EOF Then MaxA = DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
                
                If I <> MaxA Then
                    
                    AUX2 = "               Actual:      " & Format(MaxA, "000000") & "    -   " & " Contadores: " & I
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & AUX2 & vbCrLf
                    'MsgBox AUX2, vbExclamation
                End If
    
    
    
    
                'Actual
                I = DBLet(Adodc1.Recordset!Contado2, "N")
                AUX2 = MontaSQLFacCli(False)
                Aux = "Select max(codfaccl) from cabfact where numserie = '" & Adodc1.Recordset!tiporegi
                SQL = Aux & "' AND " & AUX2
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                MaxA = 0
                If Not miRsAux.EOF Then MaxA = DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
                If I <> MaxA Then
                    AUX2 = "               Siguiente: " & Format(MaxA, "000000") & "    -   " & " Contadores: " & I
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & AUX2 & vbCrLf
                End If
    
    
                If CadenaDesdeOtroForm <> "" Then Me.Tag = Me.Tag & "    -Registro: " & Adodc1.Recordset!tiporegi & " - " & Adodc1.Recordset!nomregis & vbCrLf & CadenaDesdeOtroForm & vbCrLf
                
            End If
            Adodc1.Recordset.MoveNext
            
        Wend
        Adodc1.Recordset.MoveFirst
        If NumRegElim > 1 Then Adodc1.Recordset.Move (NumRegElim - 1)
        If Me.Tag <> "" Then CadenaError = CadenaError & vbCrLf & "FACTURAS CLIENTES" & vbCrLf & Me.Tag
    End If
    Me.Tag = ""
    'Fin comprobacion
    Set miRsAux = Nothing
    If CadenaError = "" Then
        MsgBox "Comprobacion finalizada", vbInformation
    Else
        MsgBox CadenaError, vbExclamation
    End If
End Function

Private Function MontaSQLFacCli(Actual As Boolean) As String
Dim I As Integer
Dim F As Date
    If Actual Then
        I = 0
    Else
        I = 1
    End If
    F = DateAdd("yyyy", I, vParam.fechaini)
    MontaSQLFacCli = "fecfaccl >='" & Format(F, FormatoFecha) & "'"
    F = DateAdd("yyyy", I, vParam.fechafin)
    MontaSQLFacCli = MontaSQLFacCli & " AND fecfaccl<='" & Format(F, FormatoFecha) & "'"
End Function
