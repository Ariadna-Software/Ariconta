VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIVA 
   Caption         =   "Tipos de I.V.A."
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   Icon            =   "frmIVA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame frameSoportado 
      Caption         =   "Soportado"
      Height          =   1755
      Left            =   60
      TabIndex        =   25
      Top             =   2760
      Width           =   7515
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   1200
         Width           =   3795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   720
         Width           =   3795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   240
         Width           =   3795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   2355
         MaxLength       =   1
         TabIndex        =   9
         Tag             =   "Cta. repercutido|T|N|||tiposiva|cuentasn|||"
         Text            =   "T"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   2355
         MaxLength       =   1
         TabIndex        =   8
         Tag             =   "Cta. repercutido|T|N|||tiposiva|cuentasr|||"
         Text            =   "T"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   2355
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Cta. repercutido|T|N|||tiposiva|cuentaso|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta soportado N/Ded"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   1275
         Width           =   1530
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   1980
         Picture         =   "frmIVA.frx":030A
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta soportado recargo"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   795
         Width           =   1575
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   1980
         Picture         =   "frmIVA.frx":0D0C
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta soportado"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   330
         Width           =   990
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   1980
         Picture         =   "frmIVA.frx":170E
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Repercutido"
      Height          =   1215
      Left            =   60
      TabIndex        =   20
      Top             =   1440
      Width           =   7515
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   720
         Width           =   3795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   180
         Width           =   3795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   2340
         TabIndex        =   5
         Tag             =   "Cta. repercutido|T|N|||tiposiva|cuentare|||"
         Text            =   "Text1"
         Top             =   195
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Cta. repercutido|T|N|||tiposiva|cuentarr|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta repercutido recargo"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   795
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. repercutido"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   285
         Width           =   1125
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   1980
         Picture         =   "frmIVA.frx":2110
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1980
         Picture         =   "frmIVA.frx":2B12
         Top             =   795
         Width           =   240
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmIVA.frx":3514
      Left            =   5760
      List            =   "frmIVA.frx":3516
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Tipo de IVA|T|N|||tiposiva|tipodiva|||"
      Top             =   990
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   4260
      TabIndex        =   2
      Tag             =   "%  IVA|N|N|0|100|tiposiva|porceiva|#0.00||"
      Text            =   "Text1"
      Top             =   990
      Width           =   630
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   4980
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "% Recargo de IVA|N|N|0|100|tiposiva|porcerec|#0.00||"
      Text            =   "Text1"
      Top             =   990
      Width           =   645
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6555
      TabIndex        =   12
      Top             =   4635
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   660
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Código del tipo de IVA|T|N||100|tiposiva|nombriva|||"
      Text            =   "Text1"
      Top             =   990
      Width           =   3525
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   45
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "Código del tipo de IVA|N|N||100|tiposiva|codigiva||S|"
      Text            =   "Text1"
      Top             =   990
      Width           =   510
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6570
      TabIndex        =   11
      Top             =   4635
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5250
      TabIndex        =   10
      Top             =   4635
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   150
      Top             =   4680
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
      TabIndex        =   32
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
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
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
         Left            =   6120
         TabIndex        =   33
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de IVA"
      Height          =   195
      Left            =   5880
      TabIndex        =   19
      Top             =   750
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "% I.V.A."
      Height          =   195
      Index           =   4
      Left            =   4200
      TabIndex        =   18
      Top             =   750
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "% Recargo"
      Height          =   195
      Index           =   2
      Left            =   4920
      TabIndex        =   17
      Top             =   750
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   255
      Index           =   1
      Left            =   645
      TabIndex        =   16
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private DevfrmCCtas As String

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                'MsgBox "Registro insertado.", vbInformation
                PonerModo 0
                lblIndicador.Caption = ""
            End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                
                    InsertaEnElLog
                
                    TerminaBloquear
                    lblIndicador.Caption = ""
                    If SituarData1 Then
                        PonerModo 2
                    Else
                        LimpiarCampos
                        PonerModo 0
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

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1, 3
    LimpiarCampos
    PonerModo 0
Case 4
    'Modificar
    lblIndicador.Caption = ""
    TerminaBloquear
    PonerModo 2
    PonerCampos
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
            Data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = " codigiva = " & Text1(0).Text & ""
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
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        '### A mano
        '-------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
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
    Me.lblIndicador.Caption = "Leyendo datos"
    Me.lblIndicador.Refresh
    LimpiarCampos
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
lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
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
    
    'Comprobamos si se puede eliminar
    If Not SePuedeEliminar Then Exit Sub
    '### a mano
    Cad = "Seguro que desea eliminar de la BD el registro:"
    Cad = Cad & vbCrLf & "Tipo iva: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Descripción: " & Data1.Recordset.Fields(1)
    I = MsgBox(Cad, vbQuestion + vbYesNo)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
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
        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String
    
    If Not Data1.Recordset.EOF Then
    
        If Text1(0).Text <> "" Then
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
            RaiseEvent DatoSeleccionado(Cad)
        
        End If
    End If
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Combo1.BackColor = vbWhite
      SendKeys "{tab}"
      KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer


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


    LimpiarCampos
    'Si hay algun combo los cargamos
    CargarCombo
    
    'Como son cuentas, como mucho seran
    For I = 4 To 8
        Text1(I).MaxLength = vEmpresa.DigitosUltimoNivel
    Next I
    
    '## A mano
    NombreTabla = "tiposiva"
    Ordenacion = " ORDER BY codigiva"
        
    PonerOpcionesMenu
    
    'Para todos
'    Data1.UserName = vUsu.Login
'    Me.Data1.password = vUsu.Passwd
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        Text1(0).BackColor = vbYellow
    End If

End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
    'Check1.Value = 0
End Sub


Private Sub CargarCombo()
'### tipodiva
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'       0-IVA,              1-IGIC, 2-Bien de inversión, 3.- REA(Regimen especial agrario)
'       4.- NO DEDUCIBLE    5.- Importacion
' columna: tipodiva
    Combo1.Clear
    Combo1.AddItem "IVA"
    Combo1.ItemData(Combo1.NewIndex) = 0
    
    Combo1.AddItem "IGIC"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "BIEN DE INVERSIÓN"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
    Combo1.AddItem "R.E.A."
    Combo1.ItemData(Combo1.NewIndex) = 3
    
    Combo1.AddItem "NO DEDUCIBLE"
    Combo1.ItemData(Combo1.NewIndex) = 4     'NO DEDUCIBLE
    
    'Marzo 2014
    'IVA IMPORTACION
    Combo1.AddItem "Importación"
    Combo1.ItemData(Combo1.NewIndex) = 5
    
    
    
   
End Sub


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
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
DevfrmCCtas = CadenaSeleccion
End Sub

Private Sub imgCuentas_Click(Index As Integer)
 Screen.MousePointer = vbHourglass
 Set frmCCtas = New frmColCtas
 DevfrmCCtas = ""
 frmCCtas.DatosADevolverBusqueda = "0"
 frmCCtas.Show vbModal
 Set frmCCtas = Nothing
 If DevfrmCCtas <> "" Then
    Text1(4 + Index) = RecuperaValor(DevfrmCCtas, 1)
    Text2(Index).Text = RecuperaValor(DevfrmCCtas, 2)
 End If
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


'Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        KeyCode = 0
'        SendKeys "{TAB}"
'    End If
'End Sub


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
    Dim I As Integer
    Dim SQL As String
    Dim mTag As CTag
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
    
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 2, 3
            If Modo = 3 Or Modo = 4 Then
                If Text1(Index).Text = "" Then Exit Sub
                Set mTag = New CTag
                If mTag.Cargar(Text1(Index)) Then
                    If mTag.Cargado Then
                        If mTag.Comprobar(Text1(Index)) Then
                            FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
                        Else
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
                End If
                Set mTag = Nothing
             End If
        Case 4 To 8
            If Text1(Index).Text = "" Then
                 Text2(Index - 4).Text = SQL
                 Exit Sub
            End If
            DevfrmCCtas = Text1(Index).Text
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                Text1(Index).Text = DevfrmCCtas
                Text2(Index - 4).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                Text1(Index).Text = ""
                Text2(Index - 4).Text = ""
                PonerFoco Text1(Index)
            End If
            DevfrmCCtas = ""
        '....
    End Select
    '---
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
        Cad = Cad & ParaGrid(Text1(0), 10, "Código")
        Cad = Cad & ParaGrid(Text1(1), 60, "Denominacion")
        Cad = Cad & ParaGrid(Text1(2), 20, "% IVA ")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.VCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Tipos de I.V.A."
            frmB.vSelElem = 0
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
    Dim I As Integer
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    PonerCtasIVA
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Dim B As Boolean
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For I = 0 To Text1.Count - 1
            'Text1(i).BackColor = vbWhite
            Text1(0).BackColor = &H80000018
        Next I
        'chkVistaPrevia.Visible = False
    End If
    Modo = Kmodo
    'chkVistaPrevia.Visible = (Modo = 1)
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    'Modificar
    Toolbar1.Buttons(7).Enabled = B And vUsu.Nivel < 2
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B And vUsu.Nivel < 2
    mnEliminar.Enabled = B
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If Modo < 3 Then
        If DatosADevolverBusqueda <> "" Then
            cmdRegresar.visible = B Or Modo = 0
            cmdRegresar.Cancel = True
        Else
            cmdRegresar.visible = False
        End If
    End If
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = B Or Modo = 1
    cmdCancelar.visible = B Or Modo = 1
    mnOpciones.Enabled = Not B
    If B Or Modo = 1 Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 2
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    B = (Modo = 2) Or Modo = 0
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = B
        If B Then
            Text1(I).BackColor = &H80000018
        ElseIf Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    For I = 0 To imgCuentas.Count - 1
        imgCuentas(I).Enabled = Not B
    Next I
    Combo1.Enabled = Not B
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    'Comprobamos  si existe
    If Modo = 3 Then
        If DevuelveDesdeBD("codigiva", "tiposiva", "codigiva", Text1(0).Text, "N") <> "" Then
            B = False
            MsgBox "Ya existe el codigo de IVA: " & Text1(0).Text, vbExclamation
        Else
            B = True
        End If
    End If
    DatosOk = B
End Function


'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()

    Dim SQL As String
    Dim RS As ADODB.Recordset

    SQL = "Select Max(codigiva) from " & NombreTabla
    Text1(0).Text = 1
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            Text1(0).Text = RS.Fields(0) + 1
        End If
    End If
    RS.Close
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
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
Case 8
    BotonEliminar
Case 12
    mnSalir_Click
Case 14 To 17
    Desplazamiento (Button.Index - 14)
'Case 20
'    'Listado en crystal report
'    Screen.MousePointer = vbHourglass
'    CR1.Connect = Conn
'    CR1.ReportFileName = App.Path & "\Informes\list_Inc.rpt"
'    CR1.WindowTitle = "Listado incidencias."
'    CR1.WindowState = crptMaximized
'    CR1.Action = 1
'    Screen.MousePointer = vbDefault

Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).visible = bol
    Next I
End Sub


'Private Sub PonerCtasIVA()
'Dim SQL As String
'Dim RS As Recordset
'Dim i As Integer
'On Error GoTo EPonerCtasIVA
'
'SQL = "Select codmacta,nommacta FROM cuentas WHERE apudirec='S'"
'Set RS = New ADODB.Recordset
'RS.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
'
'If RS.EOF Then
'    RS.Close
'    Exit Sub
'End If
'
'
'For i = 0 To 4
'    RS.MoveFirst
'    RS.Find "codmacta = '" & Text1(i + 4).Text & "'", , adSearchForward
'    If Not RS.EOF Then
'        Text2(i).Text = RS!nommacta
'        Else
'        Text2(i).Text = "No encontrado"
'    End If
'Next i
'RS.Close
'Set RS = Nothing
'Exit Sub
'EPonerCtasIVA:
'    MuestraError Err.Number, "Poniendo valores ctas. IVA", Err.Description
'End Sub




Private Sub PonerCtasIVA()
Dim SQL As String
Dim I As Integer
On Error GoTo EPonerCtasIVA

'SQL = "Select codmacta,nommacta FROM cuentas WHERE apudirec='S'"
'Set RS = New ADODB.Recordset
'RS.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
'
'If RS.EOF Then
'    RS.Close
'    Exit Sub
'End If
'

For I = 0 To 4
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(I + 4).Text, "T")
    
    Text2(I).Text = SQL
Next I
Exit Sub
EPonerCtasIVA:
    MuestraError Err.Number, "Poniendo valores ctas. IVA", Err.Description
End Sub



Private Sub PonerFoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SePuedeEliminar() As Boolean
Dim B As Boolean
Dim Cad As String

    Screen.MousePointer = vbHourglass
    SePuedeEliminar = False
    Set miRsAux = New ADODB.Recordset
    Cad = "Select fecfaccl from cabfact where tp1faccl =" & Text1(0).Text
    Cad = Cad & " OR tp2faccl =" & Text1(0).Text
    Cad = Cad & " OR tp3faccl =" & Text1(0).Text & ";"
    B = True
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then B = False
    miRsAux.Close
    
    If B Then
            Cad = "Select fecfacpr from cabfactprov where tp1facpr =" & Text1(0).Text
            Cad = Cad & " OR tp2facpr =" & Text1(0).Text
            Cad = Cad & " OR tp3facpr =" & Text1(0).Text & ";"
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then B = False
            miRsAux.Close
            SePuedeEliminar = B
    End If
    If Not B Then MsgBox "El tipo de IVA no se puede eliminar. Hay facturas vinculadas con ese tipo de IVA.", vbExclamation
    Screen.MousePointer = vbDefault
End Function



Private Sub InsertaEnElLog()

On Error GoTo eInsertaEnElLog

    CadenaConsulta = ""
    If Text1(1).Text <> Me.Data1.Recordset!nombriva Then
        DevfrmCCtas = Text1(1).Text & " -- " & Me.Data1.Recordset!nombriva
        CadenaConsulta = DevfrmCCtas
    End If
    If ImporteFormateado(Text1(2).Text) <> DBLet(Me.Data1.Recordset!porceiva, "N") Then
        DevfrmCCtas = "(1)" & Text1(2).Text & " -- " & Me.Data1.Recordset!nombriva
        CadenaConsulta = CadenaConsulta & DevfrmCCtas
    End If
        
    If ImporteFormateado(Text1(3).Text) <> DBLet(Me.Data1.Recordset!porcerec, "N") Then
        DevfrmCCtas = "(2)" & Text1(3).Text & " -- " & Me.Data1.Recordset!porcerec
        CadenaConsulta = CadenaConsulta & DevfrmCCtas
    End If
    
    'El combo
    If Me.Combo1.ItemData(Me.Combo1.ListIndex) <> Val(Me.Data1.Recordset!tipodiva) Then
        DevfrmCCtas = "(Tipo)" & Me.Combo1.ItemData(Me.Combo1.ListIndex) & " -- " & Val(Me.Data1.Recordset!tipodiva)
        CadenaConsulta = CadenaConsulta & DevfrmCCtas
    End If
    
    For kCampo = 0 To 4
        If Text1(kCampo + 4).Text <> DBLet(Data1.Recordset.Fields(kCampo + 5), "T") Then
            DevfrmCCtas = "(Cta " & kCampo & ")" & Text1(3).Text & " -- " & DBLet(Data1.Recordset.Fields(kCampo + 5), "T")
            CadenaConsulta = CadenaConsulta & DevfrmCCtas
        End If
    Next
    If CadenaConsulta <> "" Then
        CadenaConsulta = "Actual/Anterior" & vbCrLf & CadenaConsulta
        vLog.Insertar 20, vUsu, CadenaConsulta
        CadenaConsulta = ""
    End If

eInsertaEnElLog:
    If Err.Number <> 0 Then MuestraError Err.Number, "Acciones realizadas(SLOG)", Err.Description
        

End Sub
