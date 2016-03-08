VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacLinAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lineas factura predefinidos"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7635
   ClipControls    =   0   'False
   Icon            =   "frmFacLinAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FDEFAC&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
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
      Height          =   320
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "existencia"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   6435
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   6600
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   6600
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1320
      Top             =   6720
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacLinAdd.frx":000C
      Height          =   5895
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Falta"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
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
      Left            =   240
      TabIndex        =   4
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu mnOpciones1 
         Caption         =   "Leer desde predefinido"
         Index           =   0
      End
      Begin VB.Menu mnOpciones1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnOpciones1 
         Caption         =   "Salir"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmFacLinAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TotalLineas As Currency


Private WithEvents frmPre As frmAsiPre
Attribute frmPre.VB_VarHelpID = -1
Private Modo As Byte
Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
Dim PrimeraVez As Boolean
Dim Cad As String


Private Sub cmdAceptar_Click()
  'TotalLineas llevo
    Set miRsAux = New ADODB.Recordset
    Cad = "Select count(*),sum(saldo) from tmpconext WHERE saldo<>0 and codusu = " & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        Cad = "X"
    Else
        If DBLet(miRsAux.Fields(0), "N") = 0 Then
            Cad = "X"
        Else
            If miRsAux.Fields(1) <> CCur(Text1.Tag) Then
                Cad = "Importe BD y variable incorrectos. Avise soporte tecnico"
            Else
                If miRsAux.Fields(1) <> TotalLineas Then
                    Cad = "La suma deberia ser : " & TotalLineas & "     Diferencia: " & Text1.Text
                Else
                    Cad = "Va a insertar en la factura:" & vbCrLf
                    Cad = Cad & Space(10) & "Lineas: " & miRsAux.Fields(0) & vbCrLf
                    Cad = Cad & Space(10) & "Importe: " & Format(miRsAux.Fields(1), FormatoImporte)
                    Cad = Cad & vbCrLf & vbCrLf & "¿Continuar?"
                    If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                        Cad = ""
                    Else
                        Cad = "NO"
                    End If
                End If
            End If
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Cad <> "" Then
        If Cad = "X" Then Cad = "No hay valor para ninguna de las lineas"
        
        If Cad <> "NO" Then MsgBox Cad, vbExclamation
        
        Exit Sub
    End If
    
        
        
        
        
        
    CadenaDesdeOtroForm = "OK"
    Me.Tag = 0
    Unload Me
    
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
   
    
    If MsgBox("Desea cancelar el proceso?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    CadenaDesdeOtroForm = ""
    Me.Tag = 0
    Unload Me
  
   
End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
       'txtAux.SelStart = Len(Me.txtAux.Text)
       
       txtAux.SetFocus
       
       txtAux.SelStart = 0
       txtAux.SelLength = Len(Me.txtAux.Text)
       txtAux.Refresh
       
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
         PrimeraVez = False
         mnOpciones1_Click 0
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
 
    
    Me.Tag = 1 'NO se puede cerrar mas que de boton
    
    
    If vParam.autocoste Then
        Me.Width = 10315
        
    End If
    DataGrid1.Width = Me.Width - 400
    Me.cmdCancelar.Left = Me.Width - 1365
    Me.cmdAceptar.Left = Me.Width - 2565
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    BorrarDatos
    
    CargaGrid
    

    Text1.Left = DataGrid1.Columns(5).Left + 130 'codalmac
    Text1.Width = DataGrid1.Columns(5).Width - 10
    Text1.Text = Format(TotalLineas, FormatoImporte)
    Text1.Tag = 0
    
    Label1.Left = Text1.Left - Label1.Width - 120
    
    
    PrimeraVez = True
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid()


On Error GoTo ECarga

    gridCargado = False
    
    Cad = "select cta,nommacta, pos ,ccost ,ampconce,saldo"
    Cad = Cad & " from tmpconext,cuentas  where tmpconext.cta=cuentas.codmacta AND codusu=" & vUsu.Codigo & " ORDER BY pos"

    data1.ConnectionString = Conn
    data1.RecordSource = Cad
    data1.CursorType = adOpenDynamic
    data1.LockType = adLockPessimistic
    data1.Refresh
   
    
    PrimeraVez = False
        

    
    DataGrid1.Columns(0).Caption = "Cuenta"
    DataGrid1.Columns(0).Width = 1100
    
    
        
    DataGrid1.Columns(1).Caption = "Descripcion"
    DataGrid1.Columns(1).Width = 3700
    

    DataGrid1.Columns(2).visible = False
    
    If Not vParam.autocoste Then
        DataGrid1.Columns(3).visible = False
        DataGrid1.Columns(4).visible = False
        
    Else
        DataGrid1.Columns(3).Caption = "C.C."
        DataGrid1.Columns(3).Width = 700
        DataGrid1.Columns(4).Caption = "Nombre centro coste"
        DataGrid1.Columns(4).Width = 2300
    End If
    
    
    DataGrid1.Columns(5).Caption = "Base"
    DataGrid1.Columns(5).Width = 1250
    DataGrid1.Columns(5).NumberFormat = FormatoImporte
    DataGrid1.Columns(5).Alignment = dbgRight
            
    
    
    DataGrid1.ScrollBars = dbgAutomatic
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        txtAux.visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
                txtAux.Text = data1.Recordset!Saldo
                txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(5).Left + 130 'codalmac
        txtAux.Width = DataGrid1.Columns(5).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux.visible = visible
    End If
    PonFoco txtAux
    
    If visible Then
        txtAux.TabIndex = 2
      '  txtAux.SelStart = 0
       ' txtAux.SelLength = Len(txtAux.Text)
    Else
        txtAux.TabIndex = 5
    End If
End Sub






Private Sub Form_Unload(Cancel As Integer)
    If Me.Tag = 1 Then Cancel = 1 'o aceptar o cancelar
End Sub

Private Sub frmPre_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub mnOpciones1_Click(Index As Integer)
    If Index = 0 Then
        
        If data1.Recordset.RecordCount > 0 Then
            If MsgBox("Ya existen datos. Volver a cargarlos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            BorrarDatos
            espera 0.5
        End If
        Cad = ""
        Set frmPre = New frmAsiPre
        frmPre.DatosADevolverBusqueda = "0|1|"
        frmPre.Show vbModal
        Set frmPre = Nothing
        If Cad <> "" Then
            
            Cad = " WHERE numaspre=" & RecuperaValor(Cad, 1)
            
            If vParam.autocoste Then
                'LLEVA CENTROS DE COSTE
                Cad = " left join cabccost ON linasipre.codccost =cabccost.codccost  " & Cad
            End If
            Cad = " FROM linasipre " & Cad
            
            If vParam.autocoste Then Cad = ",linasipre.codccost,nomccost" & Cad
            
            
            Cad = ")  select " & vUsu.Codigo & ",codmacta,0,linlapre" & Cad
            If vParam.autocoste Then Cad = ",ccost, ampconce" & Cad
            Cad = "INSERT INTO tmpconext(codusu,cta,  saldo,pos " & Cad
            Conn.Execute Cad
            CargaGrid
            BotonModificar
        End If
    
    
    Else
    
    
    End If
    
End Sub

Private Sub txtAux_GotFocus()
    txtAux.SelStart = 0
    txtAux.SelLength = Len(txtAux.Text)
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If
        
        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                     Me.txtAux.SelStart = 0
                Me.txtAux.SelLength = Len(Me.txtAux.Text)
                'txtaux.Refresh
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        ModificarExistencia
        
        PasarSigReg
        
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus()
Dim Importe As Currency
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then
            .Text = "0,00"
        Else
                If Not EsNumerico(.Text) Then
                    MsgBox "Importes deben ser numéricos.", vbExclamation
                    On Error Resume Next
                    .Text = "0,00"
                    PonFoco txtAux
                    Exit Sub
                End If
                
                
                'Es numerico
                Cad = TransformaPuntosComas(.Text)
                If CadenaCurrency(Cad, Importe) Then .Text = Format(Importe, "0.00")
                    
                
        
        End If
    End With

End Sub






Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim B As Boolean
       
    Modo = Kmodo
  
    

    B = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera B
   
    Select Case Kmodo
'    Case 0    'Modo Inicial
'        PonerBotonCabecera True
'        lblIndicador.Caption = ""
        
    Case 1 'Modo Buscar
'        PonerBotonCabecera False
      
'        lblIndicador.Caption = "BÚSQUEDA"
'    Case 2    'Visualización de Datos
'        PonerBotonCabecera True
'    Case 3 'Insertar Datos en el Datagrid
'        PonerBotonCabecera False 'Poner Aceptar y Cancelar Visible
'        lblIndicador.Caption = "MODIFICAR"
    End Select

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub







Private Sub BotonModificar()
    If data1.Recordset.EOF Then Exit Sub
    PonerModo 4
    CargaTxtAux True, True
    PonleFoco txtAux
End Sub


Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
    txtAux.Text = Trim(txtAux.Text)
    DatosOk = False
    If txtAux.Text <> "" Then
        If EsNumerico(txtAux.Text) Then DatosOk = True
    End If
End Function


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    'PonerOpcionesMenuGeneral Me
End Sub




Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        PonleFoco Me.txtAux
    ElseIf DataGrid1.Bookmark = data1.Recordset.RecordCount Then
       PonleFoco cmdAceptar
    End If
    

End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long


    If DatosOk Then
        
        If ActualizarExistencia() Then
            
            NumReg = data1.Recordset.AbsolutePosition
            CargaGrid
            
                    
            If NumReg < data1.Recordset.RecordCount Then
                data1.Recordset.Move NumReg - 1
            Else
                data1.Recordset.MoveLast
            End If
        End If

            
            
            ModificarExistencia = True
    Else
            ModificarExistencia = False
  
    End If
End Function




Private Function ActualizarExistencia() As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim Cantidad As Currency


    On Error GoTo EActualizar

    
        


    Cantidad = TransformaPuntosComas(txtAux.Text)
    
    If Cantidad < 0 Then Err.Raise 513, , "No se permiten negativos"
        
    If Cantidad <> data1.Recordset!Saldo Then
            
        'Actualizar la Tabla: sinven con la cantidad introducida
        '-------------------------------------------------------
'
        SQL = "UPDATE tmpconext  Set saldo = " & TransformaComasPuntos(CStr(Cantidad))
        SQL = SQL & " WHERE cta = '" & data1.Recordset!Cta & "' AND "
        SQL = SQL & " pos =" & data1.Recordset!Pos & " AND codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        
        Text1.Tag = Text1.Tag - data1.Recordset!Saldo + Cantidad
        Cantidad = TotalLineas - CCur(Text1.Tag)
        Text1.Text = Format(Cantidad, FormatoImporte)
        
    End If
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         MuestraError Err.Number, SQL, Err.Description
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
    End If
End Function


Private Sub BorrarDatos()
    Conn.Execute "DELETE FROM tmpconext WHERE codusu = " & vUsu.Codigo
End Sub
