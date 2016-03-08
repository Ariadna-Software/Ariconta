VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIntracom349 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar tipo factura intracomunitaria"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9840
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "Asignar"
      Height          =   315
      Left            =   6240
      TabIndex        =   4
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   11280
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6840
      Width           =   4695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NIF"
         Object.Width           =   3387
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5211
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nº Fact"
         Object.Width           =   2884
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Importe"
         Object.Width           =   2532
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tipo"
         Object.Width           =   4180
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Clave operacion"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "frmIntracom349"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String
    
Private Sub Command2_Click()

End Sub

Private Sub cmdAceptar_Click()

    If ListView1.Tag = 1 Then
        Screen.MousePointer = vbHourglass
        ACtualizaDatosparaFichero
        Screen.MousePointer = vbDefault
        
    End If
    Unload Me
End Sub

Private Sub cmdAsignar_Click()
Dim primerSel As Integer
Dim ParaClientes As Boolean

    If Combo2.ListIndex < 0 Then Exit Sub

    CadenaDesdeOtroForm = ""   'Veremos si ha seleccionado de dos tipos distintos (entragas/adquisiciones)
    SQL = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        With ListView1.ListItems(NumRegElim)
            If .Selected Then
                SQL = SQL & "X"
                ' T.SmallIcon = 16  o 18
                If .SmallIcon = 16 Then
                    If InStr(1, CadenaDesdeOtroForm, "E") = 0 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "E"
                Else
                    If InStr(1, CadenaDesdeOtroForm, "A") = 0 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "A"
                End If
                    
            End If
        End With
    Next
    
    
    If SQL = "" Then Exit Sub
    If Len(CadenaDesdeOtroForm) = 2 Then
        MsgBox "No puede asignar a distintos tipos de facturas", vbExclamation
        Exit Sub
    End If
    
    'Ahora comprobaremos que
    'Para PROVEEDOR se le puede asignar
    '
    'Para cliente
    '
    ' 1 y 3 es proveedores
    ParaClientes = Not (Combo2.ItemData(Combo2.ListIndex) = 1 Or Combo2.ItemData(Combo2.ListIndex) = 3)
    
    If CadenaDesdeOtroForm = "E" Then
        'El combo puede ser Entragas  o prestacion servicios
        If Not ParaClientes Then
            MsgBox "Clave de operaciones para proveedor", vbExclamation
            Exit Sub
        End If
    Else
        If ParaClientes Then
            MsgBox "Clave de operaciones para clientes", vbExclamation
            Exit Sub
        End If
    End If
    
    
    CadenaDesdeOtroForm = "Va a asignar la clave de operacion " & vbCrLf & vbCrLf & "     " & UCase(Combo2.List(Combo2.ListIndex)) & vbCrLf & vbCrLf
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Total facturas seleccionadas: " & Len(SQL)
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & vbCrLf
    If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) <> vbYes Then
        CadenaDesdeOtroForm = ""
    Else
        'Ha dicho que si
        Screen.MousePointer = vbHourglass
        If Not ListView1.SelectedItem Is Nothing Then
            primerSel = ListView1.SelectedItem.Index
        Else
            primerSel = 1
        End If
        
        For NumRegElim = 1 To ListView1.ListItems.Count
            With ListView1.ListItems(NumRegElim)
                If .Selected Then
        
                    SQL = "UPDATE usuarios.ztesoreriacomun SET opcion = " & Combo2.ItemData(Combo2.ListIndex)
                    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND codigo =" & .Tag
                    Conn.Execute SQL
                    
                    ListView1.Tag = 1
                End If
            End With
        Next
        Combo2.ListIndex = -1
        CargarDatos
        If primerSel <= ListView1.ListItems.Count Then
            Set ListView1.SelectedItem = ListView1.ListItems(primerSel)
            ListView1.SelectedItem.EnsureVisible
        End If
        CadenaDesdeOtroForm = ""
        Screen.MousePointer = vbDefault
    End If
    
    
End Sub

Private Sub Command1_Click()
'        For NumRegElim = 1 To ListView1.ColumnHeaders.Count
'            Debug.Print ListView1.ColumnHeaders(NumRegElim).Text & ":"; ListView1.ColumnHeaders(NumRegElim).Width
'        Next
'        Stop

    Conn.Execute "DElete FROM Usuarios.z347 where codusu = " & vUsu.Codigo
    espera 0.3
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag = 0 Then
        Me.Tag = 1
        
        'Para saber si ha cambiado algo. Si no toca nada, no recalcularemos
        ListView1.Tag = 0
        
        CargarDatos
        

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon
    Me.Tag = 0
    
    Me.ListView1.SmallIcons = frmPpal.ImageList1
    


    SQL = DevuelveCadenaTipoFacturas349_
    Combo2.Clear
    For NumRegElim = 0 To 6
        'Debug.Print RecuperaValor(SQL, NumRegElim + 1)
        Combo2.AddItem RecuperaValor(SQL, NumRegElim + 1)
        Combo2.ItemData(Combo2.NewIndex) = NumRegElim
    Next
    
    
End Sub



Private Sub CargarDatos()
Dim IT As ListItem
On Error GoTo eCargaDatos


    ListView1.ListItems.Clear

    SQL = "select * from usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " ORDER BY texto3,texto4"
    Set miRsAux = New ADODB.Recordset
    
    CadenaDesdeOtroForm = DevuelveCadenaTipoFacturas349_
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = miRsAux!texto6 & miRsAux!texto3  'nif
        IT.SubItems(1) = miRsAux!texto5 'nomnbre
        IT.SubItems(2) = miRsAux!texto4  'nº fac
        IT.SubItems(3) = miRsAux!fecha1   'fecha
        IT.SubItems(4) = Format(miRsAux!Importe1, FormatoImporte) 'importe
        If Val(miRsAux!importe2) = 0 Then
            IT.SmallIcon = 16
        Else
            IT.SmallIcon = 18
        End If
    
    
        '   0   ventas o entregas
        '   1   compras o adquisiciones


        Select Case miRsAux!Opcion
        Case 1
            SQL = "Adquisiciones"
        Case 2, 3, 4, 5, 6
            SQL = RecuperaValor(CadenaDesdeOtroForm, miRsAux!Opcion + 1)
        Case Else
            'EL CERO
            SQL = "Entregas"
        End Select
        IT.SubItems(5) = Mid(SQL, 1, 13)
        IT.Tag = miRsAux!Codigo
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    CadenaDesdeOtroForm = ""
    
    
    
eCargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number
    
    Set miRsAux = Nothing
End Sub



Private Sub ACtualizaDatosparaFichero()
Dim RT As ADODB.Recordset
Dim Importe As Currency
Dim ImporteAComparar As Currency
Dim Aux As String
Dim Actualizar As Boolean

    On Error GoTo eACtualizaDatosparaFichero
    
    Set RT = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset

    SQL = "select * from Usuarios.z347 WHERE codusu = " & vUsu.Codigo & " ORDER BY NIF"
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    While Not RT.EOF
    
        SQL = RT!NIF
        ImporteAComparar = RT!Importe
        RT.MoveNext
        If Not RT.EOF Then
            If RT!NIF = SQL Then
                'Es el mismo NIF que actua como cliente y proveedor
                ImporteAComparar = ImporteAComparar + RT!Importe
            Else
                RT.MovePrevious
            End If
        Else
            RT.MovePrevious
        End If
        'Para cada NIF veremos si ha cambiado datos
        SQL = "select opcion,sum(importe1) from usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
        SQL = SQL & " and texto3='" & DevNombreSQL(RT!NIF) & "' group by 1"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'codusu,nif,razosoci,dirdatos,codposta,despobla,Provincia,pais,cliprov,importe,
        SQL = ", (" & vUsu.Codigo & ",'" & DevNombreSQL(RT!NIF) & "','" & DevNombreSQL(RT!razosoci) & "','"
        SQL = SQL & DevNombreSQL(DBLet(RT!dirdatos, "T")) & "','" & DevNombreSQL(DBLet(RT!codposta, "T")) & "','"
        SQL = SQL & DevNombreSQL(DBLet(RT!despobla, "T")) & "','" & DevNombreSQL(DBLet(RT!Provincia, "T")) & "','"
        SQL = SQL & DevNombreSQL(DBLet(RT!Pais, "T")) & "',"
        
        NumRegElim = 0
        Importe = 0
        Aux = "" 'preparamos el insert
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            Importe = Importe + miRsAux.Fields(1)
            
            'codusu,nif,razosoci,dirdatos,codposta,despobla,Provincia,pais,cliprov,importe,
            Aux = Aux & SQL & miRsAux.Fields(0) & "," & TransformaComasPuntos(CStr(miRsAux.Fields(1))) & ")"
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        If ImporteAComparar <> Importe Then
            MsgBox "Importe origen: " & ImporteAComparar & vbCrLf & "Suma: " & Importe, vbCritical
        End If
        
        
        SQL = "DELETE FROM Usuarios.z347 WHERE codusu = " & vUsu.Codigo & " AND NIF = '" & DevNombreSQL(RT!NIF) & "'"
        Conn.Execute SQL
        espera 0.5
        
        
        SQL = Mid(Aux, 2)
        SQL = "INSERT INTO Usuarios.z347(codusu,nif,razosoci,dirdatos,codposta,despobla,Provincia,pais,cliprov,importe) VALUES " & SQL
        Conn.Execute SQL
    
            
        
        
        RT.MoveNext
    Wend
    RT.Close
    
eACtualizaDatosparaFichero:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RT = Nothing
    Set miRsAux = Nothing
End Sub

