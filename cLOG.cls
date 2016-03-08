VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ACCIONES DISPONIBLES
'
'
'

'
'       1.- Crear asiento manual
'       2.- Modificar asiento
'       3.- Eliminar asiento

'       4,5,6   FRACLI
'       7,8,9   FRAPRO

'       10.- Actualizar asientos
'       11.- Contabilizar FRACLI
'       12.- Contabilizar FRAPRO

'       13.- Realizar amortizacion
'       14.- Deshacer  " "
'       15.- 340 a fichero
'       16.- Acciones del CC. En la descripcioin IRA lo que he heco
'       17.- Punteo de extractos
'       18.-  proceso cierre [RENUMERAR] [CIERRE]  [DESHACER]

'       19.-  VARIOS.  Para quedar constancia pero no crear un "exproceso"


'       20.- Modificar tabla IVA

'        TESORERIA
'       101.- cobros agentes

Private Const AccionesRegistradas = 20
Private mvarDatosDescripcion As String

Dim SQL As String



Public Property Let DatosDescripcion(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarDatosDescripcion = vData
End Property


Public Property Get DatosDescripcion() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    DatosDescripcion = mvarDatosDescripcion
End Property



Public Function Insertar(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String) As Boolean


    On Error GoTo EI
    Insertar = False
    
    SQL = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
    SQL = SQL & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    SQL = SQL & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        SQL = SQL & "NULL)"
    Else
        If Len(Descripcion) > 255 Then Descripcion = Mid(Descripcion, 1, 250) & "[MAS]"
        SQL = SQL & "'" & DevNombreSQL(Descripcion) & "')"
    End If
    Conn.Execute SQL
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, "Registro LOG de acciones"

End Function


Private Function DimeAccion(N As Integer) As String
    If N > 100 Then
        'Tesoreria
        DimeAccion = DimeAccionTesoreria(N)
    Else
        DimeAccion = DimeAccion2(N)
    End If
End Function

Private Function DimeAccion2(N As Integer) As String
Dim H As Integer
Dim C As String

    If N <= 9 Then
        'las 9 primeras son mantenimientos basicos
        H = N Mod 3
        If H = 1 Then
            C = "NEW"
        ElseIf H = 2 Then
            C = "MOD"
        Else
            C = "DEL"
        End If
        H = (N - 1) \ 3
        If H = 0 Then
            C = C & " ASTO"
        ElseIf H = 1 Then
            C = C & " FRACLI"
        Else
            C = C & " FRAPRO"
        End If
            
    Else
        'El resto de acciones con un select case
        'ACTUALIZACION (CONTABILIZACION)
        If N <= 12 Then
            C = "CONTAB "
            If N = 10 Then
                C = C & " ASTO"
            ElseIf N = 11 Then
                C = C & " FRACLI"
            Else
                C = C & " FRAPRO"
            End If
            
            
            
            
        Else
            'Aqui iran yendo resto acciones
            If N <= 16 Then
                Select Case N
                Case 13
                    C = "AMORTIZ."
                Case 14
                    C = "DESAMORT."
                Case 15
                    C = "Mod340"
                Case 16
                    C = "ControlCC"
                End Select
        
            Else
                'n=17
                If N = 17 Then
                    C = "Punteo cta"
                ElseIf N = 18 Then
                    'Llevara la primera linea como que se ha hecho(
                    '[CIERRE-RENUMERAR-DESHACER}
                    C = "PROCESO CIERRE"
                    
                ElseIf N = 19 Then
                    C = "VARIOS"
                    
                Else
                    C = "Mod. IVA"
                End If
            End If
        End If
        
        
        
    End If
    DimeAccion2 = C
End Function


Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
Dim I As Integer

 On Error GoTo EDevuelveAcciones
    For I = 1 To AccionesRegistradas
        SQL = I & "|" & DimeAccion(I) & "|"
        ListaAcciones.Add SQL
    Next I
    
    DevuelveAcciones = True
    
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function


'`slog_old`
'*****************************************************************
'
'   Ahora ya no volcamos a feichero. Meteremos en una tabla myisam, sin claves ni na de na.
'
'Public Sub VolcarAFichero()
'Dim F As Date
'Dim RS As ADODB.Recordset
'Dim NF As Integer
'
'Dim Blanco As String
'
'
'    On Error GoTo EVolcar
'
'
'
'    Set RS = New ADODB.Recordset
'    F = Now   'Por si acaso esta el reloj mal
'    RS.Open "select curdate()", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then F = DBLet(RS.Fields(0), "F")
'    RS.Close
'
'    'Ya tengo la fecha.
'    F = DateAdd("m", -1, F)  'Le quito 1 meses
'    F = CDate("01/" & Month(F) & "/" & Year(F))   'Pongo el primer dia del mes anterior
'    'Le quito un dia
'    F = DateAdd("d", -1, F)
'    If Dir(App.path & "\LOG", vbDirectory) = "" Then MkDir App.path & "\LOG"
'
'    For NF = 1 To 99
'        SQL = Format(vEmpresa.codempre, "00") & "_" & Format(F, "yyyymm") & "_" & Format(NF, 0) & ".log"
'        If Dir(App.path & "\LOG\" & SQL, vbArchive) = "" Then
'            'ESTE ES EL NOMBRE
'            Exit For
'        End If
'    Next
'
'    RS.Open "Select * from slog where fecha < '" & Format(F, "yyyy-mm-dd") & "' ORDER BY fecha", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        '-------------------------------------------
'        Blanco = Space(25)
'        NF = FreeFile
'        Open App.path & "\LOG\" & SQL For Output As NF
'        While Not RS.EOF
'            'Una linea para cada accion. Separadas por ; por si hay que abrirlo en excel
'            SQL = Mid(DimeAccion(RS!Accion) & Blanco, 1, 15)
'            SQL = Format(RS!Fecha, "dd/mm/yyyy hh:mm:ss") & ";" & SQL & ";"
'            SQL = SQL & Mid(RS!Usuario & Blanco, 1, 30) & ";" & Mid(RS!PC & Blanco, 1, 30) & ";'" & RS!Descripcion & "';"
'            Print #NF, SQL
'            RS.MoveNext
'        Wend
'
'        Close #NF
'        Conn.Execute "DELETE from slog where fecha <='" & Format(F, "yyyy-mm-dd") & "'"
'
'    End If
'    RS.Close
'    Set RS = Nothing
'
'
'    Exit Sub
'EVolcar:
'    Err.Clear
'    Set RS = Nothing
'End Sub


Public Sub VolcarAFichero2()
Dim F As Date
Dim RS As ADODB.Recordset




    On Error GoTo EVolcar

    

    Set RS = New ADODB.Recordset
    F = Now   'Por si acaso esta el reloj mal
    RS.Open "select curdate()", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then F = DBLet(RS.Fields(0), "F")
    RS.Close
    
    
    
    
    'Ya tengo la fecha.
    F = DateAdd("m", -1, F)  'Le quito 3 meses
    F = CDate("01/" & Month(F) & "/" & Year(F))   'Pongo el primer dia del mes anterior
    'Le quito un dia
    F = DateAdd("d", -1, F)
    
    
    SQL = "Select * from slog where fecha < '" & Format(F, "yyyy-mm-dd") & "' "
    SQL = "INSERT INTO slog_old " & SQL
    Conn.Execute SQL
    Conn.Execute "DELETE from slog where fecha <'" & Format(F, "yyyy-mm-dd") & "'"
        
    
    
    Set RS = Nothing
    
    
    Exit Sub
EVolcar:
    SQL = "Error volcado datos acciones: " & vbCrLf & Err.Description
    SQL = SQL & vbCrLf & vbCrLf & "El programa continuará." & vbCrLf & vbCrLf & "Avise a soporte técnico."
    MsgBox SQL, vbExclamation
    Err.Clear
    Set RS = Nothing
End Sub





'Dos metodos sobre la variable mvarDatosDescripcion
' Inicializar y añadir texto
Public Sub InicializarDatosDesc()
    mvarDatosDescripcion = ""
End Sub


Public Sub AnyadeTextoDatosDes(T As String)
    
    If mvarDatosDescripcion <> "" Then mvarDatosDescripcion = mvarDatosDescripcion & ":"
    mvarDatosDescripcion = mvarDatosDescripcion & T
End Sub


Private Function DimeAccionTesoreria(N As Integer) As String
    Select Case N
    Case 101
        DimeAccionTesoreria = "Cobros agentes"
    End Select
End Function
