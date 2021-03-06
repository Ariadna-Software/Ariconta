Attribute VB_Name = "bus"
Option Explicit



Public vUsu As Usuario  'Datos usuario
Public vEmpresa As Cempresa 'Los datos de la empresa
Public vParam As Cparametros  'Los parametros
Public vConfig As Configuracion
Public vLog As cLOG   'Log de acciones

Private Procesador64bits As Boolean

'Formato de fecha
Public FormatoFecha As String
Public FormatoImporte As String

Public CadenaDesdeOtroForm As String
'Public DB As Database
Public Conn As ADODB.Connection


'Global para n� de registro eliminado
Public NumRegElim  As Long

'Para algunos campos de texto sueltos controlarlos
Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
Public AlgunAsientoActualizado As Boolean
Public TieneIntegracionesPendientes As Boolean

Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna


Public AsientoConExtModificado As Byte


'Para ver si reviso la introduccion
Public RevisarIntroduccion As Byte



'He cambiado el FechaOK. Para almacenar lo que devuelve, en algunos sitios no tengo variable
'La pongo aqui y sera comun para todos
Public varFecOk As Byte
Public Const varTxtFec = "Fecha fuera de �mbito"

    'Para los asientos k vemos desde la consulta de extractos
    '  0.- NADA
    '  1.- SIIII
Public Sub Main()


       ' MsgBox "Antes main"

       Load frmIdentifica
       CadenaDesdeOtroForm = ""

       'Necesitaremos el archivo arifon.dat
       'MsgBox "Antes identifca.show"
       frmIdentifica.Show vbModal
        
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set Conn = Nothing
            End
       End If
       
       CadenaDesdeOtroForm = ""
       frmLogin.Show vbModal
       
       If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado nonguna empresa
            Set Conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass

        'Cerramos la conexion
        Conn.Close

        
        If AbrirConexion() = False Then
            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
            End
        End If
        Screen.MousePointer = vbHourglass
        LeerEmpresaParametros
        

        RevisarIntroduccion = 0
        
        'Otras acciones
        Screen.MousePointer = vbHourglass
        
        OtrasAcciones
        
        Load frmPpal
        Screen.MousePointer = vbHourglass
        frmPpal.Show
End Sub


Public Function LeerEmpresaParametros()
        'Abrimos la empresa
        Set vEmpresa = New Cempresa
        If vEmpresa.Leer = 1 Then
            MsgBox "No se han podido cargar datos empresa. Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
           
        Set vParam = New Cparametros
        If vParam.Leer() = 1 Then
            MsgBox "No se han podido cargar los par�metros. Debe confgurar la aplicaci�n.", vbExclamation
            Set vParam = Nothing
        End If
        
        'incializamos el objeto
        Set vLog = New cLOG
        'Si estamos en localhost, y el usuario es administrador
        'Haremos la opcion de volcar la
        If LCase(vConfig.SERVER) = "localhost" And vUsu.Nivel <= 1 Then vLog.VolcarAFichero2
        
'        If Not (vEmpresa Is Nothing) And Not (vParam Is Nothing) Then
'
'            'PAra la consulta de extractos
'            CadenaDesdeOtroForm = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
'            Conn.Execute CadenaDesdeOtroForm
'
'            CadenaDesdeOtroForm = "DELETE from tmpconext where codusu= " & vUsu.Codigo
'            Conn.Execute CadenaDesdeOtroForm
'
'            CadenaDesdeOtroForm = ""
'        End If
        
        
End Function

'/////////////////////////////////////////////////////////////////
'// Se trata de identificar el PC en la BD. Asi conseguiremos tener
'// los nombres de los PC para poder asignarles un codigo
'// UNa vez asignado el codigo  se lo sumaremos (x 1000) al codusu
'// con lo cual el usuario sera distinto( aunque sea con el mismo codigo de entrada)
'// dependiendo desde k PC trabaje

Public Function GestionaPC2() As Integer
CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    FormatoFecha = DevuelveDesdeBD("codpc", "Usuarios.pcs", "nompc", CadenaDesdeOtroForm, "T")
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from Usuarios.pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        If NumRegElim > 9999 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte t�cnico.", vbCritical
            End
        End If
        FormatoFecha = "INSERT INTO Usuarios.pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        Conn.Execute FormatoFecha
    Else
        NumRegElim = Val(FormatoFecha)
    End If
    GestionaPC2 = NumRegElim
    
End If
End Function


Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoImporte = "#,###,###,##0.00"


    'Borramos uno de los archivos temporales
    If Dir(App.path & "\ErrActua.txt") <> "" Then Kill App.path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad(True)
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    Conn.Execute "Delete from zBloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    'Buscamos si hay integraciones pendientes. Lo almacenamios en la variable global: TieneIntegracionesPendientes
    TieneIntegracionesPendientes = BuscarIntegraciones(False, "??")
    
End Sub


'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
    'Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
    'Cad = Cad & ";UID=" & vConfig.User
    'Cad = Cad & ";PWD=" & vConfig.password
    'NUEVA FORMA DE ABRIR LA BD
    'Cad = "DESC=;DATA SOURCE= vUsuarios;DATABASE=" & vUsu.CadenaConexion '& ";SERVER=" & vConfig.SERVER ''-- Elimado por rafa V. 4.4.6. (OJO)
    
    If Procesador64bits Then
        Cad = "Provider=MSDASQL.1;Persist Security Info=True;Data Source=vUsuarios;DATABASE=" & vUsu.CadenaConexion
    
    Else
        '
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER
        
        Cad = Cad & ";UID=" & vConfig.User
        Cad = Cad & ";PWD=" & vConfig.password
        '---- Laura: 29/09/2006
        Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
        Cad = Cad & ";Persist Security Info=true"
    
    
    End If
    Conn.ConnectionString = Cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n.", Err.Description
End Function

Public Function AbrirConexionUsuarios() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    AbrirConexionUsuarios = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient
    Conn.CursorLocation = adUseServer
   ' Cad = "DESC=;DATA SOURCE= vUsuarios;DATABASE=USUARIOS;SERVER=" & vConfig.SERVER
    
    Procesador64bits = False
    
    'Esta es la buena
    If Procesador64bits Then   'KIKO valsur
        
        Cad = "Provider=MSDASQL.1;Persist Security Info=True;Data Source=vUsuarios"

    Else
        'Modo NORMAL. Hasta ahora este fncionaba en todos los sitios
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=usuarios;SERVER=" & vConfig.SERVER
        Cad = Cad & ";UID=" & vConfig.User
        Cad = Cad & ";PWD=" & vConfig.password
        Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
    End If
    Conn.ConnectionString = Cad
    Conn.Open
    AbrirConexionUsuarios = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexi�n usuarios.", Err.Description
End Function


'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=1"
End Sub


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "," & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = CADENA
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & ":" & Mid(CADENA, I + 1)
        End If
    Loop Until I = 0
    TransformaPuntosHoras = CADENA
End Function


Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"
                    DBLet = ""
                Case "N"
                    DBLet = 0
                Case "F"
                    DBLet = "0:00:00"
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

Public Function DBMemo(vData As Variant) As String
Dim C As String
    On Error Resume Next
    C = vData
    If Err.Number <> 0 Then
        'Borramos error
        Err.Clear
        C = ""
    End If
    DBMemo = C
End Function


'MODIFICADO. Conta nueva. Ambito fechas
Public Function FechaCorrecta2(vFecha As Date) As Byte
'--------------------------------------------------------
'   Dada una fecha dira si pertenece o no
'   al intervalo de fechas que maneja la apliacion
'   Resultados:
'       0 .- A�o actual
'       1 .- Siguiente
'       2 .- Ambito fecha. Fecha menor a la del ambito !!!!! NUEVO !!!!
'       3 .- Anterior al inicio
'       4 .- Posterior al fin
'--------------------------------------------------------
    
    If vFecha >= vParam.fechaini Then
        'Mayor que fecha inicio
        If vFecha >= vParam.FechaActiva Then
            If vFecha <= vParam.fechafin Then
                FechaCorrecta2 = 0
            Else
                'Compruebo si el a�o siguiente
                If vFecha <= DateAdd("yyyy", 1, vParam.fechafin) Then
                    FechaCorrecta2 = 1
                Else
                    FechaCorrecta2 = 4   'Fuera ejercicios
                End If
            End If
        Else
            FechaCorrecta2 = 2   'Menor que fecha actvia
        End If
    Else            '< fecha ini
        FechaCorrecta2 = 3
    End If
End Function


Public Sub MuestraError(numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        Cad = Cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If Conn.Errors.Count > 0 Then
        ControlamosError Aux
        Conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then Cad = Cad & "N�mero: " & numero & vbCrLf & "Descripci�n: " & Error(numero)
    MsgBox Cad, vbExclamation
End Sub

Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function RellenaCodigoCuenta(vCodigo As String) As String
    Dim I As Integer
    Dim J As Integer
    Dim Cont As Integer
    Dim Cad As String
    
    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    I = 0: Cont = 0
    Do
        I = I + 1
        I = InStr(I, vCodigo, ".")
        If I > 0 Then
            If Cont > 0 Then Cont = 1000
            Cont = Cont + I
        End If
    Loop Until I = 0
    
    'Habia mas de un punto
    If Cont > 1000 Or Cont = 0 Then Exit Function
    
    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    I = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - I
    Cad = ""
    For I = 1 To J
        Cad = Cad & "0"
    Next I
    
    Cad = Mid(vCodigo, 1, Cont - 1) & Cad
    Cad = Cad & Mid(vCodigo, Cont + 1)
    RellenaCodigoCuenta = Cad
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim RS As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        DevuelveDesdeBD = DBLet(RS.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(RS.Fields(1))
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD." & vbCrLf & Cad, Err.Description
End Function

'Obvio
Public Function EsCuentaUltimoNivel(Cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresa.DigitosUltimoNivel)
End Function


Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef Devuelve As String) As Boolean
'Comprueba si es numerica
Dim SQL As String

CuentaCorrectaUltimoNivel = False
If Cuenta = "" Then
    Devuelve = "Cuenta vacia"
    Exit Function
End If
If Not IsNumeric(Cuenta) Then
    Devuelve = "La cuenta debe de ser num�rica: " & Cuenta
    Exit Function
End If

'Rellenamos si procede
Cuenta = RellenaCodigoCuenta(Cuenta)

If Not EsCuentaUltimoNivel(Cuenta) Then
    Devuelve = "No es cuenta de �ltimo nivel: " & Cuenta
    Exit Function
End If

SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
If SQL = "" Then
    Devuelve = "No existe la cuenta : " & Cuenta
    Exit Function
End If

'Llegados aqui, si que existe la cuenta
CuentaCorrectaUltimoNivel = True
Devuelve = SQL
End Function

'-------------------------------------------------------------------------
'
'   Es la misma solo k no si no existe cuenta no da error
Public Function CuentaCorrectaUltimoNivelSIN(ByRef Cuenta As String, ByRef Devuelve As String) As Byte
'Comprueba si es numerica
Dim SQL As String

CuentaCorrectaUltimoNivelSIN = 0
If Cuenta = "" Then
    Devuelve = "Cuenta vacia"
    Exit Function
End If
If Not IsNumeric(Cuenta) Then
    Devuelve = "La cuenta debe de ser num�rica: " & Cuenta
    Exit Function
End If

'Rellenamos si procede
Cuenta = RellenaCodigoCuenta(Cuenta)

CuentaCorrectaUltimoNivelSIN = 1
If Not EsCuentaUltimoNivel(Cuenta) Then
    SQL = "No es cuenta de �ltimo nivel"
Else
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
    If SQL = "" Then
        SQL = "No existe la cuenta  "
    Else
        CuentaCorrectaUltimoNivelSIN = 2
    End If
End If

'Llegados aqui, si que existe la cuenta
Devuelve = SQL
End Function


'Devuelve, para un nivel determinado, cuantos digitos tienen las cuentas
' a ese nivel
Public Function DigitosNivel(numnivel As Integer) As Integer
    Select Case numnivel
    Case 1
        DigitosNivel = vEmpresa.numdigi1

    Case 2
        DigitosNivel = vEmpresa.numdigi2
        
    Case 3
        DigitosNivel = vEmpresa.numdigi3
        
    Case 4
        DigitosNivel = vEmpresa.numdigi4
        
    Case 5
        DigitosNivel = vEmpresa.numdigi5
        
    Case 6
        DigitosNivel = vEmpresa.numdigi6
        
    Case 7
        DigitosNivel = vEmpresa.numdigi7
        
    Case 8
        DigitosNivel = vEmpresa.numdigi8
        
    Case 9
        DigitosNivel = vEmpresa.numdigi9
        
    Case 10
        DigitosNivel = vEmpresa.numdigi10
        
    Case Else
        DigitosNivel = -1
    End Select
End Function

Public Function NivelCuenta(CodigoCuenta As String) As Integer
Dim lon As Integer
Dim niv As Integer
Dim I As Integer
    NivelCuenta = -1
    lon = Len(CodigoCuenta)
    I = 0
    Do
       I = I + 1
       niv = DigitosNivel(I)
       If niv > 0 Then
            If niv = lon Then
                NivelCuenta = I
                I = 11 'para salir del bucle
            End If
        Else
            I = 11 'salimos pq ya no hay nveles para las cuentas de longitud lon
        End If
    Loop Until I > 10
End Function


Public Function ExistenSubcuentas(ByRef Cuenta As String, Nivel As Integer) As Boolean
Dim I As Integer
Dim B As Boolean
Dim Cad As String
    
    I = DigitosNivel(Nivel)
    Cad = Mid(Cuenta, 1, I)
    Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cad, "T")
    If Cad = "" Then
        'NO existe la subcuenta de nivel N
        'salimos
        ExistenSubcuentas = False
        Exit Function
    End If
    If Nivel > 1 Then
        ExistenSubcuentas = ExistenSubcuentas(Cuenta, Nivel - 1)
    Else
        ExistenSubcuentas = True
    End If
End Function


Public Function CreaSubcuentas(ByRef Cuenta, HastaNivel As Integer, TEXTO As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim Cad As String
Dim Cta As String

On Error GoTo ECreaSubcuentas
CreaSubcuentas = False
For I = 1 To HastaNivel
    J = DigitosNivel(I)
    Cta = Mid(Cuenta, 1, J)
    Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T")
    If Cad = "" Then
        'CreaCuenta
        Cad = "INSERT INTO cuentas (codmacta, nommacta, apudirec, model347, razosoci, "
        Cad = Cad & " dirdatos, codposta, despobla, desprovi, nifdatos, maidatos, webdatos,"
        Cad = Cad & " obsdatos) VALUES ("
        Cad = Cad & "'" & Cta
        Cad = Cad & "', '" & TEXTO
        Cad = Cad & "', "
        Cad = Cad & "'N', 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
        Conn.Execute Cad
    End If
Next I
CreaSubcuentas = True
Exit Function
ECreaSubcuentas:
    MuestraError Err.Number, "Creando subcuentas", Err.Description
End Function




Public Function CambiarBarrasPATH(ParaGuardarBD As Boolean, CADENA) As String
Dim I As Integer
Dim Ch As String
Dim Ch2 As String

If ParaGuardarBD Then
    Ch = "\"
    Ch2 = "/"
Else
    Ch = "/"
    Ch2 = "\"
End If
I = 0
Do
    I = I + 1
    I = InStr(1, CADENA, Ch)
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Ch2 & Mid(CADENA, I + 1)
Loop Until I = 0
CambiarBarrasPATH = CADENA
End Function


Public Function ImporteSinFormato(CADENA As String) As String
Dim I As Integer
'Quitamos puntos
Do
    I = InStr(1, CADENA, ".")
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Mid(CADENA, I + 1)
Loop Until I = 0
ImporteSinFormato = TransformaPuntosComas(CADENA)
End Function



'Periodo vendran las fechas Ini y fin con pipe final
Public Sub SaldoHistorico(Cuenta As String, Periodo As String, DescCuenta As String, EsSobreEjerciciosCerrados As Boolean)
Dim RS As Recordset
Dim SQL As String
Dim RC2 As String
Dim vImp As Currency

    Screen.MousePointer = vbHourglass
    SQL = "Select Sum(timporteD),sum(timporteH) from hlinapu"
    If EsSobreEjerciciosCerrados Then SQL = SQL & "1"
    SQL = SQL & " WHERE codmacta='" & Cuenta & "'"
    
    If Not EsSobreEjerciciosCerrados Then _
        SQL = SQL & " AND fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
    SQL = SQL & " AND punteada "
    Set RS = New ADODB.Recordset
    RC2 = Cuenta & "|"
    'PUNTEADO
    RS.Open SQL & "='S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
    Else
        RC2 = RC2 & "||"
    End If
    RS.Close
    'SIN puntear
    RS.Open SQL & "<>'S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
       RC2 = RC2 & Format(RS.Fields(0), FormatoImporte) & "|"
       RC2 = RC2 & Format(RS.Fields(1), FormatoImporte) & "|"
    Else
        RC2 = RC2 & "||"
    End If
    RS.Close
    
    'En el periodo. Para cuando viene de puntear
    If Periodo <> "" Then
        SQL = "Select Sum(timporteD) , sum(timporteH) from hlinapu"
        If EsSobreEjerciciosCerrados Then SQL = SQL & "1"
        SQL = SQL & " WHERE codmacta='" & Cuenta & "' AND "
        SQL = SQL & Periodo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RS.EOF Then
            vImp = DBLet(RS.Fields(0), "N")
            vImp = vImp - DBLet(RS.Fields(1), "N")
            RC2 = RC2 & Format(vImp, FormatoImporte) & "|"
        Else
            RC2 = RC2 & "|"
        End If
    Else
        RC2 = RC2 & "|"
    End If
    RC2 = RC2 & DescCuenta & "|"
    Set RS = Nothing
    'Mostramos la ventanita de mesaje
    frmMensajes.Opcion = 1
    frmMensajes.Parametros = RC2
    frmMensajes.Show vbModal

End Sub

'Lo que hace es comprobar que si la resolucion es mayor
'que 800x600 lo pone en el 400
Public Sub AjustarPantalla(ByRef Formulario As Form)
    If Screen.Width > 13000 Then
        Formulario.Top = 400
        Formulario.Left = 400
    Else
        Formulario.Top = 0
        Formulario.Left = 0
    End If
    Formulario.Width = 12000
    Formulario.Height = 9000
End Sub


'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256.98
'   Tiene que venir num�rico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

If Importe = "" Then
    ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
End If
End Function





Public Function DiasMes(Mes As Byte, Anyo As Integer) As Integer
    Select Case Mes
    Case 2
        If (Anyo Mod 4) = 0 Then
            DiasMes = 29
        Else
            DiasMes = 28
        End If
    Case 1, 3, 5, 7, 8, 10, 12
        DiasMes = 31
    Case Else
        DiasMes = 30
    End Select
End Function





Public Function ComprobarEmpresaBloqueada(codusu As Long, ByRef Empresa As String) As Boolean
Dim Cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
Conn.Execute "Delete from Usuarios.vBloqBD where codusu=" & codusu

'Ahora comprobamos k nadie bloquea la BD
Cad = DevuelveDesdeBD("codusu", "Usuarios.vBloqBD", "conta", Empresa, "T")
If Cad <> "" Then
    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
    
    Set miRsAux = New ADODB.Recordset
    Cad = "show processlist"
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(3) = Empresa Then
            Cad = miRsAux.Fields(2)
            miRsAux.MoveLast
        End If
    
        'Siguiente
        miRsAux.MoveNext
    Wend
    
    If Cad = "" Then
        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
        Conn.Execute "Delete from Usuarios.vBloqBD where conta ='" & Empresa & "'"
        
    Else
        MsgBox "BD bloqueada.", vbCritical
        ComprobarEmpresaBloqueada = True
    End If
End If

Conn.Execute "commit"
End Function


Public Function Bloquear_DesbloquearBD(Bloquear As Boolean) As Boolean

On Error GoTo EBLo
    Bloquear_DesbloquearBD = False
    If Bloquear Then
        CadenaDesdeOtroForm = "INSERT INTO Usuarios.vBloqBD (codusu, conta) VALUES (" & vUsu.Codigo & ",'" & vUsu.CadenaConexion & "')"
    Else
        CadenaDesdeOtroForm = "DELETE FROM  Usuarios.vBloqBD WHERE codusu =" & vUsu.Codigo & " AND conta = '" & vUsu.CadenaConexion & "'"
    End If
    Conn.Execute CadenaDesdeOtroForm
    Bloquear_DesbloquearBD = True
    Exit Function
EBLo:
    'MuestraError Err.Number, "Bloq. BD"
    Err.Clear
End Function




Public Function OtrosPCsContraContabiliad(EsAlIniciar As Boolean) As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String
Dim EquipoConBD As Boolean

    On Error GoTo EOtrosPCsContraContabiliad
    
    Set MiRS = New ADODB.Recordset
    EquipoConBD = (UCase(vUsu.PC) = UCase(vConfig.SERVER)) Or (LCase(vConfig.SERVER) = "localhost")
    Cad = "show processlist"
    MiRS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If UCase(MiRS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    
                    NumRegElim = 0
                    If Equipo <> "LOCALHOST" Then
                        'Si no es localhost
                        NumRegElim = 1
                    Else
                        'HAy un proceso de loclahost. Luego, si el equipo no tiene la BD
                        If Not EquipoConBD Then NumRegElim = 1
                    End If
                    
                    
                    
                    'Si hay que insertar
                    If NumRegElim = 1 Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = Cad
    Exit Function
EOtrosPCsContraContabiliad:
    MuestraError Err.Number, Err.Description, "Leyendo PROCESSLIST"
    Set MiRS = Nothing
    If EsAlIniciar Then
        OtrosPCsContraContabiliad = "LEYENDOPC|"
    Else
        Cad = "�El sistema no puede determinar si hay PCs conectados. �Desea continuar igualmente?"
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
            OtrosPCsContraContabiliad = ""
        Else
            OtrosPCsContraContabiliad = "USUARIO ACTUAL|"
        End If
    End If
    
    
    
End Function


Public Function EsNumerico(TEXTO As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim Cad As String
    
    EsNumerico = False
    Cad = ""
    If Not IsNumeric(TEXTO) Then
        Cad = "El campo debe ser num�rico"
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, TEXTO, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then Cad = "Numero de puntos incorrecto"
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, TEXTO, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then Cad = "Numero incorrecto"
        End If
        
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
    Else
        EsNumerico = True
    End If
End Function



Public Function EsFechaOK(ByRef T As TextBox) As Boolean
Dim Cad As String
    
    Cad = T.Text
    If InStr(1, Cad, "/") = 0 Then
        If Len(T.Text) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T.Text) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/20" & Mid(Cad, 5)
        End If
    End If
    
    If IsDate(Cad) Then
        EsFechaOK = True
        T.Text = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOK = False
    End If
End Function



Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
        If Len(T) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/20" & Mid(Cad, 5)
        End If
    End If
    If IsDate(Cad) Then
        EsFechaOKString = True
        T = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

'Devuelve si hay archivos
'                                                        Llevara la forma: 01, 02  para la empresa 1 o 2..
Public Function BuscarIntegraciones(Errores As Boolean, Empresa As String) As Boolean
Dim Cad As String
On Error GoTo Ebuscarintegraciones
    
    BuscarIntegraciones = False
    If vConfig.Integraciones = "" Then Exit Function
    
    Cad = vConfig.Integraciones
    If Right(Cad, 1) <> "\" Then Cad = Cad & "\"
    If Dir(Cad, vbDirectory) = "" Then
        MsgBox "Carpeta de errores no encontrada: " & vConfig.Integraciones, vbExclamation
        Exit Function
    End If
    
    If Errores Then
        Cad = Cad & "ERRORES"
    Else
        Cad = Cad & "INTEGRA"
    End If
    
    'Facturas clientes
    If Dir(Cad & "\FRACLI\*.?" & Empresa) <> "" Then
        BuscarIntegraciones = True
        Exit Function
    End If
    
    'Facturas Proveedores
    If Dir(Cad & "\FRAPRO\*.?" & Empresa) <> "" Then
        BuscarIntegraciones = True
        Exit Function
    End If
    
    'Asientos al diario
    If Dir(Cad & "\ASIDIA\*.?" & Empresa) <> "" Then
        BuscarIntegraciones = True
        Exit Function
    End If

    'Asientos al historico
    If Dir(Cad & "\ASIHCO\*.?" & Empresa) <> "" Then
        BuscarIntegraciones = True
        Exit Function
    End If

    
    
    If Errores Then
        ComprobarFuncionamientoEspia
    End If


    Exit Function
Ebuscarintegraciones:
    MuestraError Err.Number, Err.Description, "Buscar archivos integraciones" & vbCrLf
End Function


Private Sub ComprobarFuncionamientoEspia()
Dim I As Integer
Dim Cad As String
Dim Ruta As String
Dim F As Date
Dim TieneAntiguos As Boolean

On Error GoTo Ebuscarintegraciones
    
    
    
    TieneAntiguos = False
    For I = 1 To 4
    
        Cad = vConfig.Integraciones
        If Right(Cad, 1) <> "\" Then Cad = Cad & "\"
        Cad = Cad & "AUTOMA"

        If Dir(Cad, vbDirectory) <> "" Then
            Select Case I
            Case 1
                Ruta = Cad & "\FRACLI"
            Case 2
                Ruta = Cad & "\FRAPRO"
            Case 3
                Ruta = Cad & "\ASIHCO"
            Case Else
                Ruta = Cad & "\ASIDIA"
            End Select
    
            If Dir(Ruta, vbDirectory) <> "" Then
                
                Cad = Dir(Ruta & "\*.*", vbArchive)
                Do While Cad <> ""
                    F = FileDateTime(Ruta & "\" & Cad)
                    NumRegElim = Abs(DateDiff("n", Now, F))
                    'MAYOR QUE 10 MINUTOS
                    If NumRegElim > 10 Then
                        'ERROR
                        TieneAntiguos = True
                        Cad = "" 'Para que se salga sin hacer mas
                        Exit For
                    Else
                        Cad = Dir
                    End If
    
                Loop
                
            End If
            
        End If
    
    Next I

    
    If TieneAntiguos Then
        Ruta = ""
        For I = 1 To 50
            Ruta = Ruta & "*"
        Next I
        Cad = Ruta & vbCrLf & vbCrLf & vbCrLf
        Cad = Cad & "        Existen archivos en la carpeta AUTOMA pendientes de " & vbCrLf
        Cad = Cad & "        integrar desde hace m�s de 10 minutos." & vbCrLf & vbCrLf & vbCrLf & Ruta
        MsgBox Cad, vbCritical
        
    End If
    NumRegElim = 0

    Exit Sub
Ebuscarintegraciones:
    MuestraError Err.Number, Err.Description, "Buscar archivos AUTOMA" & vbCrLf


End Sub



'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
End Sub

Public Function DevNombreSQL(CADENA As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = CADENA
End Function



'Para los balnces
Public Function FechaInicioIGUALinicioEjerecicio(FecIni As Date, EjerciciosCerrados1 As Boolean) As Byte
Dim Fecha As Date
Dim Salir As Boolean
Dim I As Integer
On Error GoTo EfechaInicioIGUALinicioEjerecicio

    FechaInicioIGUALinicioEjerecicio = 1
    If EjerciciosCerrados1 Then
        I = -1 'En ejercicios cerrados emp�zamos mirando un a�o por debajo fecini
    Else
        I = 1
    End If
    Fecha = DateAdd("yyyy", I, vParam.fechaini)
    Salir = False
    While Not Salir
        If FecIni = Fecha Then
            'Fecha inicio del listado contiene es fecha incio ejercicio
            FechaInicioIGUALinicioEjerecicio = 0
            'Modificacion del 2 de Septiembre de 2004
            'Si la fehca es incio pero el de el ejercicio siguiente
            'entonces no te lo
            If Not EjerciciosCerrados1 Then
                'Ejerecicio actual / siguiente
                If FecIni > vParam.fechaini Then
                    'Ejercicio siguiente. Con lo cual SI tengo k poner los acumulados
                    FechaInicioIGUALinicioEjerecicio = 1
                End If
            End If
            
            
            
            Salir = True
        Else
            If FecIni < Fecha Then
                Fecha = DateAdd("yyyy", -1, Fecha)
            Else
                Salir = True
            End If
        End If
    Wend
    
    Exit Function
EfechaInicioIGUALinicioEjerecicio:
    Err.Clear  'No tiene importancia
End Function





Public Function DevuelveDigitosNivelAnterior() As Integer
Dim J As Integer
    DevuelveDigitosNivelAnterior = 3
    If vEmpresa Is Nothing Then Exit Function
    If vEmpresa.numnivel < 2 Then Exit Function
    J = vEmpresa.numnivel - 1
    J = DigitosNivel(J)
    If J < 3 Then J = 3
    DevuelveDigitosNivelAnterior = J
End Function



'--------------------------------------------------------------------------
' Los numeros vendran formateados o sin formatear, pero siempre viene texto
'
Public Function CadenaCurrency(TEXTO As String, ByRef Importe As Currency) As Boolean
Dim I As Integer

    On Error GoTo ECadenaCurrency
    Importe = 0
    CadenaCurrency = False
    If Not IsNumeric(TEXTO) Then Exit Function
    I = InStr(1, TEXTO, ",")
    If I = 0 Then
        'Significa k el numero no esta  formateado y como mucho tiene punto
        Importe = CCur(TransformaPuntosComas(TEXTO))
    Else
        Importe = ImporteFormateado(TEXTO)
    End If
    
    CadenaCurrency = True
    
    Exit Function
ECadenaCurrency:
    Err.Clear
End Function


Public Function UsuariosConectados(vMens As String, Optional DejarContinuar As Boolean) As Boolean
Dim I As Integer
Dim Cad As String
Dim metag As String
Dim SQL As String
Cad = OtrosPCsContraContabiliad(False)
UsuariosConectados = False
If Cad <> "" Then
    UsuariosConectados = True
    I = 1
    metag = vMens
    If vMens <> "" Then metag = metag & vbCrLf
    metag = metag & vbCrLf & "Los siguientes PC's est�n conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
    
    Do
        SQL = RecuperaValor(Cad, I)
        If SQL <> "" Then
            metag = metag & "    - " & SQL & vbCrLf
            I = I + 1
        End If
    Loop Until SQL = ""
    If DejarContinuar Then
        'Hare la pregunta
        metag = metag & vbCrLf & "�Continuar?"
        If MsgBox(metag, vbQuestion + vbYesNoCancel) = vbYes Then UsuariosConectados = False
    Else
        'Informa UNICAMENTE
        MsgBox metag, vbExclamation
    End If
End If
End Function




Public Function HayKHabilitarCentroCoste(ByRef Cuenta As String) As Boolean
Dim Ch As String

    HayKHabilitarCentroCoste = False
    If Cuenta <> "" Then
        'Hay cuenta
         Ch = Mid(Cuenta, 1, 1)
         If Ch = vParam.grupogto Or Ch = vParam.grupovta Or Ch = vParam.grupoord Then
            HayKHabilitarCentroCoste = True
         Else
            Ch = Mid(Cuenta, 1, 3)
            If Ch = vParam.Subgrupo1 Or Ch = vParam.Subgrupo2 Then
                HayKHabilitarCentroCoste = True
            End If
        End If
    End If
End Function


Public Function EjecutaSQL(ByRef SQL As String) As Boolean
    EjecutaSQL = False
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        Err.Clear
    Else
        EjecutaSQL = True
    End If
End Function


Public Function FacturaContabilizada(Serie As String, numero As String, Anofac As String) As Boolean
Dim RT As ADODB.Recordset
Dim C As String

    Set RT = New ADODB.Recordset
    FacturaContabilizada = True
    C = "Select numasien from cabfact"
    If Serie = "" Then C = C & "prov"
    C = C & " WHERE anofac"
    If Serie = "" Then
        C = C & "pr"
    Else
        C = C & "cl"
    End If
    C = C & " = " & Anofac & " AND "
    If Serie = "" Then
        C = C & " numregis = "
    Else
        C = C & " numserie = '" & Serie & "' AND codfaccl ="
    End If
    C = C & numero
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If IsNull(RT!Numasien) Then FacturaContabilizada = False
    End If
    RT.Close
    Set RT = Nothing
End Function

Public Function DirectorioEAT() As Boolean
    On Error GoTo EDirecEAT
    DirectorioEAT = False
    If Dir("C:\AEAT", vbDirectory) = "" Then
        MsgBox "No se encuentra la carpeta de la agencia tributaria.  ( C:\AEAT )", vbExclamation
    Else
        DirectorioEAT = True
    End If
    Exit Function
EDirecEAT:
    Err.Clear
End Function





Public Function EstaLaCuentaBloqueada(ByRef codmacta As String, Fecha As Date) As Boolean
Dim I As Integer

        EstaLaCuentaBloqueada = False
        If vParam.CuentasBloqueadas <> "" Then
            I = InStr(1, vParam.CuentasBloqueadas, codmacta & ":")
            If I > 0 Then
                'La cuenta esta con fecha de bloqueo
                If Fecha >= CDate(Mid(vParam.CuentasBloqueadas, I + Len(codmacta) + 1, 10)) Then EstaLaCuentaBloqueada = True
            End If
        End If
End Function


Public Sub CerrarRs(ByRef Rsss As ADODB.Recordset)
    On Error Resume Next
    Rsss.Close
    If Err.Number <> 0 Then Err.Clear
End Sub


'*******************************************************************
'*******************************************************************
'*******************************************************************
'   Septiembre 2011
'
'  Letra serie 3 Digitos
'  Con lo cual para algunas campos (numdocum de linapu) son un maximo de
'   10 posiciones. Como antes era un digito letra ser, formateabamos con 9
'       numerofactura debe ser NUMERICO
Public Function SerieNumeroFactura(Posiciones As Integer, Serie As String, Numerofactura As String)
Dim I As Integer
Dim Cad As String
    
    I = Posiciones - Len(Numerofactura) - Len(Serie)
    If I <= 0 Then
        'Hay menos posiciones de las que podemos meter
        Cad = Right(Numerofactura, Posiciones - Len(Numerofactura))
    Else
        Cad = String(I, "0") & Numerofactura
    End If
    SerieNumeroFactura = Serie & Cad
    
    
End Function






Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim ent As Integer
Dim Cad As String
  
  ' Comprobaciones
  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un n�mero."
    Exit Function
  End If
  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If
  
  ' Redondeo.
  Cad = "0"
  If NumDigitsAfterDecimals <> 0 Then Cad = Cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Format(Number, Cad)
  
End Function

