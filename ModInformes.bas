Attribute VB_Name = "ModInformes"
Option Explicit


Dim RS As Recordset
Dim Cad As String
Dim SQL As String
Dim I As Integer


'Esto sera para el pb general
Dim TotalReg As Long
Dim Actual As Long


'Esta funcion lo que hace es genera el valor del campo
'El campo lo coge del recordset, luego sera field(i), y el tipo es para añadirle
'las coimllas, o quitarlas comas
'  Si es numero viene un 1 si no nada
Private Function ParaBD(ByRef Campo As ADODB.Field, Optional EsNumerico As Byte) As String
    
    If IsNull(Campo) Then
        ParaBD = "NULL"
    Else
        Select Case EsNumerico
        Case 1
            ParaBD = TransformaComasPuntos(CStr(Campo))
        Case 2
            'Fechas
            ParaBD = "'" & Format(CStr(Campo), "dd/mm/yyyy") & "'"
        Case Else
            ParaBD = "'" & Campo & "'"

            
        End Select
    End If
    ParaBD = "," & ParaBD
End Function


'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------

'   En este modulo se crearan los datos para los informes
'   Con lo cual cada Function Generara los datos en la tabla



'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------











Public Function InformeConceptos(ByRef vSQL As String) As Boolean

On Error GoTo EGI_Conceptos
    InformeConceptos = False
    'Borramos los anteriores
    Conn.Execute "Delete from Usuarios.zconceptos where codusu = " & vUsu.Codigo
    Cad = "INSERT INTO Usuarios.zconceptos (codusu, codconce, nomconce,tipoconce) VALUES ("
    Cad = Cad & vUsu.Codigo & ",'"
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = Cad & Format(RS.Fields(0), "000")
        SQL = SQL & "','" & RS.Fields(1) & "','" & RS.Fields(3) & "')"
        Conn.Execute SQL
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    InformeConceptos = True
Exit Function
EGI_Conceptos:
    MuestraError Err.Number
    Set RS = Nothing
End Function





Public Function ListadoEstadisticas(ByRef vSQL As String) As Boolean
On Error GoTo EListadoEstadisticas
    ListadoEstadisticas = False
    Conn.Execute "Delete from Usuarios.zestadinmo1 where codusu = " & vUsu.Codigo
    'Sentencia insert
    Cad = "INSERT INTO Usuarios.zestadinmo1 (codusu, codigo, codconam, nomconam, codinmov, nominmov,"
    Cad = Cad & "tipoamor, porcenta, codprove, fechaadq, valoradq, amortacu, fecventa, impventa) VALUES ("
    Cad = Cad & vUsu.Codigo & ","
    
    'Empezamos
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 1
    While Not RS.EOF
        
        SQL = I & ParaBD(RS!codconam, 1) & ParaBD(RS!nomconam)
        SQL = SQL & ParaBD(RS!Codinmov) & ",'" & DevNombreSQL(RS!nominmov) & "'"
        
        SQL = SQL & ParaBD(RS!tipoamor) & ParaBD(RS!coeficie) & ParaBD(RS!codprove)
        SQL = SQL & ParaBD(RS!fechaadq, 2) & ParaBD(RS!valoradq, 1) & ParaBD(RS!amortacu, 1)
        SQL = SQL & ParaBD(RS!fecventa, 2) 'FECHA
        SQL = SQL & ParaBD(RS!impventa, 1) & ")"
        Conn.Execute Cad & SQL
        
        'Sig
        RS.MoveNext
        I = I + 1
    Wend
    ListadoEstadisticas = True
EListadoEstadisticas:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Function





Public Function ListadoFichaInmo(ByRef vSQL As String) As Boolean
On Error GoTo Err1
    ListadoFichaInmo = False
    Conn.Execute "Delete from Usuarios.zfichainmo where codusu = " & vUsu.Codigo
    'Sentencia insert
    Cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, Fechaamor,Importe, porcenta) VALUES ("
    Cad = Cad & vUsu.Codigo & ","
    
    'Empezamos
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 1
    While Not RS.EOF
        SQL = RS!nominmov
        NombreSQL SQL
        SQL = I & ParaBD(RS!Codinmov) & ",'" & SQL & "'"
        SQL = SQL & ParaBD(RS!fechaadq, 2) & ParaBD(RS!valoradq, 1) & ParaBD(RS!fechainm, 2)
        SQL = SQL & ParaBD(RS!imporinm, 1) & ParaBD(RS!porcinm, 1)
        SQL = SQL & ")"
        Conn.Execute Cad & SQL
        
        'Sig
        RS.MoveNext
        I = I + 1
    Wend
    RS.Close
    If I > 1 Then
        ListadoFichaInmo = True
    Else
        MsgBox "Ningún registro con esos valores", vbExclamation
    End If
Err1:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Function




Public Function GenerarDatosCuentas(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    GenerarDatosCuentas = False
    Cad = "Delete FROM Usuarios.zCuentas where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    Cad = "INSERT INTO Usuarios.zcuentas (codusu, codmacta, nommacta, razosoci,nifdatos, dirdatos, codposta, despobla, apudirec,model347) "
    Cad = Cad & " SELECT " & vUsu.Codigo & ",ctas.codmacta, ctas.nommacta, ctas.razosoci, ctas.nifdatos, ctas.dirdatos, ctas.codposta, ctas.despobla,ctas.apudirec,ctas.model347"
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".cuentas as ctas "
    If vSQL <> "" Then Cad = Cad & " WHERE " & vSQL
    Conn.Execute Cad
    GenerarDatosCuentas = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
 
End Function



Public Function GenerarDiarios() As Boolean
On Error GoTo EGen
    GenerarDiarios = False
    Cad = "Delete FROM Usuarios.ztiposdiario where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    Cad = "INSERT INTO Usuarios.ztiposdiario (codusu, numdiari, desdiari)"
    Cad = Cad & " SELECT " & vUsu.Codigo & ",d.numdiari,d.desdiari"
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".tiposdiario as d;"
    Conn.Execute Cad
    GenerarDiarios = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function


Public Function GeneraraExtractos() As Boolean
On Error GoTo EGen
    GeneraraExtractos = False
    Cad = "Delete FROM Usuarios.ztmpconextcab where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    Cad = "Delete FROM Usuarios.ztmpconext where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    Cad = "INSERT INTO Usuarios.ztmpconextcab "
    Cad = Cad & "(codusu, cuenta, fechini, fechfin, acumantD, acumantH, acumantT, acumperD, acumperH, acumperT, acumtotD, acumtotH, acumtotT, cta)"
    Cad = Cad & " SELECT " & vUsu.Codigo & ",t.cuenta, t.fechini, t.fechfin, t.acumantD, t.acumantH, t.acumantT, t.acumperD, t.acumperH, t.acumperT, t.acumtotD, t.acumtotH, t.acumtotT, t.cta"
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".tmpconextcab as t where t.codusu =" & vUsu.Codigo & ";"
    Conn.Execute Cad
    
    
    'Las lineas
    Cad = "INSERT INTO Usuarios.ztmpconext (codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, ampconce, timporteD, timporteH, saldo, Punteada, contra, ccost)"
    Cad = Cad & " SELECT " & vUsu.Codigo & ",t.cta, t.numdiari, t.Pos, t.fechaent, t.numasien, t.linliapu, t.nomdocum, t.ampconce, t.timporteD, t.timporteH, t.saldo, t.Punteada, t.contra, t.ccost"
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".tmpconext as t where t.codusu =" & vUsu.Codigo & ";"
    Conn.Execute Cad
    GeneraraExtractos = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function


'Para la impresion de extractos y demas, MAYOR, etc
Public Function GeneraraExtractosListado(Cuenta As String) As Boolean
On Error GoTo EGen
    GeneraraExtractosListado = False
    Cad = "INSERT INTO Usuarios.ztmpconextcab "
    Cad = Cad & "(codusu, cuenta, fechini, fechfin, acumantD, acumantH, acumantT, acumperD, acumperH, acumperT, acumtotD, acumtotH, acumtotT, cta)"
    Cad = Cad & " SELECT " & vUsu.Codigo & ",t.cuenta, t.fechini, t.fechfin, t.acumantD, t.acumantH, t.acumantT, t.acumperD, t.acumperH, t.acumperT, t.acumtotD, t.acumtotH, t.acumtotT, t.cta"
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".tmpconextcab as t where t.codusu =" & vUsu.Codigo & " AND cta ='" & Cuenta & "';"
    Conn.Execute Cad
    
    'Las lineas
    Cad = "INSERT INTO Usuarios.ztmpconext (codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, ampconce, timporteD, timporteH, saldo, Punteada, contra, ccost)"
    Cad = Cad & " SELECT " & vUsu.Codigo & ",t.cta, t.numdiari, t.Pos, t.fechaent, t.numasien, t.linliapu, t.nomdocum, t.ampconce, t.timporteD, t.timporteH, t.saldo, t.Punteada, t.contra, t.ccost"
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".tmpconext as t where t.codusu =" & vUsu.Codigo & " AND cta ='" & Cuenta & "';"
    Conn.Execute Cad
    GeneraraExtractosListado = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function


Public Function IAsientosErrores(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    IAsientosErrores = False
    Cad = "Delete FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    'Las lineas
    Cad = "INSERT INTO Usuarios.zdiapendact (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum,"
    Cad = Cad & " ampconce, timporteD, timporteH, codccost)"
    Cad = Cad & " SELECT " & vUsu.Codigo
    Cad = Cad & ",linapue.numdiari, tiposdiario.desdiari, linapue.fechaent, linapue.numasien, linapue.linliapu, linapue.codmacta, cuentas.nommacta, linapue.numdocum, linapue.ampconce, linapue.timporteD, linapue.timporteH, linapue.codccost"
    Cad = Cad & " FROM (linapue LEFT JOIN tiposdiario ON linapue.numdiari = tiposdiario.numdiari) LEFT JOIN cuentas ON linapue.codmacta = cuentas.codmacta"
    If vSQL <> "" Then Cad = Cad & " WHERE " & vSQL
    
    
    
    
    Conn.Execute Cad
    
    Set RS = New ADODB.Recordset
    Cad = "select count(*) FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then Cad = ""
    End If
    RS.Close
    Set RS = Nothing
    
    If Cad <> "" Then
        MsgBox "Ningun registro por mostrar.", vbExclamation
        Exit Function
    End If
    IAsientosErrores = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Public Function IDiariosPendientes(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    IDiariosPendientes = False
    Cad = "Delete FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    'Las lineas
    Cad = "INSERT INTO Usuarios.zdiapendact (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum,"
    Cad = Cad & " ampconce, timporteD, timporteH, codccost)"
    Cad = Cad & " SELECT " & vUsu.Codigo
    Cad = Cad & ",cabapu_0.numdiari, tiposdiario_0.desdiari, cabapu_0.fechaent, cabapu_0.numasien, linapu_0.linliapu,"
    Cad = Cad & " linapu_0.codmacta, cuentas_0.nommacta, linapu_0.numdocum, linapu_0.ampconce, linapu_0.timporteD, "
    Cad = Cad & " linapu_0.timporteH, linapu_0.codccost  FROM cabapu cabapu_0, cuentas cuentas_0, linapu linapu_0, tiposdiario "
    Cad = Cad & " tiposdiario_0  WHERE linapu_0.fechaent = cabapu_0.fechaent AND linapu_0.numasien = cabapu_0.numasien AND "
    Cad = Cad & " linapu_0.numdiari = cabapu_0.numdiari AND tiposdiario_0.numdiari = cabapu_0.numdiari AND"
    Cad = Cad & " tiposdiario_0.numdiari = linapu_0.numdiari AND cuentas_0.codmacta = linapu_0.codmacta"
    If vSQL <> "" Then Cad = Cad & " AND " & vSQL
    
    
    
    
    Conn.Execute Cad
    
    Set RS = New ADODB.Recordset
    Cad = "select count(*) FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then Cad = ""
    End If
    RS.Close
    Set RS = Nothing
    
    If Cad <> "" Then
        MsgBox "Ningun registro por mostrar.", vbExclamation
        Exit Function
    End If
    IDiariosPendientes = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function




Public Function ITotalesCtaConcepto(ByRef vSQL As String, Tabla As String) As Boolean
On Error GoTo EGen
    ITotalesCtaConcepto = False
    Cad = "Delete FROM Usuarios.ztotalctaconce  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    'Las lineas
    Cad = "INSERT INTO Usuarios.ztotalctaconce (codusu, codmacta, nommacta, nifdatos, fechaent, timporteD, timporteH, codconce)"
    Cad = Cad & " SELECT " & vUsu.Codigo
    Cad = Cad & " ," & Tabla & ".codmacta, nommacta, nifdatos, fechaent,"
    Cad = Cad & " timporteD,timporteH, codconce"
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".cuentas ,"
    Cad = Cad & vUsu.CadenaConexion & "." & Tabla & " WHERE cuentas.codmacta = " & Tabla & ".codmacta"
    If vSQL <> "" Then Cad = Cad & " AND " & vSQL
    
    
    
    
    Conn.Execute Cad
    
    
    'Inserto en ztmpdiarios para los que sean TODOS los conceptos
    Cad = "Delete FROM Usuarios.ztiposdiario  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    Cad = "INSERT INTO Usuarios.ztiposdiario SELECT " & vUsu.Codigo & ",codconce,nomconce FROM conceptos"
    Conn.Execute Cad
    
    
    'Contamos para ver cuantos hay
    Cad = "Select count(*) from Usuarios.ztotalctaconce WHERE codusu =" & vUsu.Codigo
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then I = 1
        End If
    End If
    RS.Close
    Set RS = Nothing
    If I > 0 Then
        ITotalesCtaConcepto = True
    Else
        MsgBox "Ningún registro con esos valores.", vbExclamation
    End If
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Public Function IAsientosPre(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    IAsientosPre = False
    Cad = "Delete FROM Usuarios.zasipre  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "INSERT INTO Usuarios.zasipre (codusu, numaspre, nomaspre, linlapre, codmacta, nommacta"
    Cad = Cad & ", ampconce, timporteD, timporteH, codccost)"

    Cad = Cad & " SELECT " & vUsu.Codigo
    Cad = Cad & ", t1.numaspre, t1.nomaspre, t2.linlapre,t2.codmacta, t3.nommacta,t2.ampconce,"
    Cad = Cad & "t2.timported,t2.timporteh,t2.codccost FROM "
    Cad = Cad & vUsu.CadenaConexion & ".cabasipre as t1,"
    Cad = Cad & vUsu.CadenaConexion & ".linasipre as t2,"
    Cad = Cad & vUsu.CadenaConexion & ".cuentas as t3 WHERE "
    Cad = Cad & " t1.numaspre=t2.numaspre AND t2.codmacta=t3.codmacta"
    If vSQL <> "" Then Cad = Cad & " AND " & vSQL
    
    
    Conn.Execute Cad
    IAsientosPre = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function





Public Function IHcoApuntes(ByRef vSQL As String, NumeroTabla As String) As Boolean
On Error GoTo EGen
    IHcoApuntes = False
    Cad = "Delete FROM Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    
    Cad = " INSERT INTO Usuarios.zhistoapu (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum, ampconce,"
    Cad = Cad & " timporteD, timporteH, codccost) "
    Cad = Cad & "SELECT " & vUsu.Codigo & ",hcabapu" & NumeroTabla & ".numdiari, tiposdiario.desdiari, hcabapu" & NumeroTabla & ".fechaent, hcabapu" & NumeroTabla & ".numasien, hlinapu" & NumeroTabla & ".linliapu,"
    Cad = Cad & " hlinapu" & NumeroTabla & ".codmacta, cuentas.nommacta, hlinapu" & NumeroTabla & ".numdocum, hlinapu" & NumeroTabla & ".ampconce, hlinapu" & NumeroTabla & ".timporteD,"
    Cad = Cad & " hlinapu" & NumeroTabla & ".timporteH, hlinapu" & NumeroTabla & ".codccost  "
    Cad = Cad & " FROM " & vUsu.CadenaConexion & ".cuentas , " & vUsu.CadenaConexion & ".hcabapu" & NumeroTabla & " , " & vUsu.CadenaConexion & ".hlinapu" & NumeroTabla & ", " & vUsu.CadenaConexion & ".tiposdiario"
    Cad = Cad & " WHERE hlinapu" & NumeroTabla & ".fechaent = hcabapu" & NumeroTabla & ".fechaent AND hlinapu" & NumeroTabla & ".numasien = hcabapu" & NumeroTabla & ".numasien AND"
    Cad = Cad & " hlinapu" & NumeroTabla & ".numdiari = hcabapu" & NumeroTabla & ".numdiari AND cuentas.codmacta = hlinapu" & NumeroTabla & ".codmacta AND tiposdiario.numdiari ="
    Cad = Cad & " hcabapu" & NumeroTabla & ".numdiari AND tiposdiario.numdiari = hlinapu" & NumeroTabla & ".numdiari"
    If vSQL <> "" Then Cad = Cad & " AND " & vSQL
    
    
    
    
    Conn.Execute Cad
    
    Cad = DevuelveDesdeBD("count(*)", "Usuarios.zhistoapu", "codusu", vUsu.Codigo, "N")
    If Val(Cad) = 0 Then
        MsgBox "Ningun registro seleccionado", vbExclamation
    Else
        IHcoApuntes = True
    End If
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function

'                           formateada
'Vienen empipados numasien|  fechaent    |numdiari|
Public Function IHcoApuntesAlActualizarModificar(Cadena1 As String) As Boolean
On Error GoTo EGen
    IHcoApuntesAlActualizarModificar = False
    
    'No borramos, borraremos antes de llamar a esta funcion
    'Cad = "Delete FROM Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
    'Conn.Execute Cad
    
    
    Cad = " INSERT INTO Usuarios.zhistoapu (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum, ampconce,"
    Cad = Cad & " timporteD, timporteH, codccost) "
    Cad = Cad & "SELECT " & vUsu.Codigo & ",hcabapu.numdiari, tiposdiario.desdiari, hcabapu.fechaent, hcabapu.numasien, hlinapu.linliapu,"
    Cad = Cad & " hlinapu.codmacta, cuentas.nommacta, hlinapu.numdocum, hlinapu.ampconce, hlinapu.timporteD,"
    Cad = Cad & " hlinapu.timporteH, hlinapu.codccost  "
    Cad = Cad & " FROM cuentas , hcabapu,hlinapu,tiposdiario"
    Cad = Cad & " WHERE hlinapu.fechaent = hcabapu.fechaent AND hlinapu.numasien = hcabapu.numasien AND"
    Cad = Cad & " hlinapu.numdiari = hcabapu.numdiari AND cuentas.codmacta = hlinapu.codmacta AND tiposdiario.numdiari ="
    Cad = Cad & " hcabapu.numdiari AND tiposdiario.numdiari = hlinapu.numdiari"
    Cad = Cad & " AND hcabapu.numasien  =" & RecuperaValor(Cadena1, 1)
    Cad = Cad & " AND hcabapu.fechaent  ='" & RecuperaValor(Cadena1, 2)
    Cad = Cad & "' AND hcabapu.numdiari =" & RecuperaValor(Cadena1, 3)
    
    Conn.Execute Cad
    IHcoApuntesAlActualizarModificar = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function




Public Function GeneraDatosHcoInmov(ByRef vSQL As String) As Boolean

On Error GoTo EGeneraDatosHcoInmov
    GeneraDatosHcoInmov = False
        
    'Borramos tmp
    Cad = "Delete from Usuarios.zfichainmo where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    'Abrimos datos
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ningún dato a mostrar", vbExclamation
    Else
        Cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, fechaamor, Importe, porcenta) VALUES (" & vUsu.Codigo & ","
        TotalReg = 0
        While Not RS.EOF
           
            TotalReg = TotalReg + 1
            'Metemos los nuevos datos
            SQL = TotalReg & ParaBD(RS!Codinmov, 1) & ",'" & DevNombreSQL(CStr(RS!nominmov)) & "'" & ParaBD(RS!fechainm, 2)
            SQL = SQL & ",NULL,NULL" & ParaBD(RS!imporinm, 1) & ParaBD(RS!porcinm, 1) & ")"
            SQL = Cad & SQL
            Conn.Execute SQL
            RS.MoveNext
        Wend
        GeneraDatosHcoInmov = True
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EGeneraDatosHcoInmov:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function





Public Function GeneraDatosConceptosInmov() As Boolean

On Error GoTo EGeneraDatosConceptosInmov
    GeneraDatosConceptosInmov = False
        
    'Borramos tmp
    Cad = "Delete from Usuarios.ztmppresu1 where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    'Abrimos datos
    SQL = "Select * from sconam"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ningún dato a mostrar", vbExclamation
    Else
        Cad = "INSERT INTO Usuarios.ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (" & vUsu.Codigo
        While Not RS.EOF
            'Metemos los nuevos datos
            SQL = ParaBD(RS!codconam, 1) & ",'" & Format(RS!codconam, "0000") & "'" & ParaBD(RS!nomconam)
            SQL = SQL & ",0" & ParaBD(RS!perimaxi, 1) & ParaBD(RS!coefimaxi, 1) & ")"
            SQL = Cad & SQL
            Conn.Execute SQL
            RS.MoveNext
        Wend
        GeneraDatosConceptosInmov = True
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EGeneraDatosConceptosInmov:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function
