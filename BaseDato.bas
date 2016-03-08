Attribute VB_Name = "BaseDato"
Option Explicit

Private SQL As String

Dim ImpD As Currency
Dim ImpH As Currency
Dim RT As ADODB.Recordset


Dim d As String
Dim H As String
'Para los balances
Dim M1 As Integer   ' años y kmeses para el balance
Dim M2 As Integer
Dim M3 As Integer
Dim A1 As Integer
Dim A2 As Integer
Dim A3 As Integer
Dim vCta As String
Dim ImAcD As Currency  'importes
Dim ImAcH As Currency
Dim ImPerD As Currency  'importes
Dim ImPerH As Currency
Dim ImCierrD As Currency  'importes
Dim ImCierrH As Currency
Dim Contabilidad As Integer
Dim Aux As String
Dim vFecha1 As Date
Dim vFecha2 As Date
Dim VFecha3 As Date
Dim Codigo As String
Dim EjerciciosCerrados As Boolean
Dim NumAsiento As Integer
Dim Nulo1 As Boolean
Dim Nulo2 As Boolean

Dim VarConsolidado(2) As String

Dim EsBalancePerdidas_y_ganancias As Boolean

'Para la precarga de datos del balance de sumas y saldos
Dim RsBalPerGan As ADODB.Recordset


'--------------------------------------------------------------------
'--------------------------------------------------------------------
Private Function ImporteASQL(ByRef Importe As Currency) As String
ImporteASQL = ","
If Importe = 0 Then
    ImporteASQL = ImporteASQL & "NULL"
Else
    ImporteASQL = ImporteASQL & TransformaComasPuntos(CStr(Importe))
End If
End Function



'--------------------------------------------------------------------
'--------------------------------------------------------------------
' El dos sera para k pinte el 0. Ya en el informe lo trataremos.
' Con esta opcion se simplifica bastante la opcion de totales
Private Function ImporteASQL2(ByRef Importe As Currency) As String
    ImporteASQL2 = "," & TransformaComasPuntos(CStr(Importe))
End Function



'--------------------------------------------------------------------
'--------------------------------------------------------------------



Public Sub CommitConexion()
    On Error Resume Next
    Conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function FacturaCorrecta(NumF As Long, AnoF As Integer, ByRef Serie As String) As String


On Error GoTo EFacturaCorrecta
FacturaCorrecta = ""

Set RT = New ADODB.Recordset

'Calculamos el total de los importes
SQL = "Select * from cabfact where numserie = '" & Serie & "' AND codfaccl = " & NumF
SQL = SQL & " AND anofaccl = " & AnoF
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RT.EOF Then
    FacturaCorrecta = "No existe la factura(Serire/Numero/año): " & Serie & " / " & NumF & " / " & AnoF
Else
    'Si que existe la factura
    'Sumamos bases imponibles
    ImpD = RT!ba1faccl
    If Not IsNull(RT!ba2faccl) Then ImpD = ImpD + RT!ba2faccl
    If Not IsNull(RT!ba3faccl) Then ImpD = ImpD + RT!ba3faccl
    
    'IVAS
    ImpD = ImpD + RT!ti1faccl
    If Not IsNull(RT!ti2faccl) Then ImpD = ImpD + RT!ti2faccl
    If Not IsNull(RT!ti3faccl) Then ImpD = ImpD + RT!ti3faccl
    
    'Retenciones
    If Not IsNull(RT!tr1faccl) Then ImpD = ImpD + RT!tr1faccl
    If Not IsNull(RT!tr2faccl) Then ImpD = ImpD + RT!tr2faccl
    If Not IsNull(RT!tr3faccl) Then ImpD = ImpD + RT!tr3faccl
    
    'Importe retencion  (SE LER RESTA)
    If Not IsNull(RT!trefaccl) Then ImpD = ImpD - RT!trefaccl
    
    'Comprobamos que el importe que pone es el que corresponde
    If ImpD <> RT!totfaccl Then
        FacturaCorrecta = "La suma de bases, ivas, retenciones no coincide con el total factura: " & ImpD & " /  " & RT!totfaccl
    Else
        'Si coincide la suma de las facturas
        FacturaCorrecta = ""
End If
RT.Close
If FacturaCorrecta = "" Then
    'Ahora comprobamos que la suma de lineas coincide con el totalfac->impd
    SQL = "SELECT sum(impbascl) FROM linfact  WHERE linfact.numserie= '" & Serie & "'"
    SQL = SQL & " AND linfact.codfaccl= " & NumF
    SQL = SQL & " AND linfact.anofaccl=" & AnoF & ";"
    RT.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    ImpH = 0
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
    End If
    RT.Close
    If ImpD <> ImpH Then _
        FacturaCorrecta = "El importe indicado en la cabecera no coincide con el de la suma de lineas: " & ImpD & " / " & ImpH
        
End If
     
EFacturaCorrecta:
    If Err.Number <> 0 Then _
        FacturaCorrecta = Err.Number & " - " & Err.Description
        Err.Clear
    End If
    Set RT = Nothing
End Function





Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
Dim Cad As String
Dim Aux As String
Dim Ch As String
Dim Fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
Cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    I = CararacteresCorrectos(CADENA, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = Campo & " >= " & Cad & " AND " & Campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    I = 1
                    Cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        Ch = Mid(CADENA, I, 1)
                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                            Cad = Cad & Ch
                            Else
                                Aux = Mid(CADENA, I)
                                Fin = True
                        End If
                        I = I + 1
                        If I > Len(CADENA) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Cad = "" Then Cad = " = "
                    DevSQL = Campo & " " & Cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not EsFechaOKString(Cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        Cad = Format(Cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & Cad & "' AND " & Campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    Ch = Mid(CADENA, I, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        Cad = Cad & Ch
                        Else
                            Aux = Mid(CADENA, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If Cad = "" Then Cad = " = "
                DevSQL = Campo & " " & Cad & " " & Aux
            End If
        End If
    
    
    
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(CADENA, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico

        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        
        'Intervalo correcto
        'Construimos la cadena
        Cad = DevNombreSQL(Cad)
        Aux = DevNombreSQL(Aux)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & Cad & "' AND " & Campo & " <= '" & Aux & "'"
    
    
    Else
    
        'Cambiamos el * por % puesto que en ADO es el caraacter para like
        I = 1
        Aux = CADENA
        While I <> 0
            I = InStr(1, Aux, "*")
            If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
        Wend
        'Cambiamos el ? por la _ pue es su omonimo
        I = 1
        While I <> 0
            I = InStr(1, Aux, "?")
            If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
        Wend
        Cad = Mid(CADENA, 1, 2)
        If Cad = "<>" Then
            Aux = Mid(CADENA, 3)
            DevSQL = Campo & " LIKE '!" & Aux & "'"
            Else
            DevSQL = Campo & " LIKE '" & Aux & "'"
        End If
    End If


    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, CADENA, "<>")
    If I = 0 Then
        'IGUAL A valor
        Cad = " = "
        Else
            'Distinto a valor
        Cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, CADENA, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = Campo & " " & Cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
Dim I As Integer
Dim Ch As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "a" To "z"
            Case "A" To "Z"
            Case "0" To "9"
            'QUITAR#### o no.
            'Modificacion hecha 26-OCT-2006.  Es para que meta la coma como caracter en la busqueda
            Case "*", "%", "?", "_", "\", "/", ":", ".", " ", "-", "," ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "#", "@", "$"
            Case "<", ">"
            Case "Ñ", "ñ"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "F"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function






'Este modulo estaba antes del ADOBUS
Public Function BloquearAsiento(NA As String, ND As String, NF As String, ByRef MostrarMensajeError As String) As Boolean
Dim RB As Recordset

    'Pensar en la coicidencia en el tiempo de dos transacciones es improbable.Teimpo de acceso en milisegudos
    BloquearAsiento = False
    SQL = "SELECT * from cabapu "
    SQL = SQL & " WHERE numdiari =" & ND
    SQL = SQL & " AND fechaent='" & NF
    SQL = SQL & "' AND numasien=" & NA
    Set RB = New Recordset
    RB.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If RB.EOF Then
        RB.Close
        Set RB = Nothing
        If MostrarMensajeError <> "" Then
            MsgBox "Asiento inexistente o ha sido borrado", vbExclamation
        Else
            MostrarMensajeError = "Asiento inexistente o ha sido borrado"
        End If
        Exit Function
    End If
    If RB!bloqactu = 0 Then
        'Asiento no bloqueado
        'Tratar de modificarlo
        SQL = "UPDATE cabapu set bloqactu=1 "
        SQL = SQL & " WHERE numdiari =" & ND
        SQL = SQL & " AND fechaent='" & NF
        SQL = SQL & " ' AND numasien=" & NA
        RB.Close
        On Error Resume Next
        Conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
            MostrarMensajeError = ""
        Else
            BloquearAsiento = True
        End If
        On Error GoTo 0  'quitamos los errores
    Else
        RB.Close
    End If
    Set RB = Nothing
End Function












Public Function DesbloquearAsiento(NA As String, ND As String, NF As String) As Boolean
On Error Resume Next
    SQL = "UPDATE cabapu "
    SQL = SQL & " SET bloqactu=0 "
    SQL = SQL & " WHERE numdiari =" & ND
    SQL = SQL & " AND fechaent='" & NF
    SQL = SQL & "' AND numasien=" & NA
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Desbloqueo asiento: " & NA & " /  " & ND & " /     " & NF
        DesbloquearAsiento = False
    Else
        DesbloquearAsiento = True
    End If
End Function



'-------------------------------------------------------------------

Public Function CargaDatosConExt(ByRef Cuenta As String, fec1 As Date, fec2 As Date, ByRef vSQL As String, ByRef DescCuenta As String) As Byte
Dim ACUM As Double  'Acumulado anterior

On Error GoTo ECargaDatosConExt
CargaDatosConExt = 1

'Insertamos en los campos de cabecera de cuentas
NombreSQL DescCuenta
SQL = Cuenta & "    -    " & DescCuenta
SQL = "INSERT INTO tmpconextcab (codusu,cta,fechini,fechfin,cuenta) VALUES (" & vUsu.Codigo & ", '" & Cuenta & "','" & Format(fec1, "dd/mm/yyyy") & "','" & Format(fec2, "dd/mm/yyyy") & "','" & SQL & "')"
Conn.Execute SQL


''los totatales
'Dim T1, cad
'cad = "Cuenta: " & DescCuenta & vbCrLf
'T1 = Timer


If Not CargaAcumuladosTotales(Cuenta) Then Exit Function
'cad = cad & "Acum Total:" & Format(Timer - T1, "0.000") & vbCrLf
'T1 = Timer

'Los caumulados anteriores
If Not CargaAcumuladosAnteriores(Cuenta, fec1, ACUM) Then Exit Function
'cad = cad & "Anterior:   " & Format(Timer - T1, "0.000") & vbCrLf
'T1 = Timer

'GENERAMOS LA TBLA TEMPORAL
If Not CargaTablaTemporalConExt(Cuenta, vSQL, ACUM) Then Exit Function


'cad = cad & "Tabla:    " & Format(Timer - T1, "0.000") & vbCrLf
'MsgBox cad


CargaDatosConExt = 0
Exit Function
ECargaDatosConExt:
    CargaDatosConExt = 2
    MuestraError Err.Number, "Gargando datos temporales. Cta: " & Cuenta, Err.Description
End Function



Private Function CargaAcumuladosTotales(ByRef Cta As String) As Boolean
    CargaAcumuladosTotales = False
    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    SQL = SQL & " from hlinapu where codmacta='" & Cta & "'"
    SQL = SQL & " AND fechaent >=  '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
        Else
        ImpD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
    End If
    RT.Close
    Set RT = Nothing
    SQL = "UPDATE tmpconextcab SET acumtotD= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    SQL = SQL & ", acumtotH= " & TransformaComasPuntos(CStr(ImpH)) 'Format(ImpH, "#,###,##0.00")
    ImpD = ImpD - ImpH
    SQL = SQL & ", acumtotT= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL
    CargaAcumuladosTotales = True
End Function


Private Function CargaAcumuladosAnteriores(ByRef Cta As String, ByRef FI As Date, ByRef ACUM As Double) As Boolean
Dim F1 As Date

    CargaAcumuladosAnteriores = False
    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    SQL = SQL & " from hlinapu where codmacta='" & Cta & "'"
    F1 = vParam.fechaini

    Do
        If FI < F1 Then F1 = DateAdd("yyyy", -1, F1)
    Loop Until F1 <= FI
    'SQL = SQL & " AND fechaent >=  '" & Format(vParam.fechaini, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent >=  '" & Format(F1, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <  '" & Format(FI, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
    Else
        ImpD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
    End If
    RT.Close
    ACUM = ImpD - ImpH
    SQL = "UPDATE tmpconextcab SET acumantD= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & ", acumantH= " & TransformaComasPuntos(CStr(ImpH))
    SQL = SQL & ", acumantT= " & TransformaComasPuntos(CStr(ACUM))
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL
    Set RT = Nothing
    CargaAcumuladosAnteriores = True
End Function



Private Function CargaTablaTemporalConExt(Cta As String, vSele As String, ByRef ACUM As Double) As Boolean
Dim Aux As Currency
Dim ImporteD As String
Dim ImporteH As String
Dim Contador As Long
Dim RC As String

Dim Inserts As String  'Octubre 2013. Iba muy lento


On Error GoTo Etmpconext


'TIEMPOS
'Dim T1, Cadenita
'T1 = Timer
'Cadenita = "Cuenta: " & Cta & vbCrLf

CargaTablaTemporalConExt = False

'Conn.Execute "Delete from tmpconext where codusu =" & vUsu.Codigo
Set RT = New ADODB.Recordset
SQL = "Select * from hlinapu where codmacta='" & Cta & "'"
SQL = SQL & " AND " & vSele & " ORDER BY fechaent,numasien,linliapu"
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

'Cadenita = Cadenita & "Select: " & Format(Timer - T1, "0.0000") & vbCrLf
'T1 = Timer

SQL = "INSERT INTO tmpconext (codusu, POS,numdiari, fechaent, numasien, linliapu, timporteD, timporteH, saldo, Punteada,nomdocum,ampconce,cta,contra,ccost) VALUES "
'ImpD = 0 ASI LLEVAMOS EL ACUMULADO
'ImpH = 0
Contador = 0
Inserts = ""
While Not RT.EOF
    Contador = Contador + 1
    If Not IsNull(RT!timported) Then
        Aux = RT!timported
        ImpD = ImpD + Aux
        ImporteD = TransformaComasPuntos(RT!timported)
        ImporteH = "Null"
    Else
        Aux = RT!timporteH
        ImporteD = "Null"
        ImporteH = TransformaComasPuntos(RT!timporteH)
        ImpH = ImpH + Aux
        Aux = -1 * Aux
    End If
    ACUM = ACUM + Aux
    
    'Insertar
    RC = vUsu.Codigo & "," & Contador & "," & RT!NumDiari & ",'" & Format(RT!fechaent, FormatoFecha) & "'," & RT!Numasien & "," & RT!Linliapu & ","
    RC = RC & ImporteD & "," & ImporteH
    If RT!punteada <> 0 Then
        ImporteD = "SI"
        Else
        ImporteD = ""
    End If
    RC = RC & "," & TransformaComasPuntos(CStr(ACUM)) & ",'" & ImporteD & "','"
    RC = RC & DevNombreSQL(RT!numdocum) & "','" & DevNombreSQL(RT!ampconce) & "','" & Cta & "',"
    If IsNull(RT!ctacontr) Then
        RC = RC & "NULL"
    Else
        RC = RC & "'" & RT!ctacontr & "'"
    End If
    RC = RC & ","
    If IsNull(RT!codccost) Then
        RC = RC & "NULL"
    Else
        RC = RC & "'" & RT!codccost & "'"
    End If
    RC = RC & ")"
    
    
    'octubre 2013
    Inserts = Inserts & ", (" & RC
    If (Contador Mod 150) = 0 Then
        
        Inserts = Mid(Inserts, 2)
        Conn.Execute SQL & Inserts
        Inserts = ""
    End If
    'Sig
    RT.MoveNext
Wend
RT.Close

If Inserts <> "" Then
    Inserts = Mid(Inserts, 2)
    Conn.Execute SQL & Inserts
End If


'Cadenita = Cadenita & "Recorrer: " & Format(Timer - T1, "0.0000") & vbCrLf
'T1 = Timer

    SQL = "UPDATE tmpconextcab SET acumperD= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & ", acumperH= " & TransformaComasPuntos(CStr(ImpH))
    ImpD = ImpD - ImpH
    SQL = SQL & ", acumperT= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL

    CargaTablaTemporalConExt = True
    
'Cadenita = Cadenita & "Actualizar: " & Format(Timer - T1, "0.0000") & vbCrLf
'MsgBox Cadenita
Exit Function
Etmpconext:
    MuestraError Err.Number, "Generando datos saldos"
    Set RT = Nothing
End Function







'--------------------------------------------------------
'  BALANCE DE SUMAS Y SALDOS
'--------------------------------------------------------
Public Sub CargaBalanceNuevo(ByRef Cta As String, NomCuenta As String, ConApertura As Boolean, ByRef Mes1 As Integer, ByRef Mes2 As Integer, ByRef Anyo1 As Integer, ByRef Anyo2 As Integer, FechInicioEsMesInicio As Boolean, F_Ini As Date, F_Fin As Date, EjerciCerrados As Boolean, QuitarCierre As Byte, vContabili As Integer, DesdeBalancesConfigurados As Boolean, Resetea6y7 As Boolean, RecordSetPrecargado As Boolean)
'FechInicioEsMesInicio ->  QUiere decir que si el mes incio que he puesto coincide
'                          con la fecha incio entoces,  no calcularemos anteriores
'                          y si ademas desglosamos la apertura, se la restaremos a
'                          los moviemientos del periodo, NO al anterior
'
'
'
'  QUitarCierre  :  0.- NO
'                   1.- Ambos
'                   2.- Solo perdidas y ganancias
'                   3.- Cierre
'   RecordSetPrecargado ....
'                   Si precargamos el RS significa que antes de lanzar este proceso cargamos un RS
'                   con los valores de la apertura (Y O CIERRE)
'
Dim miSQL As String
Dim ActualD As Currency
Dim ActualH As Currency
Dim NuloAC As Boolean   'Del actual
Dim NuloAC1 As Boolean  'Si huberia del siguiente
Dim NuloPer As Boolean
Dim NuloAper As Boolean
Dim CalcularImporteAnterior As Boolean
    M1 = Mes1
    M2 = Mes2
    A1 = Anyo1
    A2 = Anyo2
    vCta = Cta

    
    Contabilidad = vContabili
    NombreSQL NomCuenta
    
    miSQL = "INSERT INTO Usuarios.ztmpbalancesumas (codusu,"
    miSQL = miSQL & "cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES (" & vUsu.Codigo
    miSQL = miSQL & ",'" & vCta & "','" & NomCuenta & "',"
    NuloAper = True
    If ConApertura Then
        ObtenerApertura EjerciCerrados, F_Ini, F_Fin, NuloAper
        'En impd y imph tendremos los saldos
    Else
        ImpD = 0
        ImpH = 0
    End If
    'Para la cadena de insercion
    'Modificacion 1 Junio 2004. -> Ver A_versiones
    
    
'    If ImpD = 0 Then
'        d = "NULL"
'        Else
'        d = TransformaComasPuntos(CStr(ImpD))
'    End If
'    If ImpH = 0 Then
'        H = "NULL"
'    Else
'        H = TransformaComasPuntos(CStr(ImpH))
'    End If

    d = TransformaComasPuntos(CStr(ImpD))
    H = TransformaComasPuntos(CStr(ImpH))
    

    
'    Etoy aqui, comprobando si las variabels impD, umpH
'    las puedo utiliar despues, si me han pedido quitar los saldos
'    de cierre y PyG
'
'    Bajo hay una fucnion k calculara, si lo pide, y el mes fin es fin ejercicio
'    los saldos. K luego habra k restar. El problema es saber k variables utilizar
'    y k la funcion los calcule bien
    
    miSQL = miSQL & d & "," & H & ","
    
    '----------------------------
    'Calcula Acumulados Anteriores

    

    
    'Si es el ejercicio siguiente, es decir NO estamos en cerrados
    'Vemos todos los saldos
    ActualD = 0: ActualH = 0
    NuloAC = True
    CalcularImporteAnterior = False
    If Not EjerciCerrados Then
        If vParam.fechafin < F_Ini Then
            If Resetea6y7 Then
                If Mid(Cta, 1, 1) = vParam.grupogto Or Mid(Cta, 1, 1) = vParam.grupovta Then
                    CalcularImporteAnterior = False
                Else
                    CalcularImporteAnterior = True
                    If vParam.grupoord <> "" Then
                        If Mid(Cta, 1, 1) = vParam.grupoord Then
                            CalcularImporteAnterior = False
                            If vParam.Automocion <> "" Then
                                If Mid(Cta, 1, Len(vParam.Automocion)) = vParam.Automocion Then CalcularImporteAnterior = True
                            End If
                        End If
                    End If
                End If
                
            Else
                CalcularImporteAnterior = True
            End If
        End If
     End If
    

    
    If CalcularImporteAnterior Then
            
        'Estamos en ejercicio siguientes y hay que sumar todos los
        'saldos de ejercicio actual
        CalculaAcumuladosAnterioresBalance False, F_Ini, True, NuloAC
        ActualD = ImAcD
        ActualH = ImAcH

    End If
        
    'Las variabled de acumaldo hay k reestablecerlas
    ImAcD = 0: ImAcH = 0
    NuloAC1 = True
    If Not FechInicioEsMesInicio Then
        CalculaAcumuladosAnterioresBalance EjerciCerrados, F_Ini, False, NuloAC1
        ImAcD = ImAcD - ImpD
        ImAcH = ImAcH - ImpH
    End If
    NuloAC = NuloAC1 And NuloAC
    ImAcD = ActualD + ImAcD
    ImAcH = ActualH + ImAcH

    
   'Para la cadena de insercion
'    If ImAcD = 0 Then
'        d = "NULL"
'        Else
'        d = TransformaComasPuntos(CStr(ImAcD))
'    End If
'    If ImAcH = 0 Then
'        H = "NULL"
'    Else
'        H = TransformaComasPuntos(CStr(ImAcH))
'    End If
'
'

    d = TransformaComasPuntos(CStr(ImAcD))
    H = TransformaComasPuntos(CStr(ImAcH))

    
    
    miSQL = miSQL & d & "," & H & ","
    
    
    'Calcula moviemientos periodo
    MoviemientosPeridoBalance EjerciCerrados, F_Ini, F_Fin, NuloPer
    If FechInicioEsMesInicio Then
        'Le restamos los movimientos del desglose apertura
        ImPerD = ImPerD - ImpD
        ImPerH = ImPerH - ImpH
    End If
    
    '--------------------------------------------------------
    'Nuevo: 19 de Mayo de 2003
    'Ahora, le restamos, si  asi lo pide, y si se puede, el perdidads y ganacias y cierre
    'Meteremos los valores en imacd imach
    If QuitarCierre > 0 Then
        'Modificacion 24 Noviembre
        
        If RecordSetPrecargado Then
            'Esta es la mod.
            'Tendre un RS ya cargado con los valores, y el lo que antes era un RS.open
            'ahoa sera un RS.find
            BuscarValorEnPrecargado vCta
        Else
        
        
            ObtenerPerdidasyGanancias EjerciCerrados, F_Ini, F_Fin, QuitarCierre  'El 1 significa los dos pyg   y cierre
        End If
            ImPerD = ImPerD - ImCierrD
            ImPerH = ImPerH - ImCierrH
    End If
    
    
    
    
    
    
    'Para la cadena de insercion
'    If ImPerD = 0 Then
'        d = "NULL"
'        Else
'        d = TransformaComasPuntos(CStr(ImPerD))
'    End If
'    If ImPerH = 0 Then
'        H = "NULL"
'    Else
'        H = TransformaComasPuntos(CStr(ImPerH))
'    End If
    d = TransformaComasPuntos(CStr(ImPerD))
    H = TransformaComasPuntos(CStr(ImPerH))
    miSQL = miSQL & d & "," & H & ","

    If ImpD = 0 And ImAcD = 0 And ImpH = 0 And ImAcH = 0 Then
        If ImPerD = 0 And ImPerH = 0 Then
            NuloPer = NuloPer And NuloAC And NuloAper
            'If Not TieneMoviemientosPeriodoBalance(EjerciCerrados, F_Ini, F_Fin) Then Exit Sub
            If NuloPer Then Exit Sub
        End If
    End If
    'El saldo sera
    'Apertura k esta en impd
    ' anterior que esta en imacd
    ' periodo que esta en imperd
    ImpD = ImpD + ImAcD + ImPerD
    ImpH = ImpH + ImAcH + ImPerH
    
    
    'Si estamos en balnces configurados entonces no necesito insertar en la BD
    'Lo  Unico k kiero son los valores imd y imph
    If DesdeBalancesConfigurados Then Exit Sub
    
    
    
    'Si vengo para mostarar el balance de sumas y slados entocnes sigo y luego imprimire
    If ImpD >= ImpH Then
        ImpD = ImpD - ImpH
        miSQL = miSQL & TransformaComasPuntos(CStr(ImpD)) & ",NULL)"
    Else
        ImpH = ImpH - ImpD
        miSQL = miSQL & "NULL," & TransformaComasPuntos(CStr(ImpH)) & ")"
    End If
    
    Conn.Execute miSQL
End Sub




Private Sub CalculaAcumuladosAnterioresBalance(EjeCerrado As Boolean, ByRef fec1 As Date, EsSiguiente As Boolean, ByRef NulAcum As Boolean)


    SQL = "SELECT Sum(impmesde) AS SumaDetimporteD, Sum(impmesha) AS SumaDetimporteH"
    SQL = SQL & " from "
    If Contabilidad >= 0 Then SQL = SQL & " conta" & Contabilidad & "."
    SQL = SQL & "hsaldos"
    If EjeCerrado Then SQL = SQL & "1"
    SQL = SQL & " where "
    
    
    If Not EsSiguiente Then
        'NORMAL ----------------
        'Año natural
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            Aux = " codmacta='" & vCta & "' AND anopsald = " & A1 & " AND mespsald >=" & Month(fec1) & " AND mespsald < " & M1
            SQL = SQL & Aux
        Else
            'Año partido
            If Year(fec1) = A1 Then
                'Esta dentro del mismo año
                Aux = " codmacta='" & vCta & "' AND anopsald = " & A1 & " "
                Aux = Aux & " AND mespsald >= " & Month(fec1) & " AND mespsald <" & M1
                SQL = SQL & Aux
            Else
                Aux = " codmacta='" & vCta & "' AND  anopsald =" & Year(fec1) & " and mespsald >=" & Month(fec1)
                SQL = SQL & " (" & Aux & ") OR ("
                
                
                
                'Aux = " codmacta='" & vCta & "' AND anopsald = " & A1 & " AND mespsald <=" & M1
                
                
                '25 Abril 2005
                ' MENOR MENOR MENOR, no menor o igual
                Aux = " codmacta='" & vCta & "' AND anopsald = " & A1 & " AND mespsald <" & M1
                
                SQL = SQL & Aux & ")"
            End If
        End If
        
    Else
        'Saldos para ejercicios siguiente
        'Para k acumule el saldo
        SQL = SQL & " codmacta='" & vCta & "' "
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            'AÑO NATURAL
            Aux = " anopsald = " & Year(vParam.fechaini)
            SQL = SQL & " AND (" & Aux & ")"
        Else
                Aux = " anopsald =" & Year(vParam.fechaini) & " and mespsald >=" & Month(vParam.fechaini)
                SQL = SQL & " AND ((" & Aux & ") OR ("
                'Aux = " anopsald = " & Year(vParam.fechafin) & " AND mespsald <" & Month(vParam.fechafin)
                Aux = " anopsald = " & Year(vParam.fechafin) & " AND mespsald <=" & Month(vParam.fechafin)
                SQL = SQL & Aux & "))"
        End If
    
    End If
    Nulo1 = True
    Nulo2 = True
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImAcD = 0
    Else
        ImAcD = RT.Fields(0)
        Nulo1 = False
    End If
    If IsNull(RT.Fields(1)) Then
        ImAcH = 0
    Else
        ImAcH = RT.Fields(1)
        Nulo2 = False
    End If
    NulAcum = Nulo1 And Nulo2
    RT.Close
End Sub



Private Sub MoviemientosPeridoBalance(Cerrado As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, ByRef NuloPerio As Boolean)
    SQL = "SELECT Sum(impmesde) AS SumaDetimporteD, Sum(impmesha) AS SumaDetimporteH"
    'Modificacion para las cuentas k tienen movimientos positivos y negativos y
    SQL = SQL & " from "
    If Contabilidad >= 0 Then SQL = SQL & " conta" & Contabilidad & "."
    SQL = SQL & "hsaldos"
    If Cerrado Then SQL = SQL & "1"
    
'    SQL = SQL & " where codmacta='" & vCta & "'"
'
'    'Años
'    If Year(fec1) = Year(fec2) Then
'        'Es el mismo año, luego
'        Aux = " (anopsald = " & Year(fec1) & " AND ("
'        Aux = Aux & " mespsald >=  " & M1
'        Aux = Aux & " AND mespsald <=  " & M2 & "))"
'
'    Else
'        'Añso partidos, ejercicios partidos
'        'Desde mes año inico
'        Aux = "( anopsald = " & A1 & " AND mespsald>=" & M1 & ")"
'        Aux = Aux & " OR "
'        Aux = Aux & " ( anopsald =" & A2 & " AND mespsald <=" & M2 & ")"
'    End If
    
    
    SQL = SQL & " where " ' codmacta='" & vCta & "'"
    
    'Modificacion Febrero 2004
    If Year(fec1) = Year(fec2) Then
        'Es el mismo año, luego
        Aux = " codmacta='" & vCta & "' AND anopsald = " & Year(fec1) & " AND "
        Aux = Aux & " mespsald >=  " & M1
        Aux = Aux & " AND mespsald <=  " & M2
        
    Else
        'Añso partidos, ejercicios partidos
        'Desde mes año inico
        If A1 = A2 Then
            'Ha pedido desde un año hsta parte del otro
            Aux = " codmacta='" & vCta & "' AND anopsald = " & A1 & " AND "
            Aux = Aux & " mespsald >=  " & M1
            Aux = Aux & " AND mespsald <=  " & M2
        Else
            Aux = "( codmacta='" & vCta & "' AND anopsald = " & A1 & " AND mespsald>=" & M1 & ")"
            Aux = Aux & " OR "
            Aux = Aux & "(codmacta='" & vCta & "' AND anopsald = " & A2 & " AND mespsald <=" & M2 & ")"
        End If
    End If
    
    
    SQL = SQL & Aux
    Nulo1 = False
    Nulo2 = False
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImPerD = 0
        Nulo1 = True
    Else
        ImPerD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImPerH = 0
        Nulo2 = True
    Else
        ImPerH = RT.Fields(1)
    End If
    NuloPerio = Nulo1 And Nulo2
    RT.Close
End Sub



Private Function TieneMoviemientosPeriodoBalance(Cerrado As Boolean, ByRef fec1 As Date, ByRef fec2 As Date) As Boolean
Dim I As Integer


    TieneMoviemientosPeriodoBalance = False
    
    
    SQL = "SELECT count(*)"
    SQL = SQL & " from "
    If Contabilidad >= 0 Then SQL = SQL & " conta" & Contabilidad & "."
    SQL = SQL & "hlinapu"
    If Cerrado Then SQL = SQL & "1"
        
    
    SQL = SQL & " where " ' codmacta='" & vCta & "'"
    
'    'Modificacion Febrero 2004
'    If Year(fec1) = Year(fec2) Then
'        'Es el mismo año, luego
'        Aux = " codmacta='" & vCta & "' AND anopsald = " & Year(fec1) & " AND "
'        Aux = Aux & " mespsald >=  " & M1
'        Aux = Aux & " AND mespsald <=  " & M2
'
'    Else
'
'            'Ha pedido desde un año hsta parte del otro
'            Aux = " codmacta='" & vCta & "' AND anopsald = " & A1 & " AND "
'            Aux = Aux & " mespsald >=  " & M1
'            Aux = Aux & " AND mespsald <=  " & M2
'        Else
'            Aux = "( codmacta='" & vCta & "' AND anopsald = " & A1 & " AND mespsald>=" & M1 & ")"
'            Aux = Aux & " OR "
'            Aux = Aux & "(codmacta='" & vCta & "' AND anopsald = " & A2 & " AND mespsald <=" & M2 & ")"
'        End If
'    End If
'
    
    I = Len(vCta)
    If I = vEmpresa.DigitosUltimoNivel Then
        SQL = SQL & " codmacta='" & vCta & "'"
    Else
        SQL = SQL & " codmacta like '" & vCta & Mid("__________", 1, vEmpresa.DigitosUltimoNivel - I) & "'" 'Nivel
    End If
    I = DiasMes(CByte(M2), A2)
    SQL = SQL & " AND fechaent>='" & A1 & "-" & Format(M1, "00") & "-01' AND fechaent<='" & A2 & "-" & M2 & "-" & Format(I, "00") & "'"
    
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then
            If RT.Fields(0) > 0 Then TieneMoviemientosPeriodoBalance = True
        End If
    End If
    RT.Close
    
End Function





Private Sub ObtenerApertura(EjerCerrados As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, ByRef NulAper As Boolean)
Dim Aux As String

    'El movimietno de apertura se clacula mirando el asiento de apertura (codigo
    'concepto 970)
    SQL = "SELECT Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    If EsCuentaUltimoNivel(vCta) Then
        Aux = vCta
    Else
        Aux = vCta & "%"
    End If
    
    SQL = SQL & " from "
    If Contabilidad >= 0 Then SQL = SQL & " conta" & Contabilidad & "."
    SQL = SQL & "hlinapu"
    If EjerCerrados Then SQL = SQL & "1"
    SQL = SQL & " where codmacta like '" & Aux & "'"
    SQL = SQL & " and fechaent >='" & Format(fec1, FormatoFecha) & "'"
    SQL = SQL & " and fechaent <='" & Format(fec2, FormatoFecha) & "'"
    SQL = SQL & " AND codconce= 970" '970 es el asiento de apertura
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Nulo1 = True
    Nulo2 = True
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
    Else
        ImpD = RT.Fields(0)
        Nulo1 = False
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
        Nulo2 = False
    End If
    NulAper = Nulo1 And Nulo2
    RT.Close
End Sub




Public Function AgrupacionCtasBalance(Codigo As String, nommacta As String) As Boolean
Dim C As Integer
On Error GoTo EAgrupacionCtasBalance

    AgrupacionCtasBalance = False
    ImAcD = 0
    ImAcH = 0
    ImPerD = 0
    ImPerH = 0
    ImCierrD = 0
    ImCierrH = 0
    ImpD = 0
    ImpH = 0
    vCta = Mid(Codigo & "__________", 1, vEmpresa.DigitosUltimoNivel)
    
    SQL = "Select * from Usuarios.ztmpbalancesumas where codusu =" & vUsu.Codigo
    SQL = SQL & " AND cta like '" & vCta & "'"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = 0
    While Not RT.EOF
        'Apertura
        ImAcD = ImAcD + DBLet(RT.Fields(3), "N")
        ImAcH = ImAcH + DBLet(RT.Fields(4), "N")
        'anterior
        ImPerD = ImPerD + DBLet(RT.Fields(5), "N")
        ImPerH = ImPerH + DBLet(RT.Fields(6), "N")
        'periodo
        ImCierrD = ImCierrD + DBLet(RT.Fields(7), "N")
        ImCierrH = ImCierrH + DBLet(RT.Fields(8), "N")
        'Total
        ImpD = ImpD + DBLet(RT.Fields(9), "N")
        ImpH = ImpH + DBLet(RT.Fields(10), "N")
        
        RT.MoveNext
        C = C + 1
    Wend
    RT.Close
    If C = 0 Then
        AgrupacionCtasBalance = True
        Exit Function
    End If
    
    'Acumulamos saldo en uno de los lados
    If ImpD > ImpH Then
        ImpD = ImpD - ImpH
        ImpH = 0
    Else
        ImpH = ImpH - ImpD
        ImpD = 0
    End If
    
    
    'Borramos las entradas
    SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu =" & vUsu.Codigo
    SQL = SQL & " AND cta like '" & vCta & "'"
    Conn.Execute SQL
    Conn.Execute "commit"
    espera 0.5
    
    SQL = "INSERT INTO Usuarios.ztmpbalancesumas (codusu, cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES (" & vUsu.Codigo
    Aux = Mid(Codigo & "**********", 1, vEmpresa.DigitosUltimoNivel)
    SQL = SQL & ",'" & Aux & "','" & Mid("AGRUP- " & nommacta, 1, 30) & "'"
    SQL = SQL & ImporteASQL(ImAcD) & ImporteASQL(ImAcH) & ImporteASQL(ImPerD) & ImporteASQL(ImPerH)
    SQL = SQL & ImporteASQL(ImCierrD) & ImporteASQL(ImCierrH) & ImporteASQL(ImpD) & ImporteASQL(ImpH) & ")"
    Conn.Execute SQL
    AgrupacionCtasBalance = True
    Exit Function
EAgrupacionCtasBalance:
    MuestraError Err.Number, "Agrupacion Ctas Balance"
End Function










' MAYO 2004.  Vamos a poder separar dependiendo del tipo de llamada
'         0.- No llegara hasta aqui
'         1.- Los dos  pérdidas/ganancias   y   cierre
'         2.- Perdidas y GAnancias
'         3.- Solo Cierre
Private Sub ObtenerPerdidasyGanancias(EjerCerrados As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, OpcionBusqueda As Byte)
Dim Aux As String

    'Perdidas y ganancias: 960
    'Cierre             : 980
    
    SQL = "SELECT Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    If EsCuentaUltimoNivel(vCta) Then
        Aux = vCta
    Else
        Aux = vCta & "%"
    End If
    SQL = SQL & " from "
    If Contabilidad >= 0 Then SQL = SQL & " conta" & Contabilidad & "."
    SQL = SQL & "hlinapu"
    If EjerCerrados Then SQL = SQL & "1"
    SQL = SQL & " where codmacta like '" & Aux & "'"
    SQL = SQL & " and fechaent >='" & Format(fec1, FormatoFecha) & "'"
    SQL = SQL & " and fechaent <='" & Format(fec2, FormatoFecha) & "'"
    
    '  960 P y G
    '  970 es el asiento de apertura
    '  980 Cierre
    Aux = ""
    If OpcionBusqueda < 3 Then Aux = "codconce= 960"
    If OpcionBusqueda <> 2 Then
        If Aux <> "" Then Aux = Aux & " OR "
        Aux = Aux & "codconce= 980"
    End If
    Aux = " AND (" & Aux & ")"
    SQL = SQL & Aux
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImCierrD = 0
    Else
        ImCierrD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImCierrH = 0
    Else
        ImCierrH = RT.Fields(1)
    End If
    RT.Close
End Sub


'---------------------------------------------------------
'Precarga de los datos del balance
'
Public Sub PrecargaPerdidasyGanancias(EjerCerrados As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, OpcionBusqueda As Byte)
Dim Aux As String

    'Perdidas y ganancias: 960
    'Cierre             : 980
    
    SQL = "SELECT codmacta,Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    SQL = SQL & " from "
    SQL = SQL & "hlinapu"
    If EjerCerrados Then SQL = SQL & "1"
    SQL = SQL & " where fechaent >='" & Format(fec1, FormatoFecha) & "'"
    SQL = SQL & " and fechaent <='" & Format(fec2, FormatoFecha) & "'"
    
    '  960 P y G
    '  970 es el asiento de apertura
    '  980 Cierre
    Aux = ""
    If OpcionBusqueda < 3 Then Aux = "codconce= 960"
    If OpcionBusqueda <> 2 Then
        If Aux <> "" Then Aux = Aux & " OR "
        Aux = Aux & "codconce= 980"
    End If
    Aux = " AND (" & Aux & ")"
    SQL = SQL & Aux & " GROUP BY codmacta"
        
        
    Set RsBalPerGan = New ADODB.Recordset
    RsBalPerGan.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
End Sub


Public Sub PrecargaApertura()

    SQL = "SELECT codmacta,Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    SQL = SQL & " from hlinapu"
    SQL = SQL & " where fechaent ='" & Format(vParam.fechaini, FormatoFecha) & "'"
    SQL = SQL & " AND codconce= 970 GROUP BY codmacta"
        
        
    Set RsBalPerGan = New ADODB.Recordset
    RsBalPerGan.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
End Sub



Public Sub CerrarPrecargaPerdidasyGanancias()
    RsBalPerGan.Close
    Set RsBalPerGan = Nothing
End Sub


Public Sub BuscarValorEnPrecargado(ByRef codmacta As String)


    RsBalPerGan.Find "codmacta = '" & codmacta & "'", , adSearchForward, 1
    If RsBalPerGan.EOF Or RsBalPerGan.BOF Then
        ImCierrD = 0
        ImCierrH = 0
    Else
        If IsNull(RsBalPerGan.Fields(1)) Then
            ImCierrD = 0
        Else
            ImCierrD = RsBalPerGan.Fields(1)
        End If
        If IsNull(RsBalPerGan.Fields(2)) Then
            ImCierrH = 0
        Else
            ImCierrH = RsBalPerGan.Fields(2)
        End If
    End If
End Sub






'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'
' Cuentas de explotacion
'
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'--->  OPCION  0.- Con anterior y movimientos     1.- Solo SALDO
'El ctaSQL es para no tener que copiar el SQL de insertar
Public Sub CuentaExplotacion(ByRef Cta As String, ByRef Titulo As String, ByRef Mes As Integer, ByRef Anyo As Integer, ByRef Contador As Long, Opcion As Byte, ByRef CtaSQL As String, Cerrados As Boolean, ByRef FecIni As Date, ByRef FecFin As Date, QuitarSaldos As Boolean, Contabi As Integer)
'Dim ImAcD As Double     acumulado
'Dim ImAcH As Double
'Dim ImPerD As Double    periodo
'Dim ImPerH As Double
' ImCierrD As Currency  'importes
'ImCierrH As Currency

'impd                    SALDO
'imph

    M1 = Mes
    A1 = Anyo
    vCta = Cta
    Contabilidad = Contabi
    
    Aux = "'" & Cta & "'," & Contador & ",'" & DevNombreSQL(Titulo) & "',"
    
    Set RT = New ADODB.Recordset
    If QuitarSaldos Then
        CalcularImporteCierreCtaExplotacion Cerrados, FecFin
        'ImCierrD = 0
        'ImCierrH = 0
    Else
        ImCierrD = 0
        ImCierrH = 0
    End If
    
    
    CalcularSaldoCtaExplotacion True, Cerrados, FecIni
    CalcularSaldoCtaExplotacion False, Cerrados, FecIni
    If ImAcD = 0 And ImAcH = 0 And ImPerD = 0 And ImPerH = 0 And ImCierrD = 0 And ImCierrH = 0 Then
        'En este caso no hacemos insercion
        
    Else
        
        
        'Para la cadena de insercion ANTERIOR
        'Modificacion 11 Junio. Esta quitado lo del NULL
        'If ImAcD = 0 Then
        '    d = "NULL"
        '    Else
            d = TransformaComasPuntos(CStr(ImAcD))
        'End If
        'If ImAcH = 0 Then
        '    H = "NULL"
        'Else
            H = TransformaComasPuntos(CStr(ImAcH))
        'End If
        
        
        
        Aux = Aux & d & "," & H & ","
        
        'Creo k es aqui donde kitamos los importes del cierre
        ImPerD = ImPerD - ImCierrD
        ImPerH = ImPerH - ImCierrH
        
        
        'Para la cadena de insercion PERIODO
        'Modificacion 10 Junio
        'If ImPerD = 0 Then
        '    d = "NULL"
        '    Else
            d = TransformaComasPuntos(CStr(ImPerD))
        'End If
        'If ImPerH = 0 Then
        '    H = "NULL"
        'Else
            H = TransformaComasPuntos(CStr(ImPerH))
        'End If
        Aux = Aux & d & "," & H & ","
        
        
        ImpD = ImAcD + ImPerD
        ImpH = ImAcH + ImPerH
        
        If ImpD > ImpH Then
            ImpD = ImpD - ImpH
            ImpH = 0
        Else
            ImpH = ImpH - ImpD
            ImpD = 0
        End If
        
        'Para la cadena de insercion
        'If ImpD = 0 Then
        '    d = "NULL"
        '    Else
            d = TransformaComasPuntos(CStr(ImpD))
        'End If
        'If ImpH = 0 Then
        '    H = "NULL"
        'Else
            H = TransformaComasPuntos(CStr(ImpH))
        'End If
        Aux = Aux & d & "," & H & ")"
        
        Conn.Execute CtaSQL & Aux
    End If
    Contador = Contador + 1
Set RT = Nothing
End Sub





Private Sub CalcularSaldoCtaExplotacion(Anterior As Boolean, Cerrados As Boolean, ByRef FI As Date)

    SQL = "Select SUM(impmesde) as Debe,sum(impmesha) from "
    If Contabilidad > 0 Then SQL = SQL & "conta" & Contabilidad & "."
    SQL = SQL & "hsaldos"
    If Cerrados Then SQL = SQL & "1"
    SQL = SQL & " WHERE "
    
    If Anterior Then
        'Si estamos en meses partidos
        'If M1 < Month(vParam.fechaini) Then ---> antes
        'Sirve pq realmente nunca puede ser menor, pero voy a poner la de las fechas
        'If M1 < Month(FI) Then
        If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
            'AÑOS PARTIDOS
            
            If A1 = Year(FI) Then
                SQL = SQL & " codmacta ='" & vCta & "' AND mespsald >=" & Month(FI) & " AND mespsald < " & M1 & " AND anopsald =" & A1
            Else
                'El trozo del primer año
                SQL = SQL & "( codmacta ='" & vCta & "' AND mespsald >=" & Month(FI) & " AND anopsald =" & A1 - 1 & ")"
                SQL = SQL & " OR ( codmacta ='" & vCta & "' AND mespsald <" & M1 & " AND anopsald =" & A1 & ")"
            End If
            'SQL = SQL & " AND (( mespsald between " & Month(FI) & " AND 12) and anopsald = " & A1 - 1 & ")"
            'SQL = SQL & " OR ( mespsald between  1  AND " & M1 - 1 & ") and anopsald = " & A1 & ")"
            
            'El ejercicio anterior
            'Entonces si el mes k pide es mayor k el fin de ejercicio
            
            'SQL = SQL & "( codmacta ='" & vCta & "' AND mespsald >=" & Month(FI) & " AND anopsald =" & A1 -1)
        Else
        
            SQL = SQL & " codmacta ='" & vCta & "'  AND (( mespsald between " & Month(FI) & " AND " & M1 - 1 & ") and anopsald = " & A1 & ")"
        End If
    Else
        'Saldo del periodo
        SQL = SQL & " codmacta ='" & vCta & "' AND mespsald =" & M1 & " AND anopsald =" & A1
    End If
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Anterior Then
        'ACumulado anterior
        If IsNull(RT.Fields(0)) Then
            ImAcD = 0
        Else
            ImAcD = RT.Fields(0)
        End If
        If IsNull(RT.Fields(1)) Then
            ImAcH = 0
        Else
            ImAcH = RT.Fields(1)
        End If
    Else
        'Periodo
        If IsNull(RT.Fields(0)) Then
            ImPerD = 0
        Else
            ImPerD = RT.Fields(0)
        End If
        If IsNull(RT.Fields(1)) Then
            ImPerH = 0
        Else
            ImPerH = RT.Fields(1)
        End If
    End If
    RT.Close
End Sub



'Calculamos los importes de los cierres para obtener la consulta sin ellos
Private Sub CalcularImporteCierreCtaExplotacion(Cerrados As Boolean, ByRef fFin As Date)

    SQL = "Select SUM(timporteD),sum(timporteH) from "
    If Contabilidad > 0 Then SQL = SQL & "conta" & Contabilidad & "."
    SQL = SQL & "hlinapu"
    If Cerrados Then SQL = SQL & "1"
    SQL = SQL & " WHERE codmacta  like '" & vCta
    If Len(vCta) <> vEmpresa.DigitosUltimoNivel Then SQL = SQL & "%"
    SQL = SQL & "' and codconce ="
    d = Mid(vCta, 1, 1)
    If d = vParam.grupogto Or d = vParam.grupovta Or d = vParam.grupoord Then
        SQL = SQL & "960" 'perdidas y ganacias
    Else
        SQL = SQL & "980" ' cierre
    End If
    SQL = SQL & " AND fechaent = '" & Format(fFin, FormatoFecha) & "'"
    SQL = SQL & ";"

    
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImCierrD = 0
    Else
        ImCierrD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImCierrH = 0
    Else
        ImCierrH = RT.Fields(1)
    End If

    RT.Close
End Sub





'//////////////////////////////////////////////////////////////////
'       LISTADO FACTURAS CLIENTES

Public Function ListadoFacturasClientes(vSQL As String, Ordenacion As String, NumeroFAC As Long, Agrupa As Boolean, Liquidacion As Boolean, MostrarRetencion As Boolean, NIFFiltro As String) As Boolean
Dim RS As Recordset
Dim NumeroFacturaImpresa As Long
Dim CodFacTemporal As Long

ListadoFacturasClientes = False
Conn.Execute "Delete  from Usuarios.ztmpfaclin where codusu= " & vUsu.Codigo
Conn.Execute "Delete  from Usuarios.ztmpresumenivafac where codusu= " & vUsu.Codigo
SQL = "select cabfact.* , nommacta,nifdatos from cabfact,cuentas WHERE "
If vSQL <> "" Then SQL = SQL & vSQL & " AND "
SQL = SQL & "cabfact.codmacta=cuentas.codmacta "

'Febrero 2013    Fitrlo por nif
If NIFFiltro <> "" Then SQL = SQL & " AND nifdatos = '" & NIFFiltro & "'"


SQL = SQL & " ORDER BY " & Ordenacion
Set RT = New ADODB.Recordset
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RT.EOF Then
    MsgBox "Ningun dato con esos parametros. ", vbExclamation
    RT.Close
    Exit Function
End If

'Con retencion
d = "INSERT INTO Usuarios.ztmpfaclin (codusu, codigo, Numfac, Fecha, cta, Cliente, NIF, Imponible, IVA, ImpIVA, Total,retencion,tipoIva) VALUES (" & vUsu.Codigo & ","
CodFacTemporal = 1
NumeroFacturaImpresa = 1
While Not RT.EOF
    'Este trozo es comun para los imponibles
    'Nº factura
    If NumeroFAC > 0 Then
        NumRegElim = NumeroFAC + NumeroFacturaImpresa - 1
        NumeroFacturaImpresa = NumeroFacturaImpresa + 1
    Else
        NumRegElim = RT!codfaccl
    End If
    
    'Antes sept 2011
    'Aux = ",'" & RT!NUmSerie & " " & Format(NumRegElim, "0000000000") & "','"
    
    'ahora                                le doy un spc blanco
    Aux = SerieNumeroFactura(12, RT!NUmSerie & " ", CStr(NumRegElim))
    Aux = ",'" & Aux & "','"
    
    'Fecha
    If Liquidacion Then
        Aux = Aux & Format(RT!fecliqcl, "dd/mm/yyyy") & "','"
    Else
        Aux = Aux & Format(RT!fecfaccl, "dd/mm/yyyy") & "','"
    End If
    'Cuenta
    Aux = Aux & RT!codmacta & "','" & DevNombreSQL(CStr(RT!nommacta)) & "','" & DBLet(RT!nifdatos) & "',"
    '---------------------------
    'Imponibles Tipo 1 . Siempre lo tiene
    vCta = TransformaComasPuntos(CStr(RT!ba1faccl)) & ",'" & Format(RT!pi1faccl, "0.00") & "'," & TransformaComasPuntos(CStr(RT!ti1faccl)) & "," & TransformaComasPuntos(CStr(RT!totfaccl))
    SQL = d & CodFacTemporal & Aux & vCta & ","
        
        
    'La retencion
    If MostrarRetencion Then
        If Not IsNull(RT!trefaccl) Then
            SQL = SQL & TransformaComasPuntos(CStr(RT!trefaccl))
        Else
            SQL = SQL & "NULL"
        End If
    Else
        SQL = SQL & "NULL"
    End If
    
    'El tipo de iva
    SQL = SQL & "," & RT!tp1faccl
    SQL = SQL & ")"
    Conn.Execute SQL
    
    
    If Not Agrupa Then Aux = ",NULL,NULL,NULL,NULL,NULL,"    'Falta el NIF
    
    
    
    'RECARGO EQUIVALENCIA
    If Not IsNull(RT!tr1faccl) Then
        CodFacTemporal = CodFacTemporal + 1
        
        vCta = "NULL,'re" & Format(RT!pr1faccl, "0.00") & "'," & TransformaComasPuntos(CStr(RT!tr1faccl)) & ",NULL"
        SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp1faccl & ")"
        Conn.Execute SQL
    End If
    
    
    
    
    'Si tiene eimponible 2
    If Not IsNull(RT!tp2faccl) Then
        CodFacTemporal = CodFacTemporal + 1
        vCta = TransformaComasPuntos(CStr(RT!ba2faccl)) & ",'" & Format(RT!pi2faccl, "0.00") & "'," & TransformaComasPuntos(CStr(RT!ti2faccl)) & ",NULL"
        SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp2faccl & ")"
        Conn.Execute SQL
        
        
        If Not IsNull(RT!tr2faccl) Then
            CodFacTemporal = CodFacTemporal + 1
            vCta = "NULL,'re" & Format(RT!pr2faccl, "0.00") & "'," & TransformaComasPuntos(CStr(RT!tr2faccl)) & ",NULL"
            SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp2faccl & ")"
            Conn.Execute SQL
        End If
        
        
        
    End If
    
    'Si tiene imponible 3
    If Not IsNull(RT!tp3faccl) Then
        CodFacTemporal = CodFacTemporal + 1
        vCta = TransformaComasPuntos(CStr(RT!ba3faccl)) & ",'" & Format(RT!pi3faccl, "0.00") & "'," & TransformaComasPuntos(CStr(RT!ti3faccl)) & ",NULL"
        SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp3faccl & ")"
        Conn.Execute SQL
        
        
        
        If Not IsNull(RT!tr3faccl) Then
            CodFacTemporal = CodFacTemporal + 1
            vCta = "NULL,'re" & Format(RT!pr3faccl, "0.00") & "'," & TransformaComasPuntos(CStr(RT!tr3faccl)) & ",NULL"
            SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp3faccl & ")"
            Conn.Execute SQL
        End If
        
        
    End If
    
    '19/08/2003   Esto no debe entrar
    'Retencion
'    If Not IsNull(RT!cuereten) Then
'        M1 = M1 + 1
'        vCta = "NULL,'" & Format(RT!retfaccl, "0.00") & "'," & TransformaComasPuntos(CStr(-RT!trefaccl)) & ",NULL"
'        SQL = d & M1 & Aux & ",'Reten.'," & vCta & ")"
'        Conn.Execute SQL
'    End If
'
    'Siguiente
    CodFacTemporal = CodFacTemporal + 1
    RT.MoveNext
Wend
RT.Close



M1 = 1
Set RS = New ADODB.Recordset


If NIFFiltro <> "" Then
    SQL = "Select codmacta from cuentas WHERE apudirec='S' AND nifdatos = '" & NIFFiltro & "'"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO `puede ser EOF
    SQL = ""
    If RS.EOF Then Err.Raise 513, "Ninguna cuenta para el NIF"
    While Not RS.EOF
        SQL = SQL & ", '" & RS!codmacta & "'"
        RS.MoveNext
    Wend
    RS.Close
    vSQL = vSQL & " AND codmacta IN (" & Mid(SQL, 2) & ")"
End If
SQL = "Select * from Tiposiva"
RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

While Not RS.EOF
    ImpD = 0
    ImCierrD = 0
    'Suma en Imp1
    SQL = " select sum(ti1faccl),sum(ba1faccl)  from cabfact WHERE  " & vSQL & "  AND tp1faccl=" & RS!codigiva
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    ImpH = 0
    ImCierrH = 0
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImCierrH = RT.Fields(1)
    End If
    RT.Close
    ImpD = ImpD + ImpH
    ImCierrD = ImCierrD + ImCierrH
    
    'Suma en Imp2
    SQL = " select sum(ti2faccl),sum(ba2faccl)  from cabfact WHERE  " & vSQL & "  AND tp2faccl=" & RS!codigiva
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    ImpH = 0
    ImCierrH = 0
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImCierrH = RT.Fields(1)
    End If
    ImpD = ImpD + ImpH
    ImCierrD = ImCierrD + ImCierrH
    RT.Close
    'Suma en Imp3
    SQL = " select sum(ti3faccl),sum(ba3faccl)  from cabfact WHERE  " & vSQL & "  AND tp3faccl=" & RS!codigiva
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    ImpH = 0
    ImCierrH = 0
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImCierrH = RT.Fields(1)
    End If
    ImpD = ImpD + ImpH
    ImCierrD = ImCierrD + ImCierrH
    RT.Close
    
    'Si el importe es <>0 entonces lo añadimos a tmpresumn
    If ImpD <> 0 Or ImCierrD <> 0 Then
        SQL = "INSERT INTO Usuarios.ztmpresumenivafac (codusu, orden, IVA, TotalIVA,sumabases,tipoiva,NombreIVA) VALUES (" & vUsu.Codigo & "," & M1
        SQL = SQL & ",'" & Format(RS!porceiva, "0.00") & "'," & TransformaComasPuntos(CStr(ImpD)) & "," & TransformaComasPuntos(CStr(ImCierrD)) & ","
        SQL = SQL & RS!codigiva & ",'" & DevNombreSQL(RS!nombriva) & "')"
        Conn.Execute SQL
        M1 = M1 + 1
    End If
    
    
   'Por si lleva recargo de equivalencia
    If RS!porcerec > 0 Then
        ImpD = 0
        SQL = " select sum(tr1faccl)  from cabfact WHERE  " & vSQL & "  AND tp1faccl=" & RS!codigiva
        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        ImpH = 0
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
        End If
        RT.Close
        ImpD = ImpD + ImpH
            
        SQL = " select sum(tr2faccl)  from cabfact WHERE  " & vSQL & "  AND tp2faccl=" & RS!codigiva
        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        ImpH = 0
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
        End If
        RT.Close
        ImpD = ImpD + ImpH
            
    
        SQL = " select sum(tr3faccl)  from cabfact WHERE  " & vSQL & "  AND tp3faccl=" & RS!codigiva
        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        ImpH = 0
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
        End If
        RT.Close
        ImpD = ImpD + ImpH
        
        
                        
            If ImpD <> 0 Then
                SQL = "INSERT INTO Usuarios.ztmpresumenivafac (codusu, orden, IVA, TotalIVA,Sumabases,TipoIva,nombreiva) VALUES (" & vUsu.Codigo & "," & M1
                SQL = SQL & ",'re" & Format(RS!porcerec, "0.00")
                SQL = SQL & "'," & TransformaComasPuntos(CStr(ImpD))
                SQL = SQL & ",0," & RS!codigiva & ",'" & DevNombreSQL(RS!nombriva) & "')"
                Conn.Execute SQL
                M1 = M1 + 1
            End If
        
        
            
    End If  'Del recargo de equivalencia
    
    
    'Sig
    RS.MoveNext
Wend
RS.Close
Set RS = Nothing
Set RT = Nothing
'Insertaremos la cabecera, con los importes
ListadoFacturasClientes = True



End Function







'//////////////////////////////////////////////////////////////////
'       LISTADO FACTURAS PROVEEDORES
'  FechaOfactura -  0.- Nº Fac
'                   1.- Fecha emision
'                   2.- Fecha liqudiacion
'
'           NIF: Solo un NIF
'           TipoIVA :
'                       -1-->Todos   Si no solo aquellas facturas que tengan ese tipo de iva
'                       >=0 Solo las que en las bases este ese tipo de iva
'                       Especificara la retencion
Public Function ListadoFacturasProveedores(vSQL As String, Ordenacion As String, NumeroFAC As Long, Agrupa As Boolean, FechaOfactura As Byte, Liquidacion As Boolean, MostrarRetencion As Boolean, NIFFiltro As String, TipoIVA As Integer) As Boolean
Dim C As String
Dim NumeroFacturaImpresa As Long
Dim RS As Recordset
Dim CodFacTemporal As Long

'Marzo 2014. IVA importacion
Dim LetrasTipoDeIVA As String  '  ND  No decucibñe   FN     IM importacion
Dim ImporAux As Currency
Dim ImporIVA As Currency
    
    ListadoFacturasProveedores = False
    Conn.Execute "Delete  from Usuarios.ztmpfaclinprov where codusu= " & vUsu.Codigo
    Conn.Execute "Delete  from Usuarios.ztmpresumenivafac where codusu= " & vUsu.Codigo
    
    'ABRIMOS LOS TIPOS DE IVA
    SQL = "Select * from Tiposiva"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    
    
    SQL = "select cabfactprov.* , nommacta,nifdatos from cabfactprov,cuentas WHERE "
    If vSQL <> "" Then SQL = SQL & vSQL & " AND "
    SQL = SQL & "cabfactprov.codmacta=cuentas.codmacta "
    'Febrero 2013    Fitrlo por nif
    If NIFFiltro <> "" Then SQL = SQL & " AND nifdatos = '" & NIFFiltro & "'"

    'JULIO 2013 Solo quiere aquellas facturas que tengan "un tipo de iva" espceficio
    If TipoIVA >= 0 Then SQL = SQL & " and (tp1facpr=" & TipoIVA & " or tp2facpr=" & TipoIVA & " or tp3facpr=" & TipoIVA & ")"

    SQL = SQL & " ORDER BY " & Ordenacion
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        MsgBox "Ningun dato con esos parametros. ", vbExclamation
        RT.Close
        Exit Function
    End If
    'Insertado los
    'MAYO 2005
    'd = "INSERT INTO Usuarios.ztmpfaclinprov (codusu, codigo, Numfac, FechaFac, FechaCon, cta, Cliente, NIF, Imponible, IVA, ImpIVA, Total,retencion) VALUES (" & vUsu.Codigo & ","
    d = "INSERT INTO Usuarios.ztmpfaclinprov (codusu, codigo, Numfac, FechaFac, FechaCon, cta, Cliente, NIF, Imponible, IVA, ImpIVA, Total,retencion,NoDeducible,TipoIva) VALUES (" & vUsu.Codigo & ","
    CodFacTemporal = 1
    NumeroFacturaImpresa = 1
    While Not RT.EOF
        'Este trozo es comun para los imponibles
        'Nº factura
        
        If NumeroFAC > 0 Then
            NumRegElim = NumeroFAC + NumeroFacturaImpresa - 1
            NumeroFacturaImpresa = NumeroFacturaImpresa + 1
        Else
            NumRegElim = RT!NumRegis
        End If
        Aux = ",'" & Format(NumRegElim, "0000000000") & "'"
                
        'Fecha o factura
        '  FechaOfactura -  0.- Nº Fac
'                   1.- Fecha emision
'                   2.- Fecha liqudiacion
        Select Case FechaOfactura
        Case 1
            C = ",'" & Format(RT!fecfacpr, "dd/mm/yyyy")
        Case 2
            If Liquidacion Then
                C = ",'" & Format(RT!fecrecpr, "dd/mm/yyyy")
            Else
                C = ",'" & Format(RT!fecliqpr, "dd/mm/yyyy")
            End If
        Case Else
            C = ",'" & DevNombreSQL(DBLet(RT!numfacpr, "T"))
        End Select
        
        'Fecha recepcion o liquidacion
        C = C & "','"
        If Liquidacion Then
            C = C & Format(RT!fecliqpr, "dd/mm/yyyy")
        Else
            C = C & Format(RT!fecrecpr, "dd/mm/yyyy")
        End If
        'Cuenta
        'C = C & "','" & RT!codmacta & "','" & DevNombreSQL(CStr(RT!nommacta)) & "','',"
        C = C & "','" & RT!codmacta & "','" & DevNombreSQL(CStr(RT!nommacta)) & "','" & DBLet(RT!nifdatos, "T") & "',"
        
        Aux = Aux & C
        '---------------------------
        'Imponibles Tipo 1 . Siempre lo tiene
        
        LetrasTipoDeIVA = DevuelveIVA_FraPro(RS, RT!tp1facpr, RT!Nodeducible = 1, RT!extranje)
        ImporAux = RT!ba1facpr
        If LetrasTipoDeIVA = "IM" Then ImporAux = 0
        ImporIVA = RT!ti1facpr
        If LetrasTipoDeIVA = "ISP" Then ImporIVA = 0    'Inversion sujeto pasivo. NO lleva IVA
        If LetrasTipoDeIVA = "COM" Then ImporIVA = 0    'Comunitarias
        
        vCta = TransformaComasPuntos(CStr(ImporAux)) & ",'" & Format(RT!pi1facpr, "0.00") & "'," & TransformaComasPuntos(CStr(ImporIVA)) & "," & TransformaComasPuntos(CStr(RT!totfacpr))
        SQL = d & CodFacTemporal & Aux & vCta & ","
        
        
        'La retencion
        If MostrarRetencion Then
            If Not IsNull(RT!trefacpr) Then
                SQL = SQL & TransformaComasPuntos(CStr(RT!trefacpr))
            Else
                SQL = SQL & "NULL"
            End If
        Else
            SQL = SQL & "NULL"
        End If
        
        vCta = ",'" & LetrasTipoDeIVA & "'," & RT!tp1facpr
        SQL = SQL & vCta & ")"
        Conn.Execute SQL
        If Agrupa Then
            'INSERT INTO Usuarios.ztmpfaclinprov (codusu, codigo, Numfac, FechaFac, FechaCon,
            'cta, Cliente, NIF, Imponible, IVA, ImpIVA, Total) VALUES (1001,
            'En aux va desde Numfac, importa solo codmacta
            Aux = ",NULL,NULL,NULL,'" & RT!codmacta & "',NULL,NULL,"
        Else
            Aux = ",NULL,NULL,NULL,NULL,NULL,NULL,"
        End If
        
        
        
        
        
        'Recargo de equivalencia
        'RECARGO EQUIVALENCIA
        If Not IsNull(RT!tr1facpr) Then
            CodFacTemporal = CodFacTemporal + 1
            vCta = "NULL,'re" & Format(RT!pr1facpr, "0.00") & "'," & TransformaComasPuntos(CStr(RT!tr1facpr)) & ",NULL,NULL"
            SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp1facpr & ")"
            Conn.Execute SQL
        End If
            
        
        
        'Si tiene eimponible 2
        If Not IsNull(RT!tp2facpr) Then
            LetrasTipoDeIVA = DevuelveIVA_FraPro(RS, RT!tp2facpr, RT!Nodeducible = 1, RT!extranje)
            ImporAux = RT!ba2facpr
            If LetrasTipoDeIVA = "IM" Then ImporAux = 0
            ImporIVA = RT!ti2facpr
            If LetrasTipoDeIVA = "ISP" Then ImporIVA = 0    'Inversion sujeto pasivo. NO lleva IVA
            If LetrasTipoDeIVA = "COM" Then ImporIVA = 0    'Comunitarias
            
            CodFacTemporal = CodFacTemporal + 1
            vCta = TransformaComasPuntos(CStr(ImporAux)) & ",'" & Format(RT!pi2facpr, "0.00") & "'," & TransformaComasPuntos(CStr(ImporIVA)) & ",NULL"
            C = ",'" & LetrasTipoDeIVA
            C = C & "'," & RT!tp2facpr
            SQL = d & CodFacTemporal & Aux & vCta & ",NULL" & C & ")"
            Conn.Execute SQL
            
            'Recargo equivalencia 2
            If Not IsNull(RT!tr2facpr) Then
                CodFacTemporal = CodFacTemporal + 1
                vCta = "NULL,'re" & Format(RT!pr2facpr, "0.00") & "'," & TransformaComasPuntos(CStr(RT!tr2facpr)) & ",NULL,NULL"
                SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp2facpr & ")"
                Conn.Execute SQL
            End If
            
            
            
            
        End If
        
        'Si tiene imponible 3
        If Not IsNull(RT!tp3facpr) Then
            LetrasTipoDeIVA = DevuelveIVA_FraPro(RS, RT!tp3facpr, RT!Nodeducible = 1, RT!extranje)
            ImporAux = RT!ba3facpr
            If LetrasTipoDeIVA = "IM" Then ImporAux = 0
            ImporIVA = RT!ti3facpr
            If LetrasTipoDeIVA = "ISP" Then ImporIVA = 0    'Inversion sujeto pasivo. NO lleva IVA
            If LetrasTipoDeIVA = "COM" Then ImporIVA = 0    'Comunitarias
            
            CodFacTemporal = CodFacTemporal + 1
            vCta = TransformaComasPuntos(CStr(ImporAux)) & ",'" & Format(RT!pi3facpr, "0.00") & "'," & TransformaComasPuntos(CStr(ImporIVA)) & ",NULL"
            C = ",'" & LetrasTipoDeIVA
            C = C & "'," & RT!tp3facpr
            SQL = d & CodFacTemporal & Aux & vCta & ",NULL" & C & ")"
            Conn.Execute SQL
            
            
            
            'Recargo equivalencia 3
            If Not IsNull(RT!tr3facpr) Then
                CodFacTemporal = CodFacTemporal + 1
                vCta = "NULL,'re" & Format(RT!pr3facpr, "0.00") & "'," & TransformaComasPuntos(CStr(RT!tr3facpr)) & ",NULL,NULL"
                SQL = d & CodFacTemporal & Aux & vCta & ",NULL," & RT!tp3facpr & ")"
                Conn.Execute SQL
            End If
            
        End If
        
        

        'Siguiente
        CodFacTemporal = CodFacTemporal + 1
        RT.MoveNext
    Wend
    RT.Close
    
    d = RT.Source 'ME guardo el SELECT
    
    
    
    If NIFFiltro <> "" Then
        SQL = "Select codmacta from cuentas WHERE apudirec='S' AND nifdatos = '" & NIFFiltro & "'"
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO `puede ser EOF
        SQL = ""
        If RT.EOF Then Err.Raise 513, "Ninguna cuenta para el NIF"
        While Not RT.EOF
            SQL = SQL & ", '" & RT!codmacta & "'"
            RT.MoveNext
        Wend
        RT.Close
        vSQL = vSQL & " AND codmacta IN (" & Mid(SQL, 2) & ")"
    End If

    
    
    'JULIO 2013
    'Si ha especificado un IVA, comprobaremos SI es el REA,
    'Si es el REA buscare las retenciones
    If TipoIVA >= 0 Then
        'Si es 3=REA
        SQL = DevuelveDesdeBD("tipodiva", "tiposiva", "codigiva", CStr(TipoIVA), "N")
        If SQL = "3" Then
            'OK ES REA
            M1 = InStr(1, UCase(d), " FROM ")
            'NO puede ser 0
            d = Mid(d, M1)
            SQL = "(SELECT numregis,anofacpr,cuentas.codmacta " & d & ")"
            
            d = "select * from linfactprov where (numregis,anofacpr,codtbase) IN " & SQL
            RT.Open d, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                '0003200021
                d = "'" & Format(RT!NumRegis, "0000000000") & "'"
                d = " where codusu =" & vUsu.Codigo & " and numfac=" & d & " and cta='" & RT!codtbase & "'"
                d = "update Usuarios.ztmpfaclinprov set retencion=" & TransformaComasPuntos(CStr(DBLet(RT!impbaspr, "N"))) & d
                Conn.Execute d
            
                RT.MoveNext
            Wend
            RT.Close
        End If
    End If
    
    
    
    
    'Marzo 2015
    ' Sujeto pasivo.
    '   Pasara dos veces si es que hay facuras de sujeto pasivo
    ' Par saber que hay facturas de sujeto pasvio
    
    'Febrero 2015
    
    Dim BuscaISP As Boolean
    Dim BuscaComunitarias As Boolean
    
    
        SQL = UCase(d)
        M2 = InStr(1, SQL, " FROM ")
        SQL = Mid(SQL, M2)
        'quito el order by
        M2 = InStr(1, SQL, " ORDER BY")
        SQL = Mid(SQL, 1, M2)
        
        
        d = "Select count(*)  " & SQL & " AND extranje=3 "  'SUJETO PASIVO
        RT.Open d, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        M2 = 0
        If Not RT.EOF Then M2 = DBLet(RT.Fields(0), "N")
        RT.Close
        BuscaISP = M2 > 0
        
        d = "Select count(*)  " & SQL & " AND extranje=1 "  'Comunitarias
        RT.Open d, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        M2 = 0
        If Not RT.EOF Then M2 = DBLet(RT.Fields(0), "N")
        RT.Close
        BuscaComunitarias = M2 > 0
        
        
        M2 = 1
        If BuscaISP Then M2 = M2 + 1
        If BuscaComunitarias Then M2 = M2 + 1
        
        
        'Y para que no haga los selects con "tantos" tipos de iva, veremos cuales son los que se estan utilizando
        RS.Close
        
        d = "select distinct tp1facpr,tp2facpr,tp3facpr " & SQL
        
        RS.Open d, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        d = "|"
        While Not RS.EOF
            For M1 = 0 To 2
                If Not IsNull(RS.Fields(M1)) Then
                    H = "|" & RS.Fields(M1) & "|"
                    If InStr(d, H) = 0 Then d = d & RS.Fields(M1) & "|"
                End If
            Next M1
            RS.MoveNext
        Wend
        RS.Close
        
        d = Mid(d, 1, Len(d) - 1)  'quito el ultimo pipe
        d = Mid(d, 2)
        d = Replace(d, "|", " , ")
        If d = "" Then
            d = "1=1"
        Else
            d = " codigiva IN (" & d & ")"
        End If
        
        d = "Select * from tiposiva where " & d
        RS.Open d, Conn, adOpenKeyset, adCmdText
        
        
        
        
            
        
    
    
    M1 = 1
    
    For M3 = 1 To M2   'Sera una o dos veces si hay sujeto pasivo o no
        RS.MoveFirst
        
        If M3 = 1 Then
            H = ",1"   'Indicara IVA corriente
            d = " AND not extranje in (1, 3)"   '3:=sujeto pasivo   1:=Intracomunitarias
        Else
            
            
            If M3 = 2 Then
                'Estamos buscando ISP o COM
                If BuscaComunitarias Then     '3:=sujeto pasivo   1:=Intracomunitarias
                    H = ",0"
                    d = "3"
                Else
                    H = ",2"
                    d = "1"
                End If
                d = " AND extranje =  " & d
            Else
                'Si llega al tercero entonces es INtrcoms ya que
                'BuscaComunitarias
                H = ",2"
                d = " AND extranje = 1 "   '3:=sujeto pasivo   1:=Intracomunitarias
            End If
        End If
        
        
        While Not RS.EOF
            ImpD = 0
            ImCierrD = 0
            'Suma en Imp1
            SQL = " select sum(ti1facpr),sum(ba1facpr)  from cabfactprov WHERE  " & vSQL & "  AND tp1facpr=" & RS!codigiva
            SQL = SQL & d   'sujeto pasivo .. ver arriba
            RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            ImpH = 0
            ImCierrH = 0
            If Not RT.EOF Then
                If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
                If Not IsNull(RT.Fields(1)) Then ImCierrH = RT.Fields(1)
            End If
            RT.Close
            ImpD = ImpD + ImpH
            ImCierrD = ImCierrD + ImCierrH
            
            'Suma en Imp2
            SQL = " select sum(ti2facpr),sum(ba2facpr)  from cabfactprov WHERE  " & vSQL & "  AND tp2facpr=" & RS!codigiva
            SQL = SQL & d   'sujeto pasivo .. ver arriba
            RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            ImpH = 0
            ImCierrH = 0
            If Not RT.EOF Then
                If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
                If Not IsNull(RT.Fields(1)) Then ImCierrH = RT.Fields(1)
            End If
            ImpD = ImpD + ImpH
            ImCierrD = ImCierrD + ImCierrH
            RT.Close
            'Suma en Imp3
            SQL = " select sum(ti3facpr),sum(ba3facpr)  from cabfactprov WHERE  " & vSQL & "  AND tp3facpr=" & RS!codigiva
            SQL = SQL & d   'sujeto pasivo .. ver arriba
            RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            ImpH = 0
            ImCierrH = 0
            If Not RT.EOF Then
                If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
                If Not IsNull(RT.Fields(1)) Then ImCierrH = RT.Fields(1)
            End If
            ImpD = ImpD + ImpH
            ImCierrD = ImCierrD + ImCierrH
            RT.Close
            
            'Si el importe es <>0 entonces lo añadimos a tmpresumn
            If ImpD <> 0 Or ImCierrD <> 0 Then
                SQL = "INSERT INTO Usuarios.ztmpresumenivafac (codusu, orden, IVA, TotalIVA,Sumabases,TipoIva,IvaCorriente,NombreIVA) VALUES (" & vUsu.Codigo & "," & M1
                SQL = SQL & ",'" & Format(RS!porceiva, "0.00")
                'MAYO 2005
                If RS!tipodiva = 4 Then SQL = SQL & "ND"
                SQL = SQL & "'," & TransformaComasPuntos(CStr(ImpD))
                SQL = SQL & "," & TransformaComasPuntos(CStr(ImCierrD)) & "," & RS!codigiva & H & ",'" & DevNombreSQL(RS!nombriva) & "')"
                Conn.Execute SQL
                M1 = M1 + 1
            End If
            
            
            
            'Por si lleva recargo de equivalencia
            If RS!porcerec > 0 Then
                ImpD = 0
                SQL = " select sum(tr1facpr)  from cabfactprov WHERE  " & vSQL & "  AND tp1facpr=" & RS!codigiva
                SQL = SQL & d   'sujeto pasivo .. ver arriba
                RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                ImpH = 0
                If Not RT.EOF Then
                    If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
                End If
                RT.Close
                ImpD = ImpD + ImpH
                    
                SQL = " select sum(tr2facpr)  from cabfactprov WHERE  " & vSQL & "  AND tp2facpr=" & RS!codigiva
                SQL = SQL & d   'sujeto pasivo .. ver arriba
                RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                ImpH = 0
                If Not RT.EOF Then
                    If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
                End If
                RT.Close
                ImpD = ImpD + ImpH
                    
            
                SQL = " select sum(tr3facpr)  from cabfactprov WHERE  " & vSQL & "  AND tp3facpr=" & RS!codigiva
                SQL = SQL & d   'sujeto pasivo .. ver arriba
                RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                ImpH = 0
                If Not RT.EOF Then
                    If Not IsNull(RT.Fields(0)) Then ImpH = RT.Fields(0)
                End If
                RT.Close
                ImpD = ImpD + ImpH
                    
                    
                    
                    
                If ImpD <> 0 Then
                    SQL = "INSERT INTO Usuarios.ztmpresumenivafac (codusu, orden, IVA, TotalIVA,Sumabases,TipoIva,IvaCorriente,NombreIVA) VALUES (" & vUsu.Codigo & "," & M1
                    SQL = SQL & ",'re" & Format(RS!porcerec, "0.00")
                    If RS!tipodiva = 4 Then SQL = SQL & "ND"
                    SQL = SQL & "'," & TransformaComasPuntos(CStr(ImpD))
                    SQL = SQL & ",0," & RS!codigiva & H & ",'" & DevNombreSQL(RS!nombriva) & "')"
                    Conn.Execute SQL
                    M1 = M1 + 1
                End If
                        
            End If  'Del recargo de equivalencia
            
            
            
            
            
            
            'Sig
            RS.MoveNext
        Wend
    
    Next  'De dos facses si hay (o no) sujeto pasivo
    
    
    RS.Close
    Set RS = Nothing
    Set RT = Nothing
    
    'Insertaremos la cabecera, con los importes
    ListadoFacturasProveedores = True



End Function













Private Function DevuelveIVA_FraPro(ByRef RD As ADODB.Recordset, Tipo As String, EsFacturaNoDeducible As Boolean, TipoFaPro As Integer) As String
    'Buscamos el tipo de iva
    RD.Find "codigiva = " & Tipo, , adSearchForward, 1
    
    '
    If EsFacturaNoDeducible Then
        DevuelveIVA_FraPro = "FN"
    Else
        If TipoFaPro = 3 Then
            DevuelveIVA_FraPro = "ISP"
        ElseIf TipoFaPro = 1 Then
            DevuelveIVA_FraPro = "COM"
        Else
            DevuelveIVA_FraPro = ""
    
            If Not RD.EOF Then
                If RD!tipodiva = 4 Then
                    DevuelveIVA_FraPro = "ND"
                ElseIf RD!tipodiva = 5 Then
                    DevuelveIVA_FraPro = "IM"
                End If
            End If
        End If
    End If
    
End Function







'------------------------------------------------------------------------

Public Function ImporteBalancePresupuestario(ByRef vSQL As String) As Currency

ImPerH = 0
Set RT = New ADODB.Recordset
RT.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RT.EOF Then
    If Not IsNull(RT.Fields(0)) Then
        ImpD = RT.Fields(0)
    Else
        ImpD = 0
    End If
    If Not IsNull(RT.Fields(1)) Then
        ImpH = RT.Fields(1)
    Else
        ImpH = 0
    End If
    ImPerH = ImpD - ImpH
End If
RT.Close
ImporteBalancePresupuestario = ImPerH
End Function


'Desde donde:
'       0.- Listado simulacin
'       1.- Venta / baja de elmento
Public Function HazSimulacion(ByRef vSQL As String, Fecha As Date, DesdeDonde As Byte) As Boolean
Dim FechaCalculoVentaBaja As Date
Dim I2 As Integer
On Error GoTo EHazSimulacion
    
    HazSimulacion = False
    'Obtenemos la ultmia fecha de amortizacion
    Set RT = New ADODB.Recordset
    RT.Open "Select * from samort where codigo =1", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        MsgBox "Error leyendo parámetros.", vbExclamation
        RT.Close
        Exit Function
    End If
    
    'Ademas, en M2 pondremos el tipo de amortizacion, el valor por el k habra que dividir + adelante
    Select Case Val(RT!tipoamor)
    Case 2
        'Semestral
        M2 = 2
        I2 = 6 'dato auxiliar
    Case 3
        'Trimestral
        M2 = 4
        I2 = 4
    Case 4
        'mensual
        M2 = 12
        I2 = 1
    Case Else
        'Anual
        M2 = 1
        I2 = 12
    End Select
    
    
    If RT.Fields(3) > CDate(Fecha) Then
        MsgBox "Fecha ultima amortizacion mayor que fecha operacion.", vbExclamation
        RT.Close
        Exit Function
    End If
    'En m1 almacenamos los dias del diferencia
    vFecha1 = CDate(Fecha)  'La nueva fechamo
    vFecha2 = RT.Fields(3)  'Ultmfechaamort
    
    
    
    If DesdeDonde = 1 Then
        FechaCalculoVentaBaja = DateAdd("m", I2, vFecha2)
    End If
    
    
    RT.Close
    
    
    'Borramos temporales
    'Antes
    'SQL = "Delete from tmpsimula where codusu=" & vUsu.Codigo
    'Ahora
    SQL = "Delete from Usuarios.zsimulainm where codusu=" & vUsu.Codigo
    Conn.Execute SQL
    
    
    
    'Obtenemos el recordset
    SQL = "select codinmov,fechaadq,valoradq,anovidas,amortacu,sinmov.conconam,"
    SQL = SQL & " tipoamor,coeficie,valorres,fecventa,coefimaxi,nominmov,nomconam from sconam,sinmov where sinmov.conconam=sconam.codconam"
    SQL = SQL & " AND fecventa is null AND impventa is null AND situacio<>4"
    'Junio 2005
    '-------------
    ' Indicamos que la fecha adq sea menor que la fecha simulacion
    SQL = SQL & " AND fechaadq<='" & Format(vFecha1, FormatoFecha) & "'"
    If vSQL <> "" Then SQL = SQL & " AND " & vSQL
    
    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        If DesdeDonde = 0 Then
            MsgBox "Ningún  elemento de inmovilizado con estos valores.", vbExclamation
        Else
            'Estamos vendiendo un elto, o de baja
            'Signifaca que NO hay que amortizar
            HazSimulacion = True
        End If
        RT.Close
        Exit Function
    End If

    
    'Antes
    'SQL = "INSERT INTO tmpsimula (codusu, codinmov, codconam, totalamor) VALUES (" & vUsu.Codigo & ","
    'Ahora
    SQL = "INSERT INTO Usuarios.zsimulainm (codusu, codigo, conconam, nomconam, codinmov,"
    SQL = SQL & "nominmov, fechaadq, valoradq, amortacu, totalamor) VALUES (" & vUsu.Codigo & ","
    M1 = 1
    
    'Dias totales
    If DesdeDonde = 1 Then
        A1 = DateDiff("d", vFecha2, FechaCalculoVentaBaja)
    Else
        A1 = DateDiff("d", vFecha2, vFecha1)
    End If
    If A1 <= 0 Then
        MsgBox "Diferencia entre fechas amortización es <=0", vbExclamation
        Exit Function
    End If
    
    While Not RT.EOF
        
        ObtenAmortizacionAnualSimulacion 'En IMPERD esta almacenada
        'Vemos los dias del period a aplicar el valor
        CalcularDiasAplicablesSimulacion DesdeDonde = 1, Fecha
    
        If ImPerD > 0 Then
                    
            'Vemos, en funcion de los dias
           
            ImCierrD = Round(ImPerD * (A2 / A1), 2)
            'ImPerD = Round(ImPerD, 2)
            'ImCierrD = ImPerD * ImCierrD
        
            'Calcualmos los valores
            'ImPerH = Round(ImPerD * A2 / A1, 2)
            ImPerH = ImCierrD
            
            'Ahora, si lo k ahy k amortizar es mayor de lo que queda, entonces amortizamos
            'solo lo k queda
            ImPerD = Round(RT!valoradq - RT!amortacu, 2)
            If ImPerH > ImPerD Then ImPerH = ImPerD

            d = TransformaComasPuntos(CStr(ImPerH))
            'Insertamos
            Aux = SQL & M1 & "," & RT!conconam & ",'" & RT!nomconam & "'," & RT!Codinmov
            Aux = Aux & ",'" & DevNombreSQL(RT!nominmov) & "','" & Format(RT!fechaadq, "dd/mm/yyyy") & "',"
            H = TransformaComasPuntos(CStr(RT!valoradq))
            Aux = Aux & H & ","
            H = TransformaComasPuntos(CStr(RT!amortacu))
            Aux = Aux & H & "," & d & ")"
            Conn.Execute Aux
        End If
        'Siguiente elemento
        RT.MoveNext
        M1 = M1 + 1
    Wend
    HazSimulacion = True
    Exit Function
EHazSimulacion:
    MuestraError Err.Number, Err.Description
    Set RT = Nothing
End Function

'////////////////////////////////////
' A partir del RT , recordset k tiene los datos,
' pondremos en ImAcD el valor de la amortizacion anual
' en el segundo paso pondremos la apliacble( en funcion de los dias transcurridos
Private Sub ObtenAmortizacionAnual()
    Select Case RT!tipoamor
    Case 2
        'Lineal
        ImPerD = (RT!valoradq - RT!valorres) / RT!anovidas
    Case 3
        'Degresiva
        ImPerD = (RT!valoradq - RT!amortacu) * (RT!coefimaxi / 100)
    Case 4
        'Porcentual
        ImPerD = (RT!valoradq * RT!coeficie) / 100
    Case Else
        'Tablas
        ImPerD = RT!valoradq / RT!anovidas
    End Select
    ImPerD = Round(ImPerD, 2)  'Redondeando
    'Aplicamos al period mensual, trimestr...
    ImPerD = Round(ImPerD / M2, 2)
End Sub



'///////////////////////////////////////
'
' Esto es, si son 60 dias pero solo hay que aplicar 20 entonces
Private Sub CalcularDiasAplicables2()
    
    
    If RT!fechaadq > vFecha2 Then
        If RT!fechaadq > vFecha1 Then
            'Ha comprado incluso despues de
            'la fecha de amortizacion
            A2 = 0
        Else
            'Modificado 15 Octubre 2008 al igual que calculardiasaplicablessimulacion
            A2 = DateDiff("d", RT!fechaadq, vFecha1) + 1
            If A2 > A1 Then A2 = A1
        End If
    Else
        A2 = A1
    End If
    If Not IsNull(RT!fecventa) And RT!fecventa > vFecha2 Then
        'Se ha vendido despues del la ultima amortizacion
        A2 = DateDiff("d", vFecha2, RT!fecventa)
    End If
    If Not IsNull(RT!fecventa) And RT!fecventa < vFecha2 Then
        'Se ha vendido despues del la ultima amortizacion
        A2 = 0
    End If
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'      SIMULACION
'----------------------------------------------------------------
'----------------------------------------------------------------
'nuevo a 29 Marzo 2005
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub ObtenAmortizacionAnualSimulacion()
    Select Case RT!tipoamor
    Case 2
        'Lineal
        ImPerD = (RT!valoradq - RT!valorres) / RT!anovidas
    Case 3
        'Degresiva
        ImPerD = (RT!valoradq - RT!amortacu) * (RT!coefimaxi / 100)
    Case 4
        'Porcentual
        ImPerD = (RT!valoradq * RT!coeficie) / 100
    Case Else
        'Tablas
        ImPerD = RT!valoradq / RT!anovidas
    End Select
    
    'En imperd tengo la simulacion ANUAL
    
    ImPerD = Round(ImPerD, 2)  'Redondeando
    'Aplicamos al period mensual, trimestr...
    ImPerD = Round(ImPerD / M2, 2)
End Sub



Private Sub CalcularDiasAplicablesSimulacion(EsEnVentabaja As Boolean, FechaVtaBaja As Date)
    
    
    If RT!fechaadq > vFecha2 Then
        If RT!fechaadq > vFecha1 Then
            'Ha comprado incluso despues de
            'la fecha de amortizacion
            A2 = 0
        Else
            'Se ha comprado despues de la ultima amortizacion
            'Nuevo 15 Octubre 2008
            'A2 = DateDiff("d", RT!fechaadq, vFecha1)
            A2 = DateDiff("d", RT!fechaadq, vFecha1) + 1  'Ya que el primer dia tb se utiliza
            If A2 > A1 Then A2 = A1 'Esto no deberia pasar NUNCA, pero mas vale prevenir que curar
        End If
    Else
        A2 = A1
    End If
    
    If EsEnVentabaja Then
        'Aqui veremos cuantos dias hay que aplicar la amortizacion
        A2 = DateDiff("d", vFecha2, FechaVtaBaja)
    Else
        
        If Not IsNull(RT!fecventa) Then
            If RT!fecventa >= vFecha2 Then
                'Se ha vendido despues del la ultima amortizacion
                A2 = DateDiff("d", vFecha2, RT!fecventa)
            Else
                A2 = 0
            End If
        End If
    End If
End Sub



'///////////////////////////////////////////////////////////////////////////
'
'   CALCULO AMORTIZACION
'
Public Function CalculaAmortizacion(Codinmov As Long, Fecha As Date, DivMes As Integer, UltimaAmort As Date, ParametrosContabiliza As String, mContador As Long, ByRef NumLinea As Integer, EsVentaBaja As Boolean) As Boolean
Dim RS As Recordset

On Error GoTo ECalculaAmortizacion
    CalculaAmortizacion = False

    
    
    Set RT = New ADODB.Recordset
    Aux = "select sinmov.*,sconam.coefimaxi from sconam,sinmov where sinmov.conconam=sconam.codconam"
    Aux = Aux & " AND codinmov = " & Codinmov
    RT.Open Aux, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    vFecha1 = Fecha
    vFecha2 = UltimaAmort
    M2 = DivMes
  
    ObtenAmortizacionAnual 'En IMPERD esta almacenada
    'Vemos los dias del period a aplicar el valor
    If EsVentaBaja Then
        'En i metermos los meses a  sumar a la fecha
        If DivMes = 1 Then
            'Amortizacion anual
            A1 = 12 'le suma
        ElseIf DivMes = 12 Then
            'MENSUAL
            A1 = 1
        ElseIf DivMes = 4 Then
            'TRIMESTRAL
            A1 = 3
        Else
            'Semestral
            A1 = 6
        End If
        'EL sumamos los I meses a la ultima fecha de amortizacion
        A1 = DateDiff("d", vFecha2, DateAdd("m", A1, vFecha2))
    Else
        A1 = DateDiff("d", vFecha2, vFecha1)
    End If
    CalcularDiasAplicables2
    
    
    'Ya tenesmo en A1 los dias totales y en a2 los aplicables
    If A1 = 0 Then ImPerD = 0
    ImPerH = 0
    If ImPerD > 0 Then
        'Calcualmos los valores
        ImPerH = Round(ImPerD * (A2 / A1), 2)
            
        'Ahora, si lo k ahy k amortizar es mayor de lo que queda, entonces amortizamos
        'solo lo k queda
        ImPerD = Round(RT!valoradq - RT!amortacu, 2)

        If ImPerH > ImPerD Then ImPerH = ImPerD
            
        
        'Calculo el % de amortizacion
        ImpD = Round((ImPerH / RT!valoradq) * 100, 2)
        
        
        'Metemos en hco inmovilizado
        '--------------------------
        SQL = "INSERT INTO shisin (codinmov, fechainm, imporinm, porcinm) VALUES ("
        SQL = SQL & Codinmov & ",'" & Format(vFecha1, FormatoFecha) & "',"
        H = TransformaComasPuntos(CStr(ImPerH))
        d = TransformaComasPuntos(CStr(ImpD))
        SQL = SQL & H & ","
        SQL = SQL & d & ")"
        Conn.Execute SQL

        'ParametrosContabiliza :=>  contabiliza|debe|haber|diario
        If RecuperaValor(ParametrosContabiliza, 1) = "1" Then
            'Contabilizamos insertando en diario de apuntes
                'Insertamos las lineas
                'Este trozo es comun para las del debe y las del haber
                SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce,"
                SQL = SQL & "timporteD, timporteH, codccost, ctacontr, idcontab, punteada) VALUES ("
                SQL = SQL & RecuperaValor(ParametrosContabiliza, 4) & ",'"
                SQL = SQL & Format(vFecha1, FormatoFecha)
                SQL = SQL & "'," & mContador & ","
                
                
                'amortizacion acumulada -->Haber
                Aux = NumLinea & ",'" & RT!codmact3 & "','" & Format(Codinmov, "000000") & "',"
                Aux = Aux & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
                Aux = Aux & ",'" & DevNombreSQL(RT!nominmov)
                Aux = Aux & "',NULL," & H    'H tiene el importe del inmovilizado
                'El Centro de coste es 0
                Aux = Aux & ",NULL"
                If RT!repartos = 0 Then
                    vCta = "'" & RT!codmact2 & "'"
                Else
                    vCta = "NULL"
                End If
                Aux = Aux & "," & vCta & ",'CONTAI',0)"
                Conn.Execute SQL & Aux
                NumLinea = NumLinea + 1
             
                'Cta gastos --> Debe
                If RT!repartos = 0 Then
                    
                    Aux = NumLinea & ",'" & RT!codmact2 & "','" & Format(Codinmov, "000000") & "',"
                    Aux = Aux & RecuperaValor(ParametrosContabiliza, 3)   'Concepto HABER
                    Aux = Aux & ",'" & DevNombreSQL(RT!nominmov) & "'," & H & ",NULL"      'H tiene el importe del inmovilizado
                    If vParam.autocoste Then
                        If IsNull(RT!codccost) Then
                            Aux = Aux & ",NULL"
                        Else
                            Aux = Aux & ",'" & RT!codccost & "'"
                        End If
                    Else
                        'No lleva centro de coste
                        Aux = Aux & ",NULL"
                    End If
                    Aux = Aux & ",'" & RT!codmact3 & "','CONTAI',0)"
                    Conn.Execute SQL & Aux
                    
                Else
                    'Si k tiene reparto
                    Set RS = New ADODB.Recordset
                    RS.Open "Select * from sbasin where codinmov =" & Codinmov, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    ImAcD = 0 'Tendre el sumatorio de los repartos
                    While Not RS.EOF
                        'Calculamos el importe por centual y lo metemos en IMACH
                        ImAcH = Round(((ImPerH * RS!porcenta) / 100), 2)
                        If vParam.autocoste Then
                            If IsNull(RS!codccost) Then
                                vCta = "NULL"
                            Else
                                vCta = "'" & RS!codccost & "'"
                            End If
                        Else
                            vCta = "NULL"
                        End If
                        Aux = NumLinea & ",'" & RS!codmacta2 & "','" & Format(Codinmov, "000000") & "',"
                        Aux = Aux & RecuperaValor(ParametrosContabiliza, 3)   'Concepto HABER
                        Aux = Aux & ",'" & DevNombreSQL(RT!nominmov)
                        
                        'Avanzamos al siguiente
                        RS.MoveNext
                        If RS.EOF Then
                            'Es la ultima linea. Compruebo k el sumatorio de las lineas sea le total
                            ImAcH = ImPerH - ImAcD
                        Else
                            ImAcD = ImAcD + ImAcH
                            NumLinea = NumLinea + 1
                        End If
                        H = TransformaComasPuntos(CStr(ImAcH))
                        'Aux = Aux & "',NULL," & H & "," 'H tiene el importe del inmovilizado
                        Aux = Aux & "'," & H & ",NULL," 'H tiene el importe del inmovilizado
                        Aux = Aux & vCta   'ccoste
                        Aux = Aux & ",'" & RT!codmact3 & "','CONTAI',0)"
                        Conn.Execute SQL & Aux
                    Wend
                    RS.Close
                    Set RS = Nothing
                End If
        End If
    End If
    
    
    
    'ACtualizamos eltos. inmovilizado
    'En imperh tengo lo k voy a amortizar
    'En imperd tengo la nueva amortizacon acumulada
    ImPerD = RT!amortacu + ImPerH
    H = TransformaComasPuntos(CStr(ImPerD))
    SQL = "UPDATE sinmov set amortacu=" & H
    If ImPerD = RT!valoradq Then
        'Totalmente amortizado
        SQL = SQL & ", situacio= 4"
    End If
    SQL = SQL & " WHERE codinmov=" & Codinmov
    Conn.Execute SQL
    
    CalculaAmortizacion = True
    Exit Function
ECalculaAmortizacion:
    MuestraError Err.Number, "Calcula Amortizacion" & vbCrLf & Err.Description
    Set RT = Nothing
End Function


Public Function ObtenerparametrosAmortizacion(ByRef DivMes As Integer, ByRef UltmAmort As Date, ByRef RestoParametros As String) As Boolean

    Set RT = New ADODB.Recordset
    RT.Open "Select * from samort where codigo =1", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        MsgBox "Error leyendo parámetros.", vbExclamation
        RT.Close
        ObtenerparametrosAmortizacion = False
        Exit Function
    End If
    
    'Ademas, en M2 pondremos el tipo de amortizacion, el valor por el k habra que dividir + adelante
    M1 = RT!tipoamor
    Select Case M1
    Case 2
        'Semestral
        M2 = 2
    Case 3
        'Trimestral
        M2 = 4
    Case 4
        'mensual
        M2 = 12
    Case Else
        'Anual
        M2 = 1
    End Select
    DivMes = M2
    UltmAmort = RT.Fields(3)  'Ultmfechaamort
    RestoParametros = RT!intcont & "|"
    If (RT!intcont = 1) Then RestoParametros = RestoParametros & RT!condebes & "|" & RT!conhaber & "|" & RT!NumDiari & "|"
    RestoParametros = RestoParametros & RT!Preimpreso & "|"
    RT.Close
    ObtenerparametrosAmortizacion = True
    Set RT = Nothing
End Function




Public Function LiquidacionIVA(Periodo As Byte, Anyo As Integer, Empresa As Integer, Detallado As Boolean) As Boolean
Dim RIVA As Recordset
Dim TieneDeducibles As Boolean    'Para ahorrar tiempo
Dim HayRecargoEquivalencia As Boolean  'Para ahorrar tiempo tb
'Dim BienDeInversion As String
Dim IvaInversionSujetoPasivo  As String
Dim J As Integer

    '       cliprov     0- Facturas clientes
    '                   1- RECARGO EQUIVALENCIA
    '                   2- Facturas proveedores
    '                   3- libre
    '                   4- IVAS no deducible
    '                   5- Facturas NO DEDUCIBLES
    '                   6- FRAPRO con Bien inversion
    '                   7- Compras extranjero     'Marzo 2014
    '                   8- Inversion sujeto pasivo (Abril 2015)
    If vParam.periodos = 1 Then
        'Esamos en mensual
        If Periodo > 12 Then
            MsgBox "Error en el periodo a tratar.", vbExclamation
            Exit Function
        End If
        vFecha1 = CDate("01/" & Periodo & "/" & Anyo)
        M1 = DiasMes(Periodo, Anyo)
        vFecha2 = CDate(M1 & "/" & Periodo & "/" & Anyo)
        
    Else
        'IVA TRIMESTRAL
        If Periodo > 4 Then
            MsgBox "Error en el periodo a tratar.", vbExclamation
            Exit Function
        End If
        M2 = ((Periodo - 1) * 3) + 1
        vFecha1 = CDate("01/" & M2 & "/" & Anyo)
        M2 = ((Periodo - 1) * 3) + 3
        M1 = DiasMes(CByte(M2), Anyo)
        vFecha2 = CDate(M1 & "/" & M2 & "/" & Anyo)
    End If
    
    
    
    vCta = "conta" & Empresa
    
    'Para la cadena de busqueda
    LiquidacionIVA = False
    

    
    '-----------------------------------------------
    '-----------------------------------------------
    '-----------------------------------------------
    'CLIENTES
    '-----------------------------------------------
    CargarIvasATratar True
    
    'Abrimos los ivas a tratar
    Set RIVA = New ADODB.Recordset
    SQL = "Select * from tmpliqiva where codusu = " & vUsu.Codigo
    RIVA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RIVA.EOF
        If RIVA.Fields(1) >= 0 Then
            'TotalIva
            ImpD = 0: ImpH = 0
            
            'Calulamos el IVA
            TotalIva CStr(RIVA.Fields(1)), 0, False     'ANTES TotalIva CStr(RIVA.Fields(1)), True
            
            'Insertamos en la USUARIOS.TMPLIQUIVA los valores
            SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
            d = TransformaComasPuntos(CStr(RIVA.Fields(1)))
            SQL = SQL & d
            d = TransformaComasPuntos(CStr(ImpD))
            H = TransformaComasPuntos(CStr(ImpH))
            SQL = SQL & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",0)"    'El 0 es cliente normal
            Conn.Execute SQL
        End If
        
        'Siguiente
        RIVA.MoveNext
    Wend
    RIVA.Close
    
    If Detallado Then
        'Insertamos en la tabla tmpimpbalan los valores
        GeneraIVADetallado 0
    End If
    
    
   
    '-----------------------------------------------
    '-----------------------------------------------
    '-----------------------------------------------
    '           PROVEEDORES
    '-----------------------------------------------
    CargarIvasATratar False
    
    Set RIVA = New ADODB.Recordset
    
    'MODIFIACION 31/ENERO / 2005
    'Veremos si tiene DEDUCIBLES
    SQL = " SELECT count(*) from " & vCta & ".cabfactprov WHERE fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " AND nodeducible =1 "
    RIVA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TieneDeducibles = False
    If Not RIVA.EOF Then
        If Not IsNull(RIVA.Fields(0)) Then
            If RIVA.Fields(0) > 0 Then TieneDeducibles = True
        End If
    End If
    RIVA.Close
        
    
    
    'Veremos si hay Inversion de sujeto pasivo y que tipo de IVA lo llevan
    SQL = " SELECT tp1facpr,tp2facpr,tp3facpr from " & vCta & ".cabfactprov WHERE fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    
    
    SQL = " SELECT pi1facpr,pi2facpr,pi3facpr from " & vCta & ".cabfactprov WHERE fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    SQL = SQL & " AND nodeducible =0 and extranje=3 GROUP BY 1,2,3"
    RIVA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    IvaInversionSujetoPasivo = "|"
    While Not RIVA.EOF
        For J = 0 To 2
            SQL = ""
            If Not IsNull(RIVA.Fields(J)) Then
                SQL = Format(RIVA.Fields(J), "0.00")
                
                d = "|" & SQL & "|"
                If InStr(1, IvaInversionSujetoPasivo, d) = 0 Then IvaInversionSujetoPasivo = IvaInversionSujetoPasivo & SQL & "|"
            End If
        Next
        RIVA.MoveNext
    Wend
    RIVA.Close
    If IvaInversionSujetoPasivo = "|" Then IvaInversionSujetoPasivo = ""   'NO hay ISP
    
    
    
    
    SQL = "Select * from tmpliqiva where codusu = " & vUsu.Codigo
    RIVA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RIVA.EOF
        If RIVA.Fields(1) >= 0 Then
        
            
        
            'TotalIva
            ImpD = 0: ImpH = 0
                        
            
            'Calulamos el IVA, sin el deducible
            TotalIva CStr(RIVA.Fields(1)), 1, False
            
            'ImPerD = 0: ImPerH = 0  lo hara en la funcion
            IvaDeducibleBienInversion CStr(RIVA.Fields(1))
            If ImPerD > ImpD And ImPerD <> 0 Then
                'Hay mas IVA en bien de inversion que el total del IVA. Algo esta mal
                MsgBox "Mas iva en bien inversion que el total iva", vbExclamation
            End If
            ImpD = ImpD - ImPerD
            ImpH = ImpH - ImPerH
            
            
            
            
            'EXTRANJERO
            'ImCierrD = 0: ImCierrH = 0  lo hara en la funcion
            IvaDeducibleExtranjero CStr(RIVA.Fields(1))
            If ImCierrD > ImpD And ImCierrD <> 0 Then
                'Hay mas IVA en bien de inversion que el total del IVA. Algo esta mal
                'MsgBox "Mas iva en extranjero que el total iva", vbExclamation
            End If
            ImpD = ImpD - ImCierrD
            ImpH = ImpH - ImCierrH
            
            
            'Abril 2015
            'Inversion de sujeto pasivo
            ImAcD = 0
            ImAcH = 0
            If IvaInversionSujetoPasivo <> "" Then
                SQL = Format(RIVA.Fields(1), "0.00")
                SQL = "|" & SQL & "|"
                If InStr(1, IvaInversionSujetoPasivo, SQL) > 0 Then
                    'Este IVA es el porcentaje del inversion del sujeto pasivo
                    'Realmente NO me lo resta de este, si no que creara una nueva entrada para que RESTE
                    IvaDelaInversionSujetoPasivo CStr(RIVA.Fields(1))
                    'Le sumamos el IVA al total deducible de ese porcentaje
                    ImpD = ImpD + ImAcD
                    ImpH = ImpH + ImAcH
                End If
            End If
            
            
            'Insertamos en la USUARIOS.TMPLIQUIVA los valores
            SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
            d = TransformaComasPuntos(CStr(ImpD))
            H = TransformaComasPuntos(CStr(ImpH))
            SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",2)"   '2.- IVa prov
            Conn.Execute SQL
            
            ''Si tiene Bien inversion
            If ImPerD > 0 Then
                SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                d = TransformaComasPuntos(CStr(ImPerD))
                H = TransformaComasPuntos(CStr(ImPerH))
                SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",6)"   '6.-PRO bien inversion
                Conn.Execute SQL
            End If
            
            'Si tiene IVA extranjero
            If ImCierrD > 0 Then
                SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                d = TransformaComasPuntos(CStr(ImCierrD))
                H = TransformaComasPuntos(CStr(ImCierrH))
                SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",7)"   '7.-EXTRANJERO
                Conn.Execute SQL
            End If
            
            'Si tiene IVA inversion de sujeto pasivo
            If ImAcD > 0 Then
                SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                d = TransformaComasPuntos(CStr(ImAcD))
                H = TransformaComasPuntos(CStr(ImAcH))
                SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",8)"   '8.-INVERSION DE SUJETO PASIVO
                Conn.Execute SQL
            End If
            
            
            
            If TieneDeducibles Then
                'Modificacion del 31 Enero 2005.  Facutras con IVA NO deducible
                'Es para mostrar los que tienen la marca de NO deducible
                ImpD = 0: ImpH = 0
                
                'Calulamos el IVA, sin el deducible
                TotalIva CStr(RIVA.Fields(1)), 2, False
                
                'Insertamos en la USUARIOS.TMPLIQUIVA los valores
                SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                d = TransformaComasPuntos(CStr(ImpD))
                H = TransformaComasPuntos(CStr(ImpH))
                SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",4)"   'El 4 significa DEDUCIBLE
                Conn.Execute SQL
            End If
            
            
            
            
            
            
            
            
        End If
        
        'Siguiente
        RIVA.MoveNext
    Wend
    RIVA.Close
    
    
    
    
    
    If Detallado Then
        'Insertamos en la tabla tmpimpbalan los valores
        GeneraIVADetallado 2
        
        SQL = DevuelveDesdeBD("count(*)", "Usuarios.zliquidaiva", "cliente=7 AND codusu", CStr(vUsu.Codigo))
        If SQL = "" Then SQL = "0"
        If Val(SQL) > 0 Then GeneraIVADetallado 7 'IMPORTACION
        
        If TieneDeducibles Then GeneraIVADetallado 4
    End If
    
        
    'BUSCAMOS, SI TIENE, tipos de IVA no deducible
    SQL = "Select * from tmpliqiva where codusu = " & vUsu.Codigo
    RIVA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RIVA.EOF
        If RIVA.Fields(1) >= 0 Then
            ImpD = 0: ImpH = 0
            
            'Calulamos el IVA, sin el deducible
            TotalIva CStr(RIVA.Fields(1)), 2, True
            
            'Insertamos en la USUARIOS.TMPLIQUIVA los valores
            If ImpD <> 0 And ImpH <> 0 Then
                SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                d = TransformaComasPuntos(CStr(ImpD))
                H = TransformaComasPuntos(CStr(ImpH))
                SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",5)"   'El 5 significa TIPO IVA NO DEDUCIBLE
                Conn.Execute SQL
            End If
        End If
        RIVA.MoveNext
    Wend
    RIVA.Close
    If Detallado Then GeneraIVADetallado 5
    
    
    
    '------------------------------------------------------------------
    '------------------------------------------------------------------
    ' NUEVO: 26 Julio 2005             NUEVO NUEVO NUEVO
    '------------------------------------------------------------------
    '------------------------------------------------------------------
    '   Recargo de equivalencia
    '   Recargo de equivalencia
    '   Recargo de equivalencia
    '   Recargo de equivalencia
    '   Recargo de equivalencia
    '   Recargo de equivalencia
    '   Recargo de equivalencia
    '   Recargo de equivalencia
    '------------------------------------------------------------------
    
    'CLIENTES
    '---------
    '---------
    
    HayRecargoEquivalencia = CargarRecargosATratar(True)
    If HayRecargoEquivalencia Then
            SQL = "Select * from tmpliqiva where codusu = " & vUsu.Codigo
            RIVA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RIVA.EOF
                If RIVA.Fields(1) >= 0 Then
                    'TotalIva
                    ImpD = 0: ImpH = 0
                    
                    'Calulamos el IVA, sin el deducible
                    TotalRetencion CStr(RIVA.Fields(1)), 0, False
                    
                    'Insertamos en la USUARIOS.TMPLIQUIVA los valores
                    SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                    d = TransformaComasPuntos(CStr(ImpD))
                    H = TransformaComasPuntos(CStr(ImpH))
                    SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",1)"   '1.- RETENCION
                    Conn.Execute SQL
                    
                    
                    If TieneDeducibles Then
        '                msgbox "Recargo de equivalencia
        '                'Modificacion del 31 Enero 2005.  Facutras con IVA NO deducible
        '                'Es para mostrar los que tienen la marca de NO deducible
        '                ImpD = 0: ImpH = 0
        '
        '                'Calulamos el IVA, sin el deducible
        '                TotalRetencion CStr(RIVA.Fields(1)), 2, False
        '
        '                'Insertamos en la USUARIOS.TMPLIQUIVA los valores
        '                SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
        '                d = TransformaComasPuntos(CStr(ImpD))
        '                H = TransformaComasPuntos(CStr(ImpH))
        '                SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",2)"
        '                Conn.Execute SQL
                    End If
                End If
                
                'Siguiente
                RIVA.MoveNext
            Wend
            RIVA.Close
            
            'Desglose para
            If Detallado Then
                'Insertamos en la tabla tmpimpbalan los valores
                GeneraIVADetallado 1
            End If
    End If
    
    
    'PROVEEDORES
    HayRecargoEquivalencia = CargarRecargosATratar(False)
    
    If HayRecargoEquivalencia Then
            SQL = "Select * from tmpliqiva where codusu = " & vUsu.Codigo
            RIVA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RIVA.EOF
                If RIVA.Fields(1) >= 0 Then
                    'TotalIva
                    ImpD = 0: ImpH = 0
                    
                    'Calulamos el IVA, sin el deducible
                    TotalRetencion CStr(RIVA.Fields(1)), 1, False
                    
                    
                              
                    
                    
                    
                    'Insertamos en la USUARIOS.TMPLIQUIVA los valores
                    SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                    d = TransformaComasPuntos(CStr(ImpD))
                    H = TransformaComasPuntos(CStr(ImpH))
                    SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",5)"   '5.- %RETENCION EN
                    Conn.Execute SQL
                    
                    
                    If TieneDeducibles Then
                        'Modificacion del 31 Enero 2005.  Facutras con IVA NO deducible
                        'Es para mostrar los que tienen la marca de NO deducible
                        ImpD = 0: ImpH = 0
                        
                        'Calulamos el IVA, sin el deducible
                        TotalRetencion CStr(RIVA.Fields(1)), 2, False
                        
                        'Insertamos en la USUARIOS.TMPLIQUIVA los valores
                        SQL = "INSERT INTO Usuarios.zliquidaiva (codusu, iva, bases, ivas, codempre, periodo, ano,cliente) VALUES (" & vUsu.Codigo & ","
                        d = TransformaComasPuntos(CStr(ImpD))
                        H = TransformaComasPuntos(CStr(ImpH))
                        SQL = SQL & TransformaComasPuntos(CStr(RIVA.Fields(1))) & "," & d & "," & H & "," & Empresa & "," & Periodo & "," & Anyo & ",2)"   'El 2 significa DEDUCIBLE
                        Conn.Execute SQL
                    End If
                End If
                
                'Siguiente
                RIVA.MoveNext
            Wend
            RIVA.Close
    
            'Desglose para
            'If Detallado Then
            '    'Insertamos en la tabla tmpimpbalan los valores
            '    GeneraIVADetallado 1
            'End If
    
    End If
    
    
    
    
    
    
    
    
    '
    'Ahora insertamos los intracom y campo
    InsertaDesgloseIVACampoComunitario NumRegElim
    GeneraIVADetallado 9  'El iva detallo
    
    If IvaInversionSujetoPasivo <> "" Then GeneraIVADetallado 8
    
    
    Set RIVA = Nothing
End Function


Private Sub CargarIvasATratar(IvaClientes As Boolean)
    SQL = "Delete from tmpliqiva where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    On Error Resume Next  'Por k si da fallo es k ya estaba introducido
    
    If IvaClientes Then
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    End If
    
    Set RT = New ADODB.Recordset
    For M1 = 1 To 3
        If IvaClientes Then
            SQL = "select pi" & M1 & "faccl,tp" & M1 & "faccl from " & vCta & ".cabfact WHERE " & Aux & " group by pi" & M1 & "faccl"
        Else
            SQL = "select pi" & M1 & "facpr,tp" & M1 & "facpr  from " & vCta & ".cabfactprov WHERE " & Aux & " group by pi" & M1 & "facpr"
        End If
        
        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RT.EOF
            If Not IsNull(RT.Fields(0)) Then
                d = TransformaComasPuntos(CStr(RT.Fields(0)))
                SQL = "INSERT INTO tmpliqiva (codusu, iva) VALUES (" & vUsu.Codigo & "," & d & ")"
                Conn.Execute SQL
                If Err.Number <> 0 Then Err.Clear
            End If
            RT.MoveNext
        Wend
        RT.Close
    Next M1
    Set RT = Nothing
    On Error GoTo 0
End Sub




Private Function CargarRecargosATratar(IvaClientes As Boolean) As Boolean
    CargarRecargosATratar = False
    
    
    SQL = "Delete from tmpliqiva where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    On Error Resume Next  'Por k si da fallo es k ya estaba introducido
    
    If IvaClientes Then
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    End If
    
    Set RT = New ADODB.Recordset
    For M1 = 1 To 3
        If IvaClientes Then
            SQL = "select pr" & M1 & "faccl,tr" & M1 & "faccl from " & vCta & ".cabfact WHERE " & Aux & " group by pr" & M1 & "faccl"
        Else
            SQL = "select pr" & M1 & "facpr,tr" & M1 & "facpr  from " & vCta & ".cabfactprov WHERE " & Aux & " group by pr" & M1 & "facpr"
        End If
        
        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RT.EOF
            If Not IsNull(RT.Fields(0)) Then
                d = TransformaComasPuntos(CStr(RT.Fields(0)))
                SQL = "INSERT INTO tmpliqiva (codusu, iva) VALUES (" & vUsu.Codigo & "," & d & ")"
                Conn.Execute SQL
                If Err.Number <> 0 Then Err.Clear
                CargarRecargosATratar = True
            End If
            RT.MoveNext
        Wend
        RT.Close
    Next M1
    Set RT = Nothing
    On Error GoTo 0
End Function





Private Sub TotalIva(Porcentaje As String, Clientes As Byte, SoloElDeDucible As Boolean)

    Set RT = New ADODB.Recordset
        
    If Clientes = 0 Then
        'Clientes
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        'Proveedores DEDUCIBLE
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        'Marzo 2014
        Aux = Aux & " AND extranje < 1"   'JULIO. POnia <2. Ahora <1
        Aux = Aux & " AND nodeducible = "
    
        'Modificacion 31 Enero
        'PROVEEDORES NO deducibles
        If Clientes = 1 Then
            Aux = Aux & "0"
        Else
            If Clientes = 2 Then
            
                If Not SoloElDeDucible Then
                    Aux = Aux & "1"
                Else
                    'Facturas con tipo IVA NO deducible
                    Aux = Aux & "0"
                End If
                

                'Aqqui.   Extranje puede ser
                
                'Marzo 2014.
                'Solo van los ivas que no tengan marca de extranaje <> 2
                
                
                'Antes Abril 2015
                'Aux = Aux & " AND extranje < 1"
                Aux = Aux & " AND extranje IN (0)"
                
                
            End If
        End If
        
        
        
        
    End If



    
    'Comprobaremos para los tres tipos de iva
    'En el futuro podremos desglosar por tipo de iva, empresa y demas
    d = TransformaComasPuntos(CStr(Porcentaje))
    For A1 = 1 To 3
        If Clientes = 0 Then
            SQL = "select sum(ba" & A1 & "faccl),sum(ti" & A1 & "faccl) from " & vCta & ".cabfact "
            'MODIFICACION 16 MAYO 2005
            ' IVA NO DEDUCIBLE
            SQL = SQL & "," & vCta & ".tiposiva WHERE codigiva = tp" & A1 & "faccl and tipodiva"
            If SoloElDeDucible Then
                SQL = SQL & "="
            Else
                SQL = SQL & "<>"
            End If
            SQL = SQL & "4 AND  pi" & A1 & "faccl="
                
        Else
        
            SQL = "select sum(ba" & A1 & "facpr),sum(ti" & A1 & "facpr) from " & vCta & ".cabfactprov  "

            
            SQL = SQL & "," & vCta & ".tiposiva WHERE codigiva = tp" & A1 & "facpr and tipodiva"
            If SoloElDeDucible Then
                SQL = SQL & "="
            Else
                SQL = SQL & "<>"
            End If
            SQL = SQL & "4 AND  pi" & A1 & "facpr="
            
        End If
        SQL = SQL & d
        SQL = SQL & " AND " & Aux
        
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpD = ImpD + RT.Fields(0)
            If Not IsNull(RT.Fields(1)) Then ImpH = ImpH + RT.Fields(1)
        End If
        RT.Close
    Next A1
    Set RT = Nothing
End Sub




'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'TOTAL RETENCION    nuevo                    26 JULIO 2005
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Private Sub TotalRetencion(Porcentaje As String, Clientes As Byte, SoloElDeDucible As Boolean)

    Set RT = New ADODB.Recordset
        
    If Clientes = 0 Then
        'Clientes
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        'Proveedores DEDUCIBLE
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Aux = Aux & " AND nodeducible = "
    
        'Modificacion 31 Enero
        'PROVEEDORES NO deducibles
        If Clientes = 1 Then
            Aux = Aux & "0"
        Else
            If Clientes = 2 Then
            
                If Not SoloElDeDucible Then
                    Aux = Aux & "1"
                Else
                    'Facturas con tipo IVA NO deducible
                    Aux = Aux & "0"
                End If
            End If
        End If
    End If



    
    'Comprobaremos para los tres tipos de iva
    'En el futuro podremos desglosar por tipo de iva, empresa y demas
    d = TransformaComasPuntos(CStr(Porcentaje))
    For A1 = 1 To 3
        If Clientes = 0 Then
            SQL = "select sum(ba" & A1 & "faccl),sum(tr" & A1 & "faccl) from " & vCta & ".cabfact "
            'MODIFICACION 16 MAYO 2005
            ' IVA NO DEDUCIBLE
            SQL = SQL & "," & vCta & ".tiposiva WHERE codigiva = tp" & A1 & "faccl and tipodiva"
            If SoloElDeDucible Then
                SQL = SQL & "="
            Else
                SQL = SQL & "<>"
            End If
            SQL = SQL & "4 AND  pr" & A1 & "faccl="
                
        Else
        
            SQL = "select sum(ba" & A1 & "facpr),sum(tr" & A1 & "facpr) from " & vCta & ".cabfactprov  "
            'MODIFICACION 16 MAYO 2005
            ' IVA NO DEDUCIBLE
            
            SQL = SQL & "," & vCta & ".tiposiva WHERE codigiva = tp" & A1 & "facpr and tipodiva"
            If SoloElDeDucible Then
                SQL = SQL & "="
            Else
                SQL = SQL & "<>"
            End If
            SQL = SQL & "4 AND  pr" & A1 & "facpr="
            
        End If
        SQL = SQL & d
        SQL = SQL & " AND " & Aux
        
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpD = ImpD + RT.Fields(0)
            If Not IsNull(RT.Fields(1)) Then ImpH = ImpH + RT.Fields(1)
        End If
        RT.Close
    Next A1
    Set RT = Nothing
End Sub







'0.-Cli
'1.- Recargo equivalencia
'2.- PRoveeed
'3.- Recargo equivalencia PRO
'5.- Prov. NO deducible
'7.- Extranjero
'8.-  ISP  Inversion sujeto pasivo

'9.- INTRCOMUNITARIAS

'me falta detallar
'para los recrgos de equivalencia para ello segunsea 0,1 o no deducible sera clientes
'y para 2,3 y no deducbile sera proveedores
'PERO, el 1 sera sobre los RECARGOS DE EQUIVALENCIA
Private Sub GeneraIVADetallado(Clientes As Byte)
Dim C As String
Dim Insertar As Boolean

    

    If Clientes < 2 Then
        Aux = "cl"
    Else
        Aux = "pr"
    End If
    
    Set RT = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    
    
    'Generamos el SQL para la insercion
    SQL = "Select * from  tmpimpbalance WHERE codusu=" & vUsu.Codigo
    'MAYO 2005
    'Los valores de abajo los pondremos a mano
    'SQL = SQL & " AND pasivo=" & Clientes
    'SQL = SQL & " AND codigo ="
    SQL = SQL & " AND pasivo="
    
    d = "INSERT INTO tmpimpbalance (codusu, Pasivo,codigo,importe1, importe2,descripcion,linea ) VALUES ("
    d = d & vUsu.Codigo & ","
    
    If Clientes < 2 Then
        C = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        C = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    End If
    If Clientes > 1 Then
        
        
        If Clientes = 8 Then
            'ISP
            C = C & " AND extranje =3"
        Else
            If Clientes = 9 Then
                C = C & " AND extranje = 1"
            Else
                'Resto
                C = C & " AND extranje < 1"
            End If
        End If
    
        C = C & " AND nodeducible = "
        If Clientes <> 4 Then
            C = C & "0"
        Else
            C = C & "1"
        End If
        
        

    End If
    
    
    '
    'Para las tres bases
    '------------------
    For M1 = 1 To 3
        '1 y 3 son los RECARGOS DE EQUIVALENCIA
        If Clientes <> 1 And Clientes <> 3 Then
            '-----
            'IVAS
            '-----
            Codigo = "SELECT Sum(ba" & M1 & "fac" & Aux & ") AS Sumab, Sum(ti" & M1 & "fac" & Aux & ") AS SumaT, tp" & M1 & "fac" & Aux
            'mayo2005
            'Codigo = Codigo & " From " & vCta & ".cabfact"
            Codigo = Codigo & ",tipodiva From " & vCta & ".tiposiva," & vCta & ".cabfact"
            
            
        
            
        
        Else
            '--------------------
            'RECARGO EQUIVALENCIA
            '--------------------
            Codigo = "SELECT Sum(ba" & M1 & "fac" & Aux & ") AS Sumab, Sum(tr" & M1 & "fac" & Aux & ") AS SumaT, tp" & M1 & "fac" & Aux
            'mayo2005
            Codigo = Codigo & ",tipodiva From " & vCta & ".tiposiva," & vCta & ".cabfact"
            
        
        End If
        If Clientes > 1 Then Codigo = Codigo & "prov"
        
        Codigo = Codigo & " WHERE " & C
        Codigo = Codigo & " AND "
        Codigo = Codigo & vCta & ".tiposiva.codigiva = " & vCta & ".cabfact"

        If Clientes > 1 Then Codigo = Codigo & "prov"
        Codigo = Codigo & ".tp" & M1 & "fac" & Aux
        
        
        'TEngo que separar las facturas deducibles de las no deducibles, en IVA
        'SOLO para proveedores
        
        If Clientes >= 2 Then
            If Clientes = 8 Then
                'ISP
                'No hacemos nada
                'Stop
            Else
                If Clientes = 5 Then
                    Codigo = Codigo & " and tipodiva = 4"
                ElseIf Clientes = 7 Then
                    Codigo = Codigo & " and tipodiva = 5"
                
                Else
                    Codigo = Codigo & " and not tipodiva in (4,5)"
                End If
            End If
                    
        End If
        

        
        Codigo = Codigo & " GROUP BY tp" & M1 & "fac" & Aux
        RT.Open Codigo, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
    
        While Not RT.EOF
            'Para cada tipo, para la empresa esta
            If Not IsNull(RT.Fields(2)) Then
                A2 = RT.Fields(2)
                If IsNull(RT!sumab) Then
                    ImPerD = 0
                Else
                    ImPerD = RT!sumab
                End If
                If IsNull(RT!sumat) Then
                    ImPerH = 0
                Else
                    ImPerH = RT!sumat
                End If
                
                'Si es retencion, el importe tendremos que comprobar que es superior a 0
                Insertar = True
                If Clientes = 1 Or Clientes = 3 Then
                      Insertar = Not (ImPerH = 0)
                End If
                                            'EL DEDUCIBLE
                If Insertar Then
                    InsertaIVADetallado Clientes, RT!tipodiva = 5
                    
                    'Si el IVA es ISP debe añadirse al deducible en su menu
                    If Clientes = 8 Then
                        
                        ' El codigo 2 es el IVA deducible
                        InsertaIVADetallado 2, False
                    End If
                    
                End If
            End If
            RT.MoveNext
        Wend
        RT.Close
    Next M1

    Set miRsAux = Nothing
    Set RT = Nothing
End Sub


'No le pasamos parametros pq las variables k va a utilizar son globales
Private Sub InsertaIVADetallado(Clientes2 As Byte, Nodeducible As Boolean)

'    If Nodeducible Then
'        If Clientes2 = 1 Then Cli2 = 2
'    End If
    
    H = SQL & Clientes2 & " AND Codigo ="
    miRsAux.Open H & A2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ImpD = 0
    ImpH = 0
    If miRsAux.EOF Then
        'No esta insertado
        M2 = 0
    Else
        'Ya esta insertado
        If Not IsNull(miRsAux!Importe1) Then ImpD = miRsAux!Importe1
        If Not IsNull(miRsAux!importe2) Then ImpH = miRsAux!importe2
        M2 = 1
    End If
    miRsAux.Close
    
    ImpD = ImpD + ImPerD
    ImpH = ImpH + ImPerH
    
    
    'Cargamos sobre H
    If M2 = 0 Then
        'Nuevo
        
        'Ponemos el texto del iva
        H = "Select nombriva,porceiva,tipodiva,porcerec FROM " & vCta & ".tiposiva  WHERE codigiva =" & A2
        miRsAux.Open H, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            'Nombre IVA
            H = miRsAux!nombriva & "','"
            If Clientes2 = 1 Or Clientes2 = 3 Then
                'RECARGO EQUIVALENCIA
                H = H & Format(miRsAux!porcerec, "#0.00")
            Else
                H = H & Format(miRsAux!porceiva, "#0.00")
            End If
        Else
            H = "','"
        End If
        
        miRsAux.Close
        H = ",'" & H & "')"
        'SQL
        
            
        H = d & Clientes2 & "," & A2 & "," & TransformaComasPuntos(CStr(ImpD)) & "," & TransformaComasPuntos(CStr(ImpH)) & H
    Else
        'Modificar
        H = "UPDATE tmpimpbalance SET importe1=" & TransformaComasPuntos(CStr(ImpD))
        H = H & ",importe2 =" & TransformaComasPuntos(CStr(ImpH))
        H = H & " WHERE codusu=" & vUsu.Codigo & " AND Pasivo = " & Abs(Clientes2)
        H = H & " AND codigo =" & A2
    End If
    Conn.Execute H
End Sub



Private Function InsertaDesgloseIVACampoComunitario(ByRef CodigoInsercion As Long)
Dim LetraAutoFactura2 As String
Dim TipoAutofactura As Integer
Dim BasesAutofac As Currency
Dim IvaAutofac As Currency
Dim PorcivaIntra As Currency
    
    'Buscamos las facturas con la marca INTRACOM
    Set miRsAux = New ADODB.Recordset
    Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Aux = "Select * from " & vCta & ".cabfact where " & Aux
    Aux = Aux & " AND intracom = 1"
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        With miRsAux
            InsertarDesgloseIntracomCampo Asc("c"), !pi1faccl * 100, !ba1faccl, !ti1faccl, CodigoInsercion
            If Not IsNull(!tp2faccl) Then InsertarDesgloseIntracomCampo Asc("c"), !pi2faccl * 100, !ba2faccl, !ti2faccl, CodigoInsercion
            If Not IsNull(!tp3faccl) Then InsertarDesgloseIntracomCampo Asc("c"), !pi3faccl * 100, !ba3faccl, !ti3faccl, CodigoInsercion
            .MoveNext
        End With
    Wend
    miRsAux.Close
    
    'Buscamos las provee
    Set miRsAux = New ADODB.Recordset
    
    
    
    'JULIO 2014
    'En proveedores, en la liquidacion, NO van las facturas intracomunitarias,
    'si no la parte de las autofacturas de  proveedores
    'con lo cual , las facturas normales (no las de Bienes de inversion, NO las decalramos
    Aux = "Select * from " & vCta & ".cabfactprov," & vCta & ".tiposiva where "
    Aux = Aux & vCta & ".cabfactprov.tp1facpr = " & vCta & ".tiposiva.codigiva AND"
    Aux = Aux & " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Aux = Aux & " AND extranje = 1"
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        With miRsAux
            If !tipodiva = 2 Then
                Aux = "q"
            Else
                Aux = "p"
                Aux = "" 'JULIO 2014
            End If
            If Aux <> "" Then
                InsertarDesgloseIntracomCampo Asc(Aux), !pi1facpr * 100, !ba1facpr, !ti1facpr, CodigoInsercion
                If Not IsNull(!tp2facpr) Then InsertarDesgloseIntracomCampo Asc(Aux), !pi2facpr * 100, !ba2facpr, !ti2facpr, CodigoInsercion
                If Not IsNull(!tp3facpr) Then InsertarDesgloseIntracomCampo Asc(Aux), !pi3facpr * 100, !ba3facpr, !ti3facpr, CodigoInsercion
            End If
            .MoveNext
        End With
    Wend
    miRsAux.Close
    
    'SEPTIEMBRE 2015
    ' A raiz de llamada Montifrut
    ' las intracomunitarias no las saca en el informe
    Aux = "Select * from " & vCta & ".cabfactprov," & vCta & ".tiposiva where "
    Aux = Aux & vCta & ".cabfactprov.tp1facpr = " & vCta & ".tiposiva.codigiva AND"
    Aux = Aux & " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Aux = Aux & " AND extranje = 1"
    Aux = Aux & " AND tipodiva <> 2"
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        With miRsAux
            
                InsertarDesgloseIntracomCampo Asc("p"), !pi1facpr * 100, !ba1facpr, !ti1facpr, CodigoInsercion
                If Not IsNull(!tp2facpr) Then InsertarDesgloseIntracomCampo Asc("p"), !pi2facpr * 100, !ba2facpr, !ti2facpr, CodigoInsercion
                If Not IsNull(!tp3facpr) Then InsertarDesgloseIntracomCampo Asc("p"), !pi3facpr * 100, !ba3facpr, !ti3facpr, CodigoInsercion
      
            .MoveNext
        End With
    Wend
    miRsAux.Close
    
    
    
    
    
    'Metemos en agrario las facturas de proveedores k hayan venido con el regimen especial agrario
    For A3 = 1 To 3
        Aux = "Select * from " & vCta & ".cabfactprov as TF," & vCta & ".tiposiva as TI where TF.tp" & A3 & "facpr = TI.codigiva and TI.tipodiva= 3"
        Aux = Aux & " and fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            Aux = "tp" & A3 & "facpr"
            If Not IsNull(miRsAux.Fields(Aux)) Then
            
                Aux = "ba" & A3 & "facpr"
                ImAcD = miRsAux.Fields(Aux)
                Aux = "ti" & A3 & "facpr"
                ImAcH = miRsAux.Fields(Aux)
                Aux = "pi" & A3 & "facpr"
                InsertarDesgloseIntracomCampo Asc("a"), miRsAux.Fields(Aux) * 100, ImAcD, ImAcH, CodigoInsercion
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next A3
    
    
    
    
    
    'JULIO 2014
    '----------------------------------------------------------------------------------------------------
    

    'DE momento esta a piñon
    'Veremos si en el periodo tiene autofacturas. De ahi sacaraemos el tipo de IVA, el totalbase , totaliva
    'Con ese tipo de iva, nos iremos a cabfactprov y buscaremos en el periodo, para ese tipode iva
    'Si lo tiene, deberian sumar lo mismo. Si no Avisamos, pero continuamos
    TipoAutofactura = -1
    LetraAutoFactura2 = vParam.LetraSerieAutofactura
    If LetraAutoFactura2 = "" Then LetraAutoFactura2 = "@@@"
    
    Aux = "Select tp1faccl from " & vCta & ".cabfact WHERE numserie ='" & LetraAutoFactura2 & "'"
    Aux = Aux & " and fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "' GROUP BY 1"
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Aux = ""
    While Not miRsAux.EOF
        If Aux <> "" Then
            MsgBox "Mas de un tipo de iva para las autofacturas. No se puede detallar", vbExclamation
            TipoAutofactura = -1
        Else
            TipoAutofactura = miRsAux!tp1faccl
            Aux = "S"
        End If
        miRsAux.MoveNext
    
    Wend
    miRsAux.Close
    
    'Otra comprobacion. Si tiene autofacturas
    If TipoAutofactura > 0 Then
    
        'Que las autofacturas solo tienen una base
        Aux = "Select count(*) from " & vCta & ".cabfact WHERE numserie ='" & LetraAutoFactura2 & "'"
        Aux = Aux & " and fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
        Aux = Aux & " AND tp1faccl = " & TipoAutofactura & " AND (tp2faccl >=0 or tp3faccl>=0)"
        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not miRsAux.EOF Then
            If DBLet(miRsAux.Fields(0), "N") > 0 Then
                MsgBox "Facturas intracomunitarias con mas de un tipo de IVA", vbExclamation
                TipoAutofactura = -1
            End If
        End If
        miRsAux.Close
    End If
        
    If TipoAutofactura > 0 Then
        Aux = "Select porceiva  from tiposiva where codigiva= " & TipoAutofactura
        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        PorcivaIntra = 21
        If miRsAux.EOF Then
            MsgBox "Error obteniendo porcentaje iva autofacturas", vbExclamation
        Else
            PorcivaIntra = miRsAux!porceiva
        End If
        miRsAux.Close
        
    
    
        'Calculamos lo que suma
        Aux = "Select sum(ba1faccl),sum(ti1faccl) from " & vCta & ".cabfact WHERE numserie ='" & LetraAutoFactura2 & "'"
        Aux = Aux & " and fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
        Aux = Aux & " AND tp1faccl = " & TipoAutofactura
        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'NO PUEDE SER EOF
        BasesAutofac = miRsAux.Fields(0)
        IvaAutofac = miRsAux.Fields(1)
        miRsAux.Close
        
        
        'Podriamos ver si el iva de las autofacuras en proveedore suma lo mismo que este
        Aux = "Select sum(ba1facpr),sum(ti1facpr) from " & vCta & ".cabfactprov WHERE numfacpr like '" & LetraAutoFactura2 & "%'"
        Aux = Aux & " and fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Aux = Aux & " AND tp1facpr = " & TipoAutofactura
        miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Aux = ""
        If miRsAux.EOF Then
            Aux = "No se han encontrado las autofacturas en proveedores"
        Else
            'NO PUEDE SER EOF
            If DBLet(miRsAux.Fields(0), "N") <> BasesAutofac Then Aux = Aux & " - No coinciden las bases de autofacturas"
            If DBLet(miRsAux.Fields(1), "N") <> IvaAutofac Then Aux = Aux & " - No coinciden el importe iva de autofacturas"
        End If
        miRsAux.Close
        If Aux <> "" Then
            Aux = "Detallando autofacturas " & vbCrLf & vbCrLf & Aux
            MsgBox Aux, vbExclamation
            
        Else
            InsertarDesgloseIntracomCampo Asc("c"), PorcivaIntra * 100, BasesAutofac, IvaAutofac, CodigoInsercion
            'Para provedores tb, yq que arriba lo hemos comentado el fuente
            InsertarDesgloseIntracomCampo Asc("p"), PorcivaIntra * 100, BasesAutofac, IvaAutofac, CodigoInsercion
            
        End If
    
    End If
    Set miRsAux = Nothing
End Function

'Tipo tendra un entero de la forma: 16% --> 1600
'--- etc
Private Function InsertarDesgloseIntracomCampo(Subtipo As Byte, Tipo As Integer, Base As Currency, impIVA As Currency, ByRef Co As Long)
    Co = Co + 1
    d = "INSERT INTO tmpctaexplotacioncierre (codusu, cta, acumPerD, acumPerH) VALUES ("
    d = d & vUsu.Codigo & ",'"
    d = d & Chr(Subtipo) & Format(Co, "00000") & Format(Tipo, "0000") & "',"
    d = d & TransformaComasPuntos(CStr(Base)) & ","
    d = d & TransformaComasPuntos(CStr(impIVA)) & ")"
    Conn.Execute d
    
End Function

'1 Clientes
'3 Proveedores
Private Sub GeneraRetencionDetallado(Clientes As Byte)
Dim C As String

    

    If Clientes = 1 Then
        Aux = "cl"
    Else
        Aux = "pr"
    End If
    
    Set RT = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    
    
    'Generamos el SQL para la insercion
    SQL = "Select * from  tmpimpbalance WHERE codusu=" & vUsu.Codigo
    'MAYO 2005
    'Los valores de abajo los pondremos a mano
    'SQL = SQL & " AND pasivo=" & Clientes
    'SQL = SQL & " AND codigo ="
    SQL = SQL & " AND pasivo="
    
    d = "INSERT INTO tmpimpbalance (codusu, Pasivo,codigo,importe1, importe2,descripcion,linea ) VALUES ("
    d = d & vUsu.Codigo & ","
    
    If Clientes = 0 Then
        C = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        C = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        C = C & " AND nodeducible = 0"
    End If
    
    
    '
    'Para las tres bases
    '------------------
    For M1 = 1 To 3
        Codigo = "SELECT Sum(ba" & M1 & "fac" & Aux & ") AS Sumab, Sum(tr" & M1 & "fac" & Aux & ") AS SumaT, tp" & M1 & "fac" & Aux
        'mayo2005
        'Codigo = Codigo & " From " & vCta & ".cabfact"
        Codigo = Codigo & ",tipodiva From " & vCta & ".tiposiva," & vCta & ".cabfact"
        
        If Clientes > 0 Then Codigo = Codigo & "prov"
        
        'mayo2005
        Codigo = Codigo & " WHERE " & C
        Codigo = Codigo & " AND "
        Codigo = Codigo & vCta & ".tiposiva.codigiva = " & vCta & ".cabfact"

        If Clientes > 0 Then Codigo = Codigo & "prov"
        Codigo = Codigo & ".tp" & M1 & "fac" & Aux
        
        
        Codigo = Codigo & " GROUP BY tp" & M1 & "fac" & Aux
        RT.Open Codigo, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
    
        While Not RT.EOF
            'Para cada tipo, para la empresa esta
            If Not IsNull(RT.Fields(2)) Then
                A2 = RT.Fields(2)
                If IsNull(RT!sumab) Then
                    ImPerD = 0
                Else
                    ImPerD = RT!sumab
                End If
                If IsNull(RT!sumat) Then
                    ImPerH = 0
                Else
                    ImPerH = RT!sumat
                End If
                
              
                                            'EL DEDUCIBLE
                InsertaIVADetallado Clientes, RT!tipodiva = 4
            End If
            RT.MoveNext
        Wend
        RT.Close
    Next M1

    Set miRsAux = Nothing
    Set RT = Nothing
End Sub


Private Sub IvaDeducibleBienInversion(Porcentaje As String)
Dim R As ADODB.Recordset

    ImPerD = 0: ImPerH = 0
    Set R = New ADODB.Recordset
    For M1 = 1 To 3
        Codigo = "SELECT Sum(ba" & M1 & "facpr) AS Sumab, Sum(ti" & M1 & "facpr) AS SumaT"
        Codigo = Codigo & " From " & vCta & ".tiposiva tiposiva," & vCta & ".cabfactprov cabfactprov"
        Codigo = Codigo & " WHERE cabfactprov.tp" & M1 & "facpr = tiposiva.codigiva and tipodiva=2"  'Bien inversion  #####PONIA un 3
        Codigo = Codigo & " AND fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Codigo = Codigo & " AND pi" & M1 & "facpr = " & TransformaComasPuntos(Porcentaje)
        R.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not R.EOF Then
            If Not IsNull(R!sumab) Then
                ImPerD = ImPerD + R!sumab
                ImPerH = ImPerH + DBLet(R!sumat, "N")
            End If
        End If
        R.Close
    Next M1
    Set R = Nothing
    
End Sub


Private Sub IvaDeducibleExtranjero(Porcentaje As String)
Dim R As ADODB.Recordset

    ImCierrD = 0: ImCierrH = 0
    Set R = New ADODB.Recordset
    For M1 = 1 To 3
        Codigo = "SELECT Sum(ba" & M1 & "facpr) AS Sumab, Sum(ti" & M1 & "facpr) AS SumaT"
        Codigo = Codigo & " From " & vCta & ".tiposiva tiposiva," & vCta & ".cabfactprov cabfactprov"
        Codigo = Codigo & " WHERE cabfactprov.tp" & M1 & "facpr = tiposiva.codigiva and tipodiva=5"  'EXTRANJERO
        Codigo = Codigo & " AND fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Codigo = Codigo & " AND pi" & M1 & "facpr = " & TransformaComasPuntos(Porcentaje)
        R.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not R.EOF Then
            If Not IsNull(R!sumab) Then
                ImCierrD = ImCierrD + R!sumab
                ImCierrH = ImCierrH + DBLet(R!sumat, "N")
            End If
        End If
        R.Close
    Next M1
    Set R = Nothing
    
End Sub


Private Sub IvaDelaInversionSujetoPasivo(Porcentaje As String)
Dim R As ADODB.Recordset

    ImAcD = 0: ImAcH = 0
    Set R = New ADODB.Recordset
    For M1 = 1 To 3
        Codigo = "SELECT Sum(ba" & M1 & "facpr) AS Sumab, Sum(ti" & M1 & "facpr) AS SumaT"
        Codigo = Codigo & " From " & vCta & ".cabfactprov cabfactprov"
        Codigo = Codigo & " WHERE  "
        Codigo = Codigo & "  fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Codigo = Codigo & " AND pi" & M1 & "facpr = " & TransformaComasPuntos(Porcentaje)
        Codigo = Codigo & " AND extranje=3"
        R.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not R.EOF Then
            If Not IsNull(R!sumab) Then
                ImAcD = ImAcD + R!sumab
                ImAcH = ImAcH + DBLet(R!sumat, "N")
            End If
        End If
        R.Close
    Next M1
    Set R = Nothing
    
End Sub



'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'   Consulta extractos sobre ejercicios traspasados
'
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Public Function CargaDatosConExtCerrados(ByRef Cuenta As String, fec1 As Date, fec2 As Date, ByRef vSQL As String, ByRef DescCuenta As String, ByRef AnyoInicioEjercicio As String) As Byte
Dim ACUM As Currency  'Acumulado anterior

On Error GoTo ECargaDatosConExtCerrados
CargaDatosConExtCerrados = 1
'Insertamos en los campos de cabecera de cuentas
NombreSQL DescCuenta
SQL = Cuenta & "    -    " & DescCuenta
SQL = "INSERT INTO tmpconextcab (codusu,cta,fechini,fechfin,cuenta) VALUES (" & vUsu.Codigo & ", '" & Cuenta & "','" & Format(fec1, "dd/mm/yyyy") & "','" & Format(fec2, "dd/mm/yyyy") & "','" & SQL & "')"
Conn.Execute SQL

'los totatales
'If Not CargaAcumuladosTotalesCerrados(Cuenta) Then Exit Function

'Los caumulados anteriores
If Not CargaAcumuladosAnterioresCerrados(Cuenta, fec1, ACUM, AnyoInicioEjercicio) Then Exit Function

'GENERAMOS LA TBLA TEMPORAL
If Not CargaTablaTemporalConExtCerrados(Cuenta, vSQL, ACUM) Then Exit Function

CargaDatosConExtCerrados = 0
Exit Function
ECargaDatosConExtCerrados:
    CargaDatosConExtCerrados = 2
    MuestraError Err.Number, "Gargando datos temporales. Cta: " & Cuenta, Err.Description
End Function



Private Function CargaAcumuladosTotalesCerrados(ByRef Cta As String) As Boolean
    CargaAcumuladosTotalesCerrados = False
    ImpD = 0
    ImpH = 0
    SQL = "UPDATE tmpconextcab SET acumtotD= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    SQL = SQL & ", acumtotH= " & TransformaComasPuntos(CStr(ImpH)) 'Format(ImpH, "#,###,##0.00")
    ImpD = ImpD - ImpH
    SQL = SQL & ", acumtotT= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL
    CargaAcumuladosTotalesCerrados = True
End Function


Private Function CargaAcumuladosAnterioresCerrados(ByRef Cta As String, ByRef FI As Date, ByRef ACUM As Currency, ByRef AnyoInicioEjercicio As String) As Boolean
Dim UnaAnyoMenos As Boolean
    CargaAcumuladosAnterioresCerrados = False
    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    SQL = SQL & " from hlinapu1 where codmacta='" & Cta & "'"
    H = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & AnyoInicioEjercicio
    
    'Year (FI)
    SQL = SQL & " AND fechaent >=  '" & Format(H, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <  '" & Format(FI, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
    Else
        ImpD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
    End If
    RT.Close
    ACUM = ImpD - ImpH
    SQL = "UPDATE tmpconextcab SET acumantD= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & ", acumantH= " & TransformaComasPuntos(CStr(ImpH))
    SQL = SQL & ", acumantT= " & TransformaComasPuntos(CStr(ACUM))
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL
    Set RT = Nothing
    CargaAcumuladosAnterioresCerrados = True
End Function



Private Function CargaTablaTemporalConExtCerrados(Cta As String, vSele As String, ByRef ACUM As Currency) As Boolean
Dim Aux As Currency
Dim ImporteD As String
Dim ImporteH As String
Dim Contador As Long
Dim RC As String
On Error GoTo Etmpconext


CargaTablaTemporalConExtCerrados = False

'Conn.Execute "Delete from tmpconext where codusu =" & vUsu.Codigo
Set RT = New ADODB.Recordset
SQL = "Select * from hlinapu1 where codmacta='" & Cta & "'"
SQL = SQL & " AND " & vSele & " ORDER BY fechaent,numasien,linliapu"
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
SQL = "INSERT INTO tmpconext (codusu, POS,numdiari, fechaent, numasien, linliapu, timporteD, timporteH, saldo, Punteada,nomdocum,ampconce,cta,contra,ccost) VALUES ("
'ImpD = 0 ASI LLEVAMOS EL ACUMULADO
'ImpH = 0
Contador = 0
While Not RT.EOF
    Contador = Contador + 1
    
    If Not IsNull(RT!timported) Then
        Aux = RT!timported
        ImpD = ImpD + Aux
        ImporteD = TransformaComasPuntos(RT!timported)
        ImporteH = "Null"
    Else
        Aux = RT!timporteH
        ImporteD = "Null"
        ImporteH = TransformaComasPuntos(RT!timporteH)
        ImpH = ImpH + Aux
        Aux = -1 * Aux
    End If
    ACUM = ACUM + Aux
    
    'Insertar
    RC = vUsu.Codigo & "," & Contador & "," & RT!NumDiari & ",'" & Format(RT!fechaent, FormatoFecha) & "'," & RT!Numasien & "," & RT!Linliapu & ","
    RC = RC & ImporteD & "," & ImporteH
    If RT!punteada <> 0 Then
        ImporteD = "SI"
        Else
        ImporteD = ""
    End If
    RC = RC & "," & TransformaComasPuntos(CStr(ACUM)) & ",'" & ImporteD & "','"
    RC = RC & RT!numdocum & "','" & DevNombreSQL(RT!ampconce) & "','" & Cta & "',"
    If IsNull(RT!ctacontr) Then
        RC = RC & "NULL"
    Else
        RC = RC & "'" & RT!ctacontr & "'"
    End If
    RC = RC & ","
    If IsNull(RT!codccost) Then
        RC = RC & "NULL"
    Else
        RC = RC & "'" & RT!codccost & "'"
    End If
    RC = RC & ")"
    Conn.Execute SQL & RC
    'Sig
    RT.MoveNext
Wend
RT.Close

    SQL = "UPDATE tmpconextcab SET acumperD= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & ", acumperH= " & TransformaComasPuntos(CStr(ImpH))
    ImpD = ImpD - ImpH
    SQL = SQL & ", acumperT= " & TransformaComasPuntos(CStr(ImpD))
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute SQL

    CargaTablaTemporalConExtCerrados = True

Exit Function
Etmpconext:
    MuestraError Err.Number, "Generando datos saldos"
    Set RT = Nothing
End Function



Public Function BloqueoManual(Bloquear As Boolean, Tabla As String, Clave As String) As Boolean
    If Bloquear Then
        SQL = "INSERT INTO zbloqueos (codusu, tabla, clave) VALUES (" & vUsu.Codigo
        SQL = SQL & ",'" & UCase(Tabla) & "','" & UCase(Clave) & "')"
    Else
        SQL = "DELETE FROM zbloqueos where codusu = " & vUsu.Codigo & " AND tabla ='"
        SQL = SQL & Tabla & "'"
    End If
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        Err.Clear
        BloqueoManual = False
    Else
        BloqueoManual = True
    End If
End Function





'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'   Cuenta explotacion por centro de coste
'
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Cadena fechas tendra, enpipado, mesinicio,añoinicio, mespedido,anopedido,mesfin,anofin
'
Public Sub CtaExploCentroCoste(ByRef CuentaYCC As String, AcuPosterior As Boolean, UltimaFechaHco As Date)
Dim HcoTb As Boolean
    
    
    vCta = RecuperaValor(CuentaYCC, 1)
    Codigo = RecuperaValor(CuentaYCC, 2)
    
    Set RT = New ADODB.Recordset
    
    
    'Acumulado anterior
    '------------------
    HcoTb = False
    If vFecha1 <= UltimaFechaHco Then
        HcoTb = True
    Else
        If vFecha2 <= UltimaFechaHco Then HcoTb = True
    End If
    ImportesCtaExpCC M1, A1, M2 - 1, A2, HcoTb
    ImAcD = ImpD
    ImAcH = ImpH
        
        
    'Acumulado del periodo
    HcoTb = vFecha2 <= UltimaFechaHco
    ImportesCtaExpCC M2, A2, M2, A2, HcoTb
    ImPerD = ImpD
    ImPerH = ImpH
    
    'Estamos pidiendo el acumulado posterior al periodo
    If AcuPosterior Then
        HcoTb = False
        If vFecha2 <= UltimaFechaHco Then
            HcoTb = True
        Else
            If VFecha3 <= UltimaFechaHco Then HcoTb = True
        End If
        ImportesCtaExpCC M2 + 1, A2, M3, A3, HcoTb
        ImCierrD = ImpD
        ImCierrH = ImpH
    Else
        ImCierrD = 0
        ImCierrH = 0
    End If
    
    
    'Ahora ya tenemos todo, nos generar el sql  : ImporteASQL  zctaexpcc
    'codccost, nomccost, codmacta, nommacta, acumD, acumH, perid, periH, postD, postH
    If ImAcD = 0 Then
        If ImAcH = 0 Then
            If ImPerD = 0 Then
                If ImPerH = 0 Then
                    If ImCierrD = 0 Then
                        If ImCierrH = 0 Then
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    'ACumulado
    ImpD = ImAcD + ImPerD + ImCierrD
    ImpH = ImAcH + ImPerH + ImCierrH
    ImpD = ImpD - ImpH
    If ImpD = 0 Then
        ImpD = 0
        ImpH = 0
    Else
        If ImpD < 0 Then
            ImpH = Abs(ImpD)
            ImpD = 0
        Else
            ImpH = 0
        End If
    End If
        
    
    H = RecuperaValor(CuentaYCC, 4)
    d = "'" & Codigo & "','" & H & "','"
    H = RecuperaValor(CuentaYCC, 3)
    d = d & vCta & "','" & H & "'"
    
    'Los importe
    d = d & ImporteASQL2(ImAcD)
    d = d & ImporteASQL2(ImAcH)
    d = d & ImporteASQL2(ImPerD)
    d = d & ImporteASQL2(ImPerH)
    d = d & ImporteASQL2(ImCierrD)
    d = d & ImporteASQL2(ImCierrH)
    
    
    
    
    
    'Saldos
    d = d & ImporteASQL2(ImpD)
    d = d & ImporteASQL2(ImpH)
    H = SQL & d & ")"
    'On Error Resume Next
    Conn.Execute H
End Sub





Public Sub AjustaValoresCtaExpCC(CadenaFechas As String)
    M1 = Val(RecuperaValor(CadenaFechas, 1))
    A1 = Val(RecuperaValor(CadenaFechas, 2))
    M2 = Val(RecuperaValor(CadenaFechas, 3))
    A2 = Val(RecuperaValor(CadenaFechas, 4))
    M3 = Val(RecuperaValor(CadenaFechas, 5))
    A3 = Val(RecuperaValor(CadenaFechas, 6))
    vFecha1 = CDate("01/" & M1 & "/" & A1)
    vFecha2 = CDate("01/" & M2 & "/" & A2)
    VFecha3 = CDate("01/" & M3 & "/" & A3)
    SQL = "INSERT INTO Usuarios.zctaexpcc (codusu, codccost, nomccost, codmacta, nommacta, acumD, acumH, perid, periH, postD, postH,saldoD,SaldoH)"
    SQL = SQL & " VALUES (" & vUsu.Codigo & ","
    
    
End Sub
    
    
    
Private Sub ImportesCtaExpCC(MInicio As Integer, AnoInicio As Integer, MesFin As Integer, AnoFin As Integer, DeHco As Boolean)
Dim SQL As String
    
On Error GoTo EImportesCtaExpCC
    SQL = "select sum(debccost),sum(habccost),codccost,codmacta from hsaldosanal"
    If AnoInicio = AnoFin Then
        Aux = " codmacta = '" & vCta & "'"
        Aux = Aux & " AND codccost = '" & Codigo & "'"
        Aux = Aux & " AND anoccost = " & AnoFin
        Aux = Aux & " AND mesccost >= " & MInicio & " AND mesccost <=" & MesFin
    Else
    
        'Como no dejo que haya mas de una año entre las fechas
        Aux = "(codmacta = '" & vCta & "'"
        Aux = Aux & " AND codccost = '" & Codigo & "'"
        Aux = Aux & " AND anoccost = " & AnoInicio & " AND mesccost >=" & MInicio & ")"
        
        Aux = Aux & " OR (codmacta = '" & vCta & "'"
        Aux = Aux & " AND codccost = '" & Codigo & "'"
        Aux = Aux & " AND anoccost =" & AnoFin & " AND mesccost <= " & MesFin & ")"
    End If
    Aux = " WHERE " & Aux & " GROUP BY codmacta"
    'Hacemos en hsaldosanal
    ImpD = 0
    ImpH = 0
    RT.Open SQL & Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImpD = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImpH = RT.Fields(1)
    End If
    RT.Close
    
    'Si tenemos k sacar desde hco sacamos tb
    If Not DeHco Then Exit Sub
    
    RT.Open SQL & "1 " & Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImpD = ImpD + RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImpH = ImpH + RT.Fields(1)
    End If
    RT.Close
    Exit Sub
EImportesCtaExpCC:
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    Err.Clear
End Sub






'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------













'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'           L I B R O        R E S U M E N
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'
'
'
'
'
'   Para los asiento k haya k quitar saldos estos los guardaremos en la tabla
'
'


Public Sub FijaValoresLibroResumen(FIni As Date, fFin As Date, Nivel As Integer, EjerciciosCerr As Boolean, NumAsiento As String)
    SQL = "INSERT INTO Usuarios.zdirioresum (codusu, clave, fecha, asiento, cuenta, titulo, concepto, debe, haber) VALUES (" & vUsu.Codigo & ","
    vFecha1 = FIni
    vFecha2 = fFin
    A3 = Nivel
    EjerciciosCerrados = EjerciciosCerr
    
    'Numero de asiento
    A2 = 1
    If NumAsiento <> "" Then
        If IsNumeric(NumAsiento) Then A2 = CInt(NumAsiento)
    End If
        
        
    'M2 sera el contador para cada registro
    M2 = 1
End Sub



Public Function ProcesaLibroResumen(Mes As Long, Ano As Integer, I1 As Currency, I2 As Currency)
Dim Opcion As Byte
    ' 0.- Mes normal
    ' 1.- Mes con apertura
    ' 2.- Mes con cierre
Dim TienekKitar As Boolean

    A1 = Ano
    M1 = CInt(Mes)
  

    
    Set RT = New ADODB.Recordset
    'Comprobamos si tiene el mes de apertura de ejercicio
    Opcion = 0
    If M1 = Month(vFecha1) Then
        Opcion = 1
    Else
        If (M1 = Month(vFecha2)) Then Opcion = 2
    End If
        
        
    TienekKitar = False
    If Opcion = 1 Then
        TienekKitar = True
        NumAsiento = A2
        GeneraAperturaResumen 0
    Else
        If Opcion = 2 Then
            TienekKitar = True
            NumAsiento = A2 + 1
            GeneraAperturaResumen 1
            GeneraAperturaResumen 2
            GeneraAperturaResumen 3
            NumAsiento = A2
        Else
            NumAsiento = A2
        End If
    End If
    
    'hacemos el mes
    If I1 <> 0 Or I2 <> 0 Then
        'Insertaremos el acumulado que nos han indicado
        vCta = CStr(DiasMes(CByte(M1), A1))
        vCta = vCta & "/" & M1 & "/" & A1
        VFecha3 = CDate(vCta)
        vCta = "'','ACUMULADO ANTERIOR'"
        Codigo = ""
        
        ImpD = I1
        ImpH = I2
        InsertaParaListadoDiarioResum
    End If
    HacerMes TienekKitar
    A2 = NumAsiento
    'Si tiene fin hacer fin
    Set RT = Nothing
End Function




Private Sub HacerMes(HayKRestarSaldos As Boolean)
Dim RTT As Recordset
   If EjerciciosCerrados Then
       vCta = "1"
   Else
       vCta = ""
   End If
   Aux = "select sum(impmesde),sum(impmesha),cuentas.codmacta,cuentas.nommacta from hsaldos" & vCta
   Aux = Aux & ",cuentas where hsaldos" & vCta
   Aux = Aux & ".codmacta=cuentas.codmacta and "
   Aux = Aux & " cuentas.codmacta like '" & Mid("__________", 1, A3) & "'"   'Nivel
   Aux = Aux & " and anopsald=" & A1
   Aux = Aux & " and mespsald=" & M1
   Aux = Aux & " group by codmacta order by codmacta"
   Set RTT = New ADODB.Recordset
   RTT.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   If RTT.EOF Then
        RTT.Close
        Exit Sub
    End If
    
    
   d = CStr(DiasMes(CByte(M1), A1))
   d = d & "/" & M1 & "/" & A1
   VFecha3 = CDate(d)
   'para cada valor insertaremos en la tabla
   While Not RTT.EOF
        d = RTT.Fields(2)
        If HayKRestarSaldos Then
            FijaImporteResta (d)
        Else
            ImAcD = 0
            ImAcH = 0
        End If
        
        ImpD = 0
        ImpH = 0
        If Not IsNull(RTT.Fields(0)) Then ImpD = RTT.Fields(0)
        If Not IsNull(RTT.Fields(1)) Then ImpH = RTT.Fields(1)
   
        ImpD = ImpD - ImAcD
        ImpH = ImpH - ImAcH
        
        
        vCta = "'" & RTT.Fields(2) & "','" & RTT.Fields(3) & "'"
        Codigo = ""
        InsertaParaListadoDiarioResum
                
        
        RTT.MoveNext
    Wend
    RTT.Close
    Set RTT = Nothing
    NumAsiento = NumAsiento + 1
End Sub






Private Sub FijaImporteResta(ByRef KCuenta As String)
Dim Au As String


    Au = "Select Debe, Haber from tmpdiarresum WHERE codusu =" & vUsu.Codigo & " AND codmacta ='" & KCuenta & "';"
    ImAcD = 0
    ImAcH = 0
    RT.Open Au, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImAcD = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImAcH = RT.Fields(1)
    End If
    RT.Close
    
End Sub






'Opcion
'   0.- Apertura
'   1.- PyG
'   2.- Cierre
'   3.- pyG y cierre
Private Sub GeneraAperturaResumen(Opcion As Byte)
Dim RS As Recordset

    Conn.Execute "DELETE from tmpdiarresum where codusu =" & vUsu.Codigo
    
    Aux = "Select codmacta from hlinapu"
    If EjerciciosCerrados Then Aux = Aux & "1"
    Aux = Aux & " WHERE "
    Select Case Opcion
    Case 0
        'Apertura. El primero
        Aux = Aux & "codconce = 970"
    Case 1
        'Py G
        'Eje: Diciembre:    234
        '     Py G:         235
        '     Cierre:       236
        Aux = Aux & "codconce = 960"
    Case 2
        Aux = Aux & "codconce = 980"
    Case 3
        Aux = Aux & "(codconce = 960 or codconce = 980)"
        'Este no insertara
    End Select
    'Fechas
    Aux = Aux & " AND fechaent >='" & Format(vFecha1, FormatoFecha)
    Aux = Aux & "' AND fechaent <='" & Format(vFecha2, FormatoFecha)
    Aux = Aux & "' GROUP BY codmacta"
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, Conn, adOpenForwardOnly, adCmdText
    While Not RS.EOF
        vCta = RS.Fields(0)
        vCta = Mid(vCta, 1, A3)
        Insertatmpdiarresum
        RS.MoveNext
    Wend
    RS.Close
    
       
    'Ya tenemos en tmpdiarresum
    'las subcuentas del diario
    Codigo = "Select * from tmpdiarresum where codusu =" & vUsu.Codigo
    RS.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        vCta = RS.Fields(1)
        'para cada cuenta, obtienes los importes
        CalcularImporteCierreAperturaPyG Opcion
        If ImCierrD <> 0 Or ImCierrH <> 0 Then
            d = TransformaComasPuntos(CStr(ImCierrD))
            H = TransformaComasPuntos(CStr(ImCierrH))
            InsertarTMP
        Else
            Conn.Execute "DELETE from tmpdiarresum where codusu =" & vUsu.Codigo & " AND codmacta='" & vCta & "';"
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    
    'Ahora si la opcion es 3 no seguimos. Esto de abajo es para insertar
    If Opcion = 3 Then Exit Sub
    
    'Volvemos abri el temporal del diario resumen
    Codigo = "select tmpdiarresum.*, nommacta from tmpdiarresum,cuentas where tmpdiarresum.codmacta = cuentas.codmacta and codusu =" & vUsu.Codigo
    Codigo = Codigo & " order by codmacta"
    RS.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Select Case Opcion
    Case 0
        VFecha3 = vFecha1
        Codigo = "APERTURA"
    Case 1
        Codigo = "PERDIDAS Y GANACIAS"
        VFecha3 = vFecha2
    Case 2
        Codigo = "CIERRE"
        VFecha3 = vFecha2
    End Select
    Codigo = Codigo & " AL "
    
        
    If Not RS.EOF Then
        While Not RS.EOF
            ImpD = RS.Fields(2)
            ImpH = RS.Fields(3)
            vCta = "'" & RS.Fields(1) & "','" & RS.Fields(4) & "'"
            InsertaParaListadoDiarioResum
            RS.MoveNext
        Wend
        'Aumentamos el contador
        NumAsiento = NumAsiento + 1
    End If
    RS.Close
    Set RS = Nothing
End Sub


Private Sub InsertarTMP()
    Aux = "UPDATE tmpdiarresum set debe = " & d & ", Haber = " & H
    Aux = Aux & " WHERE codmacta = '" & vCta & "' and codusu = " & vUsu.Codigo
    Conn.Execute Aux
End Sub

Private Sub Insertatmpdiarresum()
On Error Resume Next
Conn.Execute "INSERT INTO tmpdiarresum (codusu, codmacta) VALUES (" & vUsu.Codigo & ",'" & vCta & "')"
If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub InsertaParaListadoDiarioResum()
If ImpD <> 0 Then
    d = TransformaComasPuntos(CStr(ImpD))
    H = "NULL"
    InsertaParaListadoDiarioResum2 "DEBE"
End If
If ImpH <> 0 Then
    H = TransformaComasPuntos(CStr(ImpH))
    d = "NULL"
    InsertaParaListadoDiarioResum2 "HABER"
End If



End Sub

Private Sub InsertaParaListadoDiarioResum2(DebeHaber As String)
Dim C As String

' clave, fecha, asiento, cuenta, titulo, concepto, debe, haber)
C = SQL & M2 & ",'" & Format(VFecha3, "dd/mm/yyyy") & "'," & NumAsiento & "," & vCta
C = C & ",'" & Codigo & DebeHaber & "'," & d & "," & H & ")"
On Error Resume Next
Conn.Execute C
If Err.Number <> 0 Then MuestraError Err.Number, "Insertando en temporal para informes"
M2 = M2 + 1
End Sub



Private Sub CalcularImporteCierreAperturaPyG(Opcion As Byte)

    Aux = "Select SUM(timporteD),sum(timporteH) from hlinapu"
    If EjerciciosCerrados Then Aux = Aux & "1"
    Aux = Aux & " WHERE codmacta  like '" & vCta & "%"
    Aux = Aux & "' and "
    Select Case Opcion
    Case 0
        'Apertura. El primero
        Aux = Aux & "codconce = 970"
    Case 1
        Aux = Aux & "codconce = 960"
    Case 2
        Aux = Aux & "codconce = 980"
    Case 3
        Aux = Aux & "(codconce = 960 or codconce = 980)"
        'Este no insertara
    End Select
    Aux = Aux & " AND fechaent >= '" & Format(vFecha1, FormatoFecha) & "'"
    Aux = Aux & " AND fechaent <= '" & Format(vFecha2, FormatoFecha) & "'"
    Aux = Aux & ";"

    
    RT.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImCierrD = 0
    Else
        ImCierrD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImCierrH = 0
    Else
        ImCierrH = RT.Fields(1)
    End If
    RT.Close
End Sub





'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'           Detalles cuenta por centro de coste
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'
'
'
'
'
'
'
'

Public Sub FijaValoresCtapoCC(FIniEjer As Date, FIni As Date, fFin As Date, EjerciciosCerr As Boolean)
    'SQL para las lineas
    SQL = "INSERT INTO Usuarios.zlinccexplo (codusu, codccost, codmacta, linapu, docum,"
    SQL = SQL & "fechaent, ampconce, perD, perH, saldo,ctactra, desctra) VALUES ("
    SQL = SQL & vUsu.Codigo & ","
    
    'La cabecera
    Aux = "INSERT INTO Usuarios.zcabccexplo (codusu, codccost, codmacta, nommacta, "
    Aux = Aux & " nomccost,TieneAcum, acumD, acumH, acumS, totD, totH, totS) VALUES ("
    Aux = Aux & vUsu.Codigo & ","
    
    vFecha1 = FIni
    vFecha2 = fFin
    VFecha3 = FIniEjer
    EjerciciosCerrados = EjerciciosCerr
End Sub

'En nombres iran empipaditos nommacta y nomccost
Public Sub Cta_por_CC(ByRef vCuenta As String, vCCos As String, Nombres As String)
Dim miSQL As String
Dim Cad As String
    vCta = vCuenta
    Codigo = vCCos
    Set RT = New ADODB.Recordset
    If vFecha1 > VFecha3 Then
        'Calculamos anteriores
        ImAcD = 0
        ImAcH = 0
        CalculaAnterioresCtaPorCC
    Else
        ImAcD = 0
        ImAcH = 0
    End If
    'En impcierrD llevare el saldo
    ImCierrD = ImAcD - ImAcH
    
    'Importes totales
    ImPerD = ImAcD
    ImPerH = ImAcH
    
'    miSQL = "Select * from hlinapu"
'    If EjerciciosCerrados Then miSQL = miSQL & "1"
'    miSQL = miSQL & " WHERE codmacta ='" & vCta & "'"
'    miSQL = miSQL & " AND codccost ='" & Codigo & "'"
'    miSQL = miSQL & " AND fechaent >='" & Format(vFecha1, FormatoFecha) & "'"
'    miSQL = miSQL & " AND fechaent <='" & Format(vFecha2, FormatoFecha) & "'"
'    miSQL = miSQL & " ORDER BY fechaent, linliapu"

'SELECT Tabla1.num, Tabla1.cta, cuentas1.nommacta
'FROM Tabla1 LEFT JOIN cuentas1 ON Tabla1.cta = cuentas1.codmacta;

    miSQL = "Select *, cuentas.nommacta from hlinapu"
    If EjerciciosCerrados Then miSQL = miSQL & "1"
    miSQL = miSQL & " LEFT JOIN cuentas ON hlinapu"
    If EjerciciosCerrados Then miSQL = miSQL & "1"
    miSQL = miSQL & ".ctacontr = cuentas.codmacta WHERE hlinapu"
    If EjerciciosCerrados Then miSQL = miSQL & "1"
    miSQL = miSQL & ".codmacta ='" & vCta & "'"
    miSQL = miSQL & " AND codccost ='" & Codigo & "'"
    miSQL = miSQL & " AND fechaent >='" & Format(vFecha1, FormatoFecha) & "'"
    miSQL = miSQL & " AND fechaent <='" & Format(vFecha2, FormatoFecha) & "'"
    miSQL = miSQL & " ORDER BY fechaent, linliapu"



    RT.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    A1 = 0
    miSQL = SQL & "'" & Codigo & "','" & vCta & "',"
    While Not RT.EOF
        A1 = A1 + 1
        Cad = A1 & ",'" & DevNombreSQL(DBLet(RT!numdocum)) & "','"
        Cad = Cad & Format(RT!fechaent, FormatoFecha) & "','" & DevNombreSQL(DBLet(RT!ampconce)) & "',"
        If IsNull(RT!timported) Then
            ImpD = 0
            d = "NULL"
        Else
            ImpD = RT!timported
            d = TransformaComasPuntos(CStr(RT!timported))
        End If
            
        'importe HABER
        If IsNull(RT!timporteH) Then
            ImpH = 0
            H = "NULL"
        Else
            ImpH = RT!timporteH
            H = TransformaComasPuntos(CStr(RT!timporteH))
        End If
        Cad = Cad & d & "," & H & ","
        
        ImPerD = ImPerD + ImpD
        ImPerH = ImPerH + ImpH
        
        'Saldo
        ImCierrH = ImpD - ImpH
        ImCierrD = ImCierrD + ImCierrH
        H = TransformaComasPuntos(CStr(ImCierrD))
        Cad = Cad & H
        
        
        'Ctra partida
        If IsNull(RT!ctacontr) Then
            Cad = Cad & ",'',''"
        Else
            Cad = Cad & ",'" & RT!ctacontr & "','" & DevNombreSQL(DBLet(RT!nommacta)) & "'"
        End If
        
        
        Cad = Cad & ")"
        'Ejecutamos
        Conn.Execute miSQL & Cad
    
        'Sig
        RT.MoveNext
    Wend
    RT.Close
    
    
    
    'La cabecera
    '->INSERT INTO zcabccexplo (codusu, codccost, codmacta, nommacta, nomccost,
    '-->acumD, acumH, acumS, totD, totH, totS) VALUES (
    Cad = "'" & Codigo & "','" & vCta & "','" & DevNombreSQL(RecuperaValor(Nombres, 1))
    Cad = Cad & "','" & DevNombreSQL(RecuperaValor(Nombres, 2)) & "',"
    
    '-----------------------------------------
    'Acumulado anterior
    If ImAcD = 0 And ImAcH = 0 Then
        Cad = Cad & """"""
    Else
        Cad = Cad & """S"""
    End If
    Cad = Cad & ","
    If ImAcH = 0 Then
        H = "NULL"
    Else
        H = TransformaComasPuntos(CStr(ImAcH))
    End If
    'Acumulado anterior
    If ImAcD = 0 Then
        d = "NULL"
    Else
        d = TransformaComasPuntos(CStr(ImAcD))
    End If
    Cad = Cad & d & "," & H
    
    
        
    
    
    'SALDO anterior
    ImAcD = ImAcD - ImAcH
    If ImAcD = 0 Then
        H = "NULL"
    Else
        H = TransformaComasPuntos(CStr(ImAcD))
    End If
    Cad = Cad & "," & H

    '--------------------- TOTALES
    d = TransformaComasPuntos(CStr(ImPerD))
    H = TransformaComasPuntos(CStr(ImPerH))
    Cad = Cad & "," & d & "," & H
    
    'Saldo final
    d = TransformaComasPuntos(CStr(ImCierrD))
    Cad = Cad & "," & d & ")"
    
    Conn.Execute Aux & Cad
    
End Sub


Private Sub CalculaAnterioresCtaPorCC()
Dim C As String

    C = "Select sum(timported),sum(timporteh) from hlinapu"
    If EjerciciosCerrados Then C = C & "1"
    C = C & " WHERE codmacta ='" & vCta & "'"
    C = C & " AND codccost ='" & Codigo & "'"
    C = C & " AND fechaent >='" & Format(VFecha3, FormatoFecha) & "'"
    C = C & " AND fechaent <='" & Format(vFecha1, FormatoFecha) & "'"
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImAcD = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImAcH = RT.Fields(1)
    End If
    RT.Close
End Sub





'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'           Cuenta explotacion comparativa
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'
'
'
'
'
'
'
'
Public Function GeneraCtaExplComparativa(ByRef Mes As Integer, Acumulado As Boolean, Anyo As Integer, UtlimoFechaHCO As Date, Nivel As Integer) As Boolean
On Error GoTo EGeneraCtaExplComparativa
    GeneraCtaExplComparativa = False
    A1 = Anyo
    M1 = Mes
    'Si es -1 siginifca k es a ultimo nivel
    If Nivel < 0 Then
        A3 = vEmpresa.DigitosUltimoNivel
    Else
        A3 = DigitosNivel(Nivel)
    End If
    If Acumulado Then
        NumAsiento = 1
    Else
        NumAsiento = 0
    End If
    
    
    
    
    
    
    
    'ELiminamos datos anteriores
    Conn.Execute "Delete From Usuarios.zexplocomp where codusu = " & vUsu.Codigo
    
    'Para la insercion
    d = "INSERT INTO Usuarios.zexplocomp (codusu, cta, cuenta, importe1, importe2, activo) VALUES (" & vUsu.Codigo & ","
    
    'Hacemos los saldos de actual
    '---------------------------
    'Comprobamos si leemos de hco o de normal
    vFecha1 = "25/" & M1 & "/" & A1
    vFecha2 = "25/" & M1 & "/" & A1
    EjerciciosCerrados = vFecha1 <= UtlimoFechaHCO
    InsertarSaldosCuentas True


    'Hacemos la de saldo anterior
    'Comprobamos si leemos de hco o de normal
    
    vFecha1 = "25/" & M1 & "/" & A1 - 1
    vFecha2 = vFecha1
    EjerciciosCerrados = vFecha1 <= UtlimoFechaHCO
    InsertarSaldosCuentas False
    
    'Salimos
    GeneraCtaExplComparativa = True
    Exit Function
EGeneraCtaExplComparativa:
    MuestraError Err.Number, "Genera Cta. Expl. Comparativa"
End Function




Private Sub InsertarSaldosCuentas(PrimeraVez As Boolean)
Dim QuitarPyG As Integer
Dim I As Integer
Dim RS As ADODB.Recordset

    'Si el mes contiene al mes de cierre hay que quitar perdidas y ganacias
    QuitarPyG = 0
    If Month(vFecha1) = Month(vParam.fechafin) Then
        If Not EjerciciosCerrados Then
            If vFecha1 < vParam.fechaini Then QuitarPyG = 1
        Else
            'En ejercicios cerrados siempre quitamos
            QuitarPyG = 1
        End If
    End If


    Aux = "hsaldos"
    If EjerciciosCerrados Then Aux = Aux & "1"
    H = Mid("__________", 1, A3 - 1)
    SQL = "Select cuentas.codmacta ,nommacta,sum(impmesde) as debe ,sum(impmesha) as haber from " & Aux & ",cuentas  "
    SQL = SQL & "  WHERE cuentas.codmacta=" & Aux & ".codmacta "
    SQL = SQL & " AND (cuentas.codmacta like '" & vParam.grupogto & H & "' or cuentas.codmacta like '" & vParam.grupovta & H & "')"
    If NumAsiento = 1 Then
        'Esta pidiendo acumulado
        'Para ello vamos a ir calculando la fecha de inicio de ejercicio
        'Primero vemos si es año siguiente
        If vFecha1 > vParam.fechafin Then
            vFecha2 = DateAdd("yyyy", 1, vParam.fechaini)
        Else
            I = -1
            Do
               I = I + 1
               vFecha2 = DateAdd("yyyy", -I, vParam.fechaini)
            Loop Until vFecha2 < vFecha1
        End If
        If Year(vFecha1) = Year(vFecha2) Then
            SQL = SQL & " AND anopsald = " & Year(vFecha1)
            SQL = SQL & " AND mespsald>=" & Month(vFecha2) & " AND mespsald <=" & Month(vFecha1)
        Else
            SQL = SQL & " AND (( anopsald = " & Year(vFecha2) & " AND mespsald>=" & Month(vFecha2) & ") OR "
            SQL = SQL & " ( anopsald =" & Year(vFecha1) & " AND mespsald <=" & Month(vFecha1) & "))"
        End If
    Else
        'Solo el mes actual
        SQL = SQL & " AND ( anopsald = " & Year(vFecha1) & " AND mespsald =" & Month(vFecha1) & ")"
    End If
    SQL = SQL & " group by codmacta"
    
    'Ahora recolocamos las fechas en el incio y fin de ejercicio
    If QuitarPyG = 1 Then
        'Solo si tenemos k calcualr perdidas y ganancias
        'manda vfecha=2    FALTA revisar que pasa con vfecha2, que no esta haciendolo bien
        vFecha1 = vFecha2
        I = Year(vFecha1)
        If Year(vParam.fechafin) <> Year(vParam.fechaini) Then I = I + 1
        vFecha2 = CDate(CStr(Format(vParam.fechafin, "dd/mm/") & I))
        'Preparamos el recordset para despues
        Set RT = New ADODB.Recordset
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'INSERT INTO zexplocomp (codusu, cta, cuenta, importe1, importe2, activo)
    While Not RS.EOF
        ImAcD = DBLet(RS!Debe, "N")
        ImAcH = DBLet(RS!Haber, "N")
        vCta = RS!codmacta
        
        If QuitarPyG = 1 Then
            CalcularImporteCierreAperturaPyG 1
        Else
            ImCierrD = 0
            ImCierrH = 0
        End If
        
        
        'Le quitamos el asiento de perdidas y ganancias
        ImAcD = ImAcD - ImCierrD
        ImAcH = ImAcH - ImCierrH
        
        If Mid(vCta, 1, 1) = vParam.grupovta Then
            Codigo = "0"
            ImAcD = ImAcH - ImAcD
        Else
            Codigo = "1"
            ImAcD = ImAcD - ImAcH
        End If
        H = TransformaComasPuntos(CStr(ImAcD))
        I = 0
        If PrimeraVez Then
            H = H & ",NULL"
        Else
            'Comprobamos si existe
            If ExisteEntradaExplotacionComarativa Then
                I = 1
            Else
                H = "NULL," & H
            End If
        End If

        'A la BD
        If I = 0 Then
            'NUEVO
            SQL = d & "'" & vCta & "','" & DevNombreSQL(CStr(RS!nommacta)) & "'," & H & "," & Codigo & ")"
        Else
            SQL = "UPDATE Usuarios.zexplocomp SET importe2= " & H
            SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cta='" & vCta & "'"
        End If
        Conn.Execute SQL
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Set RT = Nothing
End Sub






Private Function ExisteEntradaExplotacionComarativa() As Boolean
Dim RS As ADODB.Recordset
    
    ExisteEntradaExplotacionComarativa = False
    Set RS = New ADODB.Recordset
    SQL = "Select cta from Usuarios.zexplocomp where codusu=" & vUsu.Codigo & " AND cta='" & vCta & "'"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then ExisteEntradaExplotacionComarativa = True
    RS.Close
    Set RS = Nothing
End Function


Public Function ImprimirCtaExploComp() As Boolean
Dim Fin As Boolean

    ImprimirCtaExploComp = False
    'Eliminar temporal
    SQL = "Delete from Usuarios.zexplocompimpre where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    Set miRsAux = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    
    SQL = "Select * from Usuarios.zexplocomp where codusu = " & vUsu.Codigo
    SQL = SQL & " AND activo = "
    'Orden
    d = " ORDER BY Cta"
    
    'Activo
    miRsAux.Open SQL & "1" & d, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Pasivo
    RT.Open SQL & "0" & d, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Inicializar variables
    A1 = 0
    SQL = "INSERT INTO Usuarios.zexplocompimpre (codusu, codigo, cta, cuenta, importe1, importe2, 2cta, 2cuenta, 2importe1, 2importe2) VALUES (" & vUsu.Codigo & ","
    
    
    'ATENCION###
    'Los importes hay k cambiarlos de orden. Esto es  El k va en 1er lugar ira 2º y viceversa
    Do
        A1 = A1 + 1
        Fin = RT.EOF And miRsAux.EOF
        Aux = SQL & A1 & ","
        If Not Fin Then
            'Cargamos primero activo
            If miRsAux.EOF Then
               Codigo = "NULL,NULL,NULL,NULL"
            Else
                ImAcD = DBLet(miRsAux!Importe1, "N")
                ImAcH = DBLet(miRsAux!importe2, "N")
                'Antes de cambiarlo de orden
                'Codigo = "'" & miRsAux!Cta & "','" & miRsAux!Cuenta & "'" & ImporteASQL(ImAcD) & ImporteASQL(ImAcH)
                'Ahora. Lo mismo sera en la de abajo
                Codigo = "'" & miRsAux!Cta & "','" & DevNombreSQL(miRsAux!Cuenta) & "'" & ImporteASQL(ImAcH) & ImporteASQL(ImAcD)
                miRsAux.MoveNext
            End If
            Aux = Aux & Codigo & ","
            
            'Luego pasivo
            'Cargamos primero activo
            If RT.EOF Then
               Codigo = "NULL,NULL,NULL,NULL"
            Else
                ImAcD = DBLet(RT!Importe1, "N")
                ImAcH = DBLet(RT!importe2, "N")
                Codigo = "'" & RT!Cta & "','" & DevNombreSQL(RT!Cuenta) & "'" & ImporteASQL(ImAcH) & ImporteASQL(ImAcD)
                RT.MoveNext
            End If
            Aux = Aux & Codigo & ")"
            
            Conn.Execute Aux
        End If
    Loop Until Fin
    RT.Close
    miRsAux.Close
    Set RT = Nothing
    Set miRsAux = Nothing
    If A1 > 0 Then ImprimirCtaExploComp = True 'Si tiene datos
End Function


'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
'
'
'           Traspasos.  Persa y ACE
'
'
'
'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
'Nos vamos a permitir modificar el label PERSA  desde aqui

Public Function TraspasoPERSA(Actual As Boolean) As Boolean
Dim Debe As Boolean
Dim F1 As Date

    On Error GoTo ETraspasoPERSA
    TraspasoPERSA = False
    
    'Primer archivo
    Aux = App.path & "\conta" & vEmpresa.codempre & "_maestro"
    If Dir(Aux) <> "" Then Kill Aux
    A1 = FreeFile
    Open Aux For Output As #A1
    Set RT = New ADODB.Recordset
    Codigo = Space(30)
    frmListado.lblPersa2.Caption = "Cuentas:"
    frmListado.lblPersa2.Refresh
    RT.Open "Select codmacta,nommacta,apudirec from cuentas", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        frmListado.lblPersa.Caption = RT!codmacta
        frmListado.lblPersa.Refresh
        Aux = Mid(RT.Fields(0) & Codigo, 1, 10)
        Aux = Aux & Mid(RT.Fields(1) & Codigo, 1, 24)
        Aux = Aux & RT.Fields(2)
        Print #A1, Aux
        'Sig
        RT.MoveNext
    Wend
    RT.Close
    Close (A1)

    'Segundo archivo
    Aux = App.path & "\conta" & vEmpresa.codempre & "_histor"
    If Dir(Aux) <> "" Then Kill Aux
    A1 = FreeFile
    Open Aux For Output As #A1
    
    'SQL
    SQL = "Select codmacta,fechaent,numasien,timported,timporteh,ctacontr,numdiari,linliapu,idcontab"
    SQL = SQL & " FROM hlinapu WHERE "
    
    If Actual Then
        F1 = vParam.fechaini
    Else
        F1 = DateAdd("yyyy", 1, vParam.fechaini)
    End If
    SQL = SQL & " fechaent >= '" & Format(F1, FormatoFecha) & "'"
    
    
    If Actual Then
        F1 = vParam.fechafin
    Else
        F1 = DateAdd("yyyy", 1, vParam.fechafin)
    End If
    SQL = SQL & " AND fechaent <= '" & Format(F1, FormatoFecha) & "'"
    
    
    
    SQL = SQL & " ORDER BY numdiari,fechaent,numasien,linliapu"
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Textos
    frmListado.lblPersa2.Caption = "Asientos:"
    frmListado.lblPersa2.Refresh
    'Bucle principal
    While Not RT.EOF
        frmListado.lblPersa.Caption = RT!Numasien
        frmListado.lblPersa.Refresh
        
        ImpD = 0   'Sera el imponible
        NumRegElim = RT!Numasien
        vFecha1 = RT!fechaent
        
        
        H = Mid(RT!codmacta, 1, 3)
        
        'importe de la linea en ImCierrD
        If IsNull(RT!timported) Then
            Debe = False
            ImCierrD = RT!timporteH
        Else
            Debe = True
            ImCierrD = RT!timported
        End If
        
        
        
        'COMPROBAMOS CLIENTES
        If H = "477" And RT!idcontab = "FRACLI" Then
            A2 = RT!NumDiari
            DevuelveIVa NumRegElim, True
        End If
        
        'PROVEEDORES
        If H = "472" And RT!idcontab = "FRAPRO" Then
            A2 = RT!NumDiari
            DevuelveIVa NumRegElim, False
        End If
        
        'Grabamos la linea
        'Fijo
        Aux = "999999"
        Aux = Aux & "1"
        
        If Debe Then
            Aux = Aux & "51"
        Else
            Aux = Aux & "52"
        End If
        
        'Fecha
        Aux = Aux & Format(vFecha1, "ddmmyy")
        d = "   "
        If ImpD <> 0 Then  'IMPONIBLE
            If H = "477" Then
                d = "1" & Format(ImpH, "00")
            Else
                If H = "472" Then d = "2" & Format(ImpH, "00")
            End If
        End If
        Aux = Aux & d
        '3 posiciones en blanco
        Aux = Aux & "   "
        
        'Numasien
        Aux = Aux & Format(NumRegElim, "00000")
        
        If Debe Then
            M1 = 19
        Else
            M1 = 20
        End If
        Aux = Aux & M1
        
        Aux = Aux & Mid(RT!codmacta & Codigo, 1, 10)
        Aux = Aux & Mid(DBLet(RT!ctacontr) & Codigo, 1, 10)
        Aux = Aux & "      "  '6 blancos
        
        
        ImAcD = Abs(ImCierrD)
        ImCierrH = Abs(ImpD)
        d = Format(ImAcD, "000000000.00")
        M1 = InStr(1, d, ",")
        If M1 > 0 Then d = Mid(d, 1, M1 - 1) & Mid(d, M1 + 1)
        Aux = Aux & d
        
        If ImCierrD < 0 Then  'Este es el importe en la linea de apntes
            Aux = Aux & "-"
        Else
            Aux = Aux & " "
        End If
        
        
        'Imprimimos una A fija
        Aux = Aux & "A"
        
        d = Format(ImCierrH, "000000000.00")
        M1 = InStr(1, d, ",")
        If M1 > 0 Then d = Mid(d, 1, M1 - 1) & Mid(d, M1 + 1)
        Aux = Aux & d
        
        'Imprimimos la linea
        Print #A1, Aux
        
        'Siguiente
        RT.MoveNext
    Wend
    Close (A1)
    RT.Close
ETraspasoPERSA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description & vbCrLf & "Linea" & vbCrLf & Aux & vbCrLf & "Reg: " & NumRegElim, "Traspaso PERSA"
    Else
        TraspasoPERSA = True
    End If
    Set RT = Nothing
End Function




Private Sub DevuelveIVa(ByRef NumeroAsiento As Long, Clientes As Boolean)
Dim RS As Recordset
Dim Seguir As Boolean
Dim Tiene As Boolean

    ImpH = 0
    SQL = "Select * from cabfact"
    
    If Not Clientes Then
        SQL = SQL & "prov"
        d = "pr"
    Else
        d = "cl"
    End If
    SQL = SQL & " WHERE numdiari = " & A2 & " AND fechaent = '" & Format(vFecha1, FormatoFecha) & "' AND numasien =" & NumeroAsiento
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'Hemos encontrado la factura. Veamos si corresponde a la base 1,2, o tres
        'Base 1
        Tiene = False
        Seguir = True
        If RS.Fields("ti1fac" & d) = ImCierrD Then
                ImpD = RS.Fields("ba1fac" & d)
                ImpH = RS.Fields("pi1fac" & d)
                Tiene = True
                Seguir = False
        End If
        If Seguir Then
            Seguir = Not IsNull(RS.Fields("ti2fac" & d))
            If Seguir Then
                If ImCierrD = RS.Fields("ti2fac" & d) Then
                    ImpD = RS.Fields("ba2fac" & d)
                    ImpH = RS.Fields("pi2fac" & d)
                    Tiene = True
                    Seguir = False
                End If
            End If
        End If
                        
        If Seguir Then
            Seguir = Not IsNull(RS.Fields("ti3fac" & d))
            If Seguir Then
                If ImCierrD = RS.Fields("ti3fac" & d) Then
                    ImpH = RS.Fields("pi3fac" & d)
                    ImpD = RS.Fields("ba3fac" & d)
                    Tiene = True
                    Seguir = False
                End If
            End If
        End If
                    
    End If
    If Tiene Then
        If Clientes Then
            NumRegElim = RS!codfaccl
        Else
            NumRegElim = RS!NumRegis
        End If
        NumRegElim = NumRegElim Mod 100000
    End If
    RS.Close
    Set RS = Nothing
End Sub



'-----------------------------------------------------------------------
' Desde aqui generaremos el encabezado para las cartas.
'
'       Tipo 0: Cartas para el 347 del IVA
'       Tipo 1: Factura venta inmovilizado
'
'
Public Function CargaEncabezadoCarta(Opcion As Byte, Optional ByRef contacto As String)
    
    'El contatacto par el futuro
    Codigo = DevNombreSQL(contacto)
    
    'Borramos el anterior
    SQL = "DELETE FROm Usuarios.z347carta WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL

    'Cadena insert
    SQL = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos,"
    SQL = SQL & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, Asunto, Referencia, contacto) VALUES ("
    SQL = SQL & vUsu.Codigo
    
    'Los datos de la empresa son comunes
    'El resto de sql lo montamos en H
    H = ""
    MontaDatosEmpresa
    SQL = SQL & H
        
    'Por si da fallo, o para el inmovilizado
    H = ""
    If Opcion = 0 Then
        'Para el 347. Cogeremos los datos de un achivo
        Monta347
    Else
        MontaFacturaVenta
    End If
    SQL = SQL & H & ")"
    Conn.Execute SQL
End Function


Private Sub MontaDatosEmpresa()
    Set RT = New ADODB.Recordset
    RT.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If RT.EOF Then
        MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
        H = ",'','','','','',''"  '6 campos
    Else
        H = ",'" & DBLet(RT!nifempre) & "','" & vEmpresa.nomempre & "','"
        d = DBLet(RT!siglasvia) & " " & DBLet(RT!Direccion) & "  " & DBLet(RT!numero) & ", " & DBLet(RT!puerta)
        H = H & d & "','" & DBLet(RT!codpos) & "','" & DBLet(RT!poblacion) & "','" & DBLet(RT!provincia) & "'"
    End If
    RT.Close
    Set RT = Nothing
End Sub


Private Sub Monta347()
Dim Fin As Boolean
On Error GoTo Emon347
    M1 = FreeFile
    'Archivo con los datos
    vCta = App.path & "\txt347.dat"
    If Dir(vCta) = "" Then
        H = ",'','','','','','','','',''"   '8 pares
        H = H & ",'" & Codigo & "'"
        Exit Sub
    End If
    
    Open vCta For Input As #M1
    M2 = 0
    d = ""
    While Not Fin
        M2 = M2 + 1
        If M2 <= 9 Then
            'Las lineas van por pares, y hay 8 pares
            Line Input #M1, vCta
            Line Input #M1, vCta
            d = d & ",'" & vCta & "'"
            Fin = EOF(M1)
        Else
            Fin = True
        End If
    Wend
    
    Close #M1
    
    If M2 = 9 Then
        d = d & ",'" & Codigo & "'"
        H = d
    End If
    Exit Sub
Emon347:
    MuestraError Err.Number, "Fichero datos para el 347"
End Sub


Private Sub MontaFacturaVenta()

    'SABEMOS K en CadenaDesdeOtroForm estan los valores a guardar
    For A1 = 1 To 6
        H = H & ",'" & RecuperaValor(CadenaDesdeOtroForm, A1) & "'"
    Next A1
    
    For A1 = 7 To 10
        H = H & ",NULL"
    Next A1
    
    
End Sub













'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'
'
'       Comprobar Formula de Configuracion de balabnce
'
'
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
Public Function CompruebaFormulaConfigBalan(NumBalan As Integer, Formula As String) As String

    CompruebaFormulaConfigBalan = ""

    NumAsiento = NumBalan
    
    
    'Kitamos todos los esapacios en blanco
    Formula = Trim(Formula)
    Do
        A1 = InStr(1, Formula, " ")
        If A1 > 0 Then Formula = Mid(Formula, 1, A1 - 1) & Mid(Formula, A1 + 1)
    Loop Until A1 = 0
    
    
    
    'Comprobamos k los caracteres son correctos
    M2 = 1 'Bien
    For M1 = 1 To Len(Formula)
        d = Mid(Formula, M1, 1)
        Select Case d
        Case "0" To "9"
            'Son los numeros
            
        Case "+", "-"
            'El mas y el menos
            
        Case "A", "B"
        
        Case Else
            M2 = 0
            Exit For
        End Select
    Next M1
    
    If M2 = 0 Then
        CompruebaFormulaConfigBalan = "Caracteres incorrectos"
        Exit Function
    End If

    'para cada campo de la formula buscamos "+" o "-"
    M3 = 1
    Set RT = New ADODB.Recordset
    Do
        M1 = 0
        A1 = InStr(1, Formula, "-")
        A2 = InStr(1, Formula, "+")
        If A1 <> 0 Or A2 <> 0 Then
            If A1 = 0 Then A1 = 32000
            If A2 = 0 Then A2 = 32000
            
            If A1 > A2 Then
                M1 = A2
            Else
                M1 = A1
            End If
        Else
            If Formula <> "" Then
                M1 = Len(Formula) + 1
            End If
        End If
        
        If M1 > 0 Then
            d = Mid(Formula, 1, 1)     'activo pasivo, A o B
            H = Mid(Formula, 2, M1 - 2) 'Codigo
            
            If Not ExisteCodigoBalance Then
                CompruebaFormulaConfigBalan = "No existe codigo balance:  " & d & H
                Set RT = Nothing
                Exit Function
            End If
            
            Formula = Mid(Formula, M1 + 1)
        End If
        Loop Until M1 = 0
        
End Function



Private Function ExisteCodigoBalance() As Boolean
On Error GoTo EExisteCodigoBalance
    ExisteCodigoBalance = False
    SQL = "Select * from sperdid where numbalan=" & NumAsiento & " AND pasivo='" & d
    SQL = SQL & "' AND Codigo=" & H
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT!Codigo) Then
            ExisteCodigoBalance = True
        End If
    End If
    RT.Close
    Exit Function
EExisteCodigoBalance:
    MuestraError Err.Number
End Function




'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'           Impresion de balances configurables
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------


'--> Se puede mejorar puesto k algunas tablas temporales se cargan con datos k luego no
'    son utilizados
'



'Imprime el listado para que vena las cuentas que entran dentro de k punto etc etc
Public Function GeneraDatosBalanConfigImpresion(NumBalan As Integer)

        SQL = "Delete from "
        SQL = SQL & "tmpImpBalance where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        Conn.Execute "DELETE FROM Usuarios.ztmpimpbalan WHERE codusu = " & vUsu.Codigo
        
        Set RT = New ADODB.Recordset
        M3 = 1 'Sera el orden de insercion
        
        'Sera el numero de balance
        NumAsiento = NumBalan
        
        
        'Vemos si es de perdidas y ganacias
        SQL = DevuelveDesdeBD("perdidas", "sbalan", "numbalan", CStr(NumBalan), "N")
        EsBalancePerdidas_y_ganancias = (Val(SQL) = 1)
        
        
        
        'Vamos a utilizar la temporal de balances donde dejara los valores
        SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        
        Contabilidad = -1



       CargaArbol 0, 0, -1, "", Month(vParam.fechaini), Year(vParam.fechaini), Month(vParam.fechafin), Year(vParam.fechafin), "", True
 
 
 
 
 
         SQL = "Select * from "
        If Contabilidad > 0 Then SQL = SQL & "Conta" & Contabilidad & "."
        SQL = SQL & "sperdid where numbalan=" & NumBalan & " AND padre"
        Aux = "INSERT INTO Usuarios.ztmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,LibroCD,QueCuentas) VALUES (" & vUsu.Codigo
        Codigo = "Select importe1,importe2,quecuentas from "
        If Contabilidad > 0 Then Codigo = Codigo & "Conta" & Contabilidad & "."
        Codigo = Codigo & "tmpimpbalance where codusu=" & vUsu.Codigo & " AND pasivo='"
        M1 = 1
 

 
        CargaArbolImpresion -1, "", 1, False, False
End Function


Public Function GeneraDatosBalanceConfigurable(NumBalan As Integer, Mes1 As Integer, Anyo1 As Integer, Mes2 As Integer, Anyo2 As Integer, LibroCD As Boolean, vContabilidad As String)
Dim QuitarUno As Boolean
Dim EsPyGNoAbreviado As Boolean
Dim AuxPyG As String

    If vContabilidad = "-1" Then vContabilidad = "-1|"
    
    Set RT = New ADODB.Recordset
    
    
    'Metemos en las variable varconsolidado
    '       (0): fechas hco  . Posiciones fijas dd/mm/yyyy|
    '        1:  quitar 1     0|  o 1|
    '        2:  quitar 2 """"
            
    d = vContabilidad
    VarConsolidado(0) = "": VarConsolidado(1) = "": VarConsolidado(2) = ""
    
    While d <> ""
            'Vemos cual es
            M1 = InStr(1, d, "|")
            A1 = CInt(Mid(d, 1, M1 - 1))
            d = Mid(d, M1 + 1)
            
            Contabilidad = A1
            
            
            VFecha3 = CDate("01/12/1900")
            SQL = "Select max(fechaent) from "
            If Contabilidad > 0 Then SQL = SQL & "Conta" & Contabilidad & "."
            SQL = SQL & "hcabapu1"
            RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RT.EOF Then
                If Not IsNull(RT.Fields(0)) Then
                    VFecha3 = Format(RT.Fields(0), "dd/mm/yyyy")
                End If
            End If
            RT.Close
            
            
            VarConsolidado(0) = VarConsolidado(0) & VFecha3 & "|"
            

            '------------------------------------------------
            'Si hay k quitar saldos para cada ano
            'Comprobamos si hay que quitar el pyg y el cierre
            QuitarUno = False
            EjerciciosCerrados = (CDate("15/" & Mes1 & "/" & Anyo1) < VFecha3)
            'Si el mes contiene el cierre, entonces adelante
            If Month(vParam.fechafin) = Mes1 Then
                'Si estamos en ejerccicios cerrados seguro que hay asiento de cierre y p y g
                If EjerciciosCerrados Then
                    QuitarUno = True
                Else
                'Si no lo comprobamos. Concepto=960 y 980
                    QuitarUno = HayAsientoCierreBalances(CByte(Mes1), Anyo1)
                End If
            End If
            VarConsolidado(1) = VarConsolidado(1) & Abs(QuitarUno) & "|"
    
    
            'Si hay k quitar saldos para el segundo
            'Comprobamos si hay que quitar el pyg y el cierre
            QuitarUno = False
            If Mes2 > 0 Then
                EjerciciosCerrados = (CDate("15/" & Mes2 & "/" & Anyo2) < VFecha3)
                'Si el mes contiene el cierre, entonces adelante
                If Month(vParam.fechafin) = Mes2 Then
                    'Si estamos en ejerccicios cerrados seguro que hay asiento de cierre y p y g
                    If EjerciciosCerrados Then
                        QuitarUno = True
                    Else
                    'Si no lo comprobamos. Concepto=960 y 980
                        QuitarUno = HayAsientoCierreBalances(CByte(Mes2), Anyo2)
                    End If
                End If
            End If
            VarConsolidado(2) = VarConsolidado(2) & Abs(QuitarUno) & "|"
    
            
        Wend
    
    
    
    
    
    

        
        'Borramos las temporales
        SQL = "Delete from "
        SQL = SQL & "tmpImpBalance where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        Conn.Execute "DELETE FROM Usuarios.ztmpimpbalan WHERE codusu = " & vUsu.Codigo
        
        Set RT = New ADODB.Recordset
        M3 = 1 'Sera el orden de insercion
        
'
'        'Fecha del utlimo en hco
'        VFecha3 = CDate("01/12/1900")
'        SQL = "Select max(fechaent) from "
'        If Contabilidad > 0 Then SQL = SQL & "Conta" & Contabilidad & "."
'        SQL = SQL & "hcabapu1"
'        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If Not RT.EOF Then
'            If Not IsNull(RT.Fields(0)) Then
'                VFecha3 = Format(RT.Fields(0), "dd/mm/yyyy")
'            End If
'        End If
'        RT.Close
        
        
        
       
        'Sera el numero de balance
        NumAsiento = NumBalan
        
        
        'Vemos si es de perdidas y ganacias
        SQL = DevuelveDesdeBD("perdidas", "sbalan", "numbalan", CStr(NumBalan), "N")
        EsBalancePerdidas_y_ganancias = (Val(SQL) = 1)
        
        
        
        'Vamos a utilizar la temporal de balances donde dejara los valores
        SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        
        Contabilidad = -1
        
        
        
        
        
        CargaArbol 0, 0, -1, "", Mes1, Anyo1, Mes2, Anyo2, vContabilidad, False
    
            
        
        
        'Cuando termina de cargar el arbol vamos calculando las sumas
        SQL = "SELECT * FROM "
        'Al ponerle Conta?.   lo k damos a entender es k lee la configuracion de su PROIPA sperdi
        If Contabilidad > 0 Then SQL = SQL & "Conta" & Contabilidad & "."
        SQL = SQL & "sperdid where numbalan=" & NumBalan & " AND tipo = 1"
        SQL = SQL & " ORDER BY orden"
        
        'Modificacion 12 Febrero 2004
        '----------------------------
        '  A igual numero de orden, ordena por creacion entonces da la casualidad de que
        ' muestra hace primero el BV del pasivo k el AiV
        SQL = SQL & ",Pasivo"
        
        
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RT.EOF
            CalculaSuma DBLet(RT!Formula), Val(RT!A_Cero) <> 0
            d = TransformaComasPuntos(CStr(ImpD))
            'UPDATEAMOS
            
            SQL = "UPDATE "
            If Contabilidad > 0 Then SQL = SQL & "Conta" & Contabilidad & "."
            SQL = SQL & "tmpImpBalance SET importe1 =" & d
            If M2 > 0 Then
                H = TransformaComasPuntos(CStr(ImpH))
                SQL = SQL & ",importe2 = " & H
            End If
            SQL = SQL & " where codusu = " & vUsu.Codigo
            SQL = SQL & " AND Pasivo='" & RT!Pasivo & "' AND Codigo=" & RT!Codigo
            Conn.Execute SQL
            RT.MoveNext
        Wend
        RT.Close
    
    
    'QUITAR. Ya ha calculado los saldos todos juntitos
    '
    'If Contabilidad < 0 Then
        'Una vez todos los importes y demas vamos a generar los datos en impresion
        'Lo unico a tener en cuenta es k en las formulas si es menor k 0 no se imprime
        '-----------------------------------------------------------------------------
        
        SQL = "Select * from "
        If Contabilidad > 0 Then SQL = SQL & "Conta" & Contabilidad & "."
        SQL = SQL & "sperdid where numbalan=" & NumBalan & " AND padre"
        Aux = "INSERT INTO Usuarios.ztmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,LibroCD,QueCuentas) VALUES (" & vUsu.Codigo
        Codigo = "Select importe1,importe2,quecuentas from "
        If Contabilidad > 0 Then Codigo = Codigo & "Conta" & Contabilidad & "."
        Codigo = Codigo & "tmpimpbalance where codusu=" & vUsu.Codigo & " AND pasivo='"
        M1 = 1
        
        
        
        
        'Para que no imprima en la primera columna lo que seria la suma de "los hijos"
        EsPyGNoAbreviado = False
        If EsBalancePerdidas_y_ganancias Then

            AuxPyG = DevuelveDesdeBD("nombalan", "sbalan", "numbalan", CStr(NumBalan), "N")
            AuxPyG = UCase(AuxPyG)
            If InStr(1, AuxPyG, "ABREV") = 0 Then EsPyGNoAbreviado = True
        End If
        
        
        CargaArbolImpresion -1, "", 1, LibroCD, EsPyGNoAbreviado
        
    'End If
    
    Set RT = Nothing
End Function

'Public Sub ImprimeBalPerConsolidado2(NumBalan As Integer, LibroCD As Boolean)
''Public Sub imprimiBalPerConsolidado(NumBalan As Integer, Mes1 As Integer, Anyo1 As Integer, Mes2 As Integer, Anyo2 As Integer, LibroCD As Boolean, vContabilidad As Integer)
'        SQL = "Select * from sperdid where numbalan=" & NumBalan & " AND padre"
'        Aux = "INSERT INTO Usuarios.ztmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,LibroCD) VALUES (" & vUsu.Codigo
'        Codigo = "Select importe1,importe2 from "
'        'If Contabilidad > 0 Then Codigo = Codigo & "Conta" & Contabilidad & "."
'        Codigo = Codigo & "tmpimpbalance where codusu=" & vUsu.Codigo & " AND pasivo='"
'        M1 = 1
'
'        Set RT = New ADODB.Recordset
'        CargaArbolImpresion -1, "", 1, LibroCD
'End Sub

'Private Sub CargaArbol(ByRef vImporte As Currency, ByRef vimporte2 As Currency, Padre As Integer, Pasivo As String, ByRef Mes1 As Integer, ByRef Anyo1 As Integer, ByRef Mes2 As Integer, ByRef Anyo2 As Integer, Quitar1 As Boolean, QUitar2 As Boolean, Contabilidades As String)
Private Sub CargaArbol(ByRef vImporte As Currency, ByRef vimporte2 As Currency, Padre As Integer, Pasivo As String, ByRef Mes1 As Integer, ByRef Anyo1 As Integer, ByRef Mes2 As Integer, ByRef Anyo2 As Integer, ByRef Contabilidades As String, EsListado As Boolean)
Dim RS As ADODB.Recordset
Dim nodImporte As Currency
Dim MiAux As String
Dim OtroImporte As Currency
Dim OtroImporte2 As Currency
Dim QueCuentas As String

'Nuevo PGC.  Puede ser que UN nodo raiz sea la SUMA , con lo cual pasan dos cosas:
'   .- 1: Puede que tenga nodos colgando, que habra que calcular
'           el importe que habra que pintarle sera LA de la formula
'   .- 2:  Un nodo raiz No es una formula.
'           el importe que pinmtara sera el de la suma
Dim Tipo As Integer

    If Padre < 0 Then
        MiAux = " is null" 'NODO RAIZ
    Else
        MiAux = " = " & Padre & " AND Pasivo = '" & Pasivo & "'"
    End If
    
    
    MiAux = "sperdid where numbalan=" & NumAsiento & " AND padre" & MiAux
    'AQUI No pondremos lo de contabilidad, pq todos
    'repito TODOS cojeran los datos del balance configurable de la empresa en la que
    'estoy
    'If Contabilidad > 0 Then MiAux = "Conta" & Contabilidad & "." & MiAux
    MiAux = "Select * from " & MiAux
    
    

    Set RS = New ADODB.Recordset
    RS.Open MiAux & " ORDER By Orden", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
    
       
    
        OtroImporte = 0
        OtroImporte2 = 0
        'If RS!Tipo = 1 Then Stop
        If vParam.NuevoPlanContable Then
            Tipo = 0  'ASi seguro que entro
        Else
            Tipo = RS!Tipo
        End If
        If Tipo = 0 Then
            If RS!tienenctas = 1 Then
                QueCuentas = ""
                If EsListado Then
                    'LISTADITO
                    
                    QueCuentas = PonerCuentasBalances(RS!Pasivo, RS!Codigo)
                Else
                    'IMPORTES
                    OtroImporte = CalculaImporteCtas(RS!Pasivo, RS!Codigo, Mes1, Anyo1, True, Contabilidades)
                
                    If Mes2 > 0 Then OtroImporte2 = CalculaImporteCtas(RS!Pasivo, RS!Codigo, Mes2, Anyo2, False, Contabilidades)
                End If
            Else
                CargaArbol OtroImporte, OtroImporte2, RS!Codigo, RS!Pasivo, Mes1, Anyo1, Mes2, Anyo2, Contabilidades, EsListado
                QueCuentas = ""
                
            End If
        Else
            QueCuentas = ""
        End If
        
        
        vImporte = vImporte + OtroImporte
        vimporte2 = vimporte2 + OtroImporte2
        
        'Insertamos la linea
        'en aux hay
        'codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,orden) VALUES (1
        MiAux = "'" & RS!Pasivo & "'," & RS!Codigo & ",'" & RS!texlinea & "','"
        d = TransformaComasPuntos(CStr(OtroImporte))
        MiAux = MiAux & RS!deslinea & "'," & d & ","
        d = TransformaComasPuntos(CStr(OtroImporte2))
        MiAux = MiAux & d & ",0," & M3 & ",'" & QueCuentas & "')"
        Aux = "INSERT INTO "
        'If Contabilidad > 0 Then Aux = Aux & "Conta" & Contabilidad & "."
        Aux = Aux & "tmpimpbalance (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,orden,quecuentas) VALUES (" & vUsu.Codigo & ","
        MiAux = Aux & MiAux
        Conn.Execute MiAux
    
        M3 = M3 + 1
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    
End Sub

''''''''' 'Borramos los temporales
''''''''        SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
''''''''        Conn.Execute SQL
''''''''
''''''''        I = 0
''''''''        Rs.MoveFirst
''''''''        While Not Rs.EOF
''''''''            CargaBalance Rs.Fields(0), Rs.Fields(1), Apertura, cmbFecha(IndiceCombo).ListIndex + 1, cmbFecha(IndiceCombo + 1).ListIndex + 1, CInt(txtAno(IndiceCombo).Text), CInt(txtAno(IndiceCombo + 1).Text), MesInicioContieneFechaInicioEjercicio, FechaIncioEjercicio, FechaFinEjercicio, EjerciciosCerrados, QuitarSaldos, vConta
''''''''            'Progress
''''''''            PB.Value = Round((I / Cont), 3) * 1000
''''''''            PB.Refresh
''''''''            'Siguiente cta
''''''''            I = I + 1
''''''''            Rs.MoveNext
''''''''        Wend
''''''''    End If
''''''''    Rs.Close


Private Function CalculaImporteCtas(Pasivo As String, Codigo As Integer, ByRef mess1 As Integer, ByRef anyos1 As Integer, Año1_o2 As Boolean, ByRef Contabilidades As String) As Currency
Dim RT As ADODB.Recordset
Dim X As Integer
Dim Y As Integer
Dim vI1 As Currency
Dim vI2 As Currency
Dim QuitarUno As Boolean
Dim Contador As Integer
Dim ContaX As String

        If Contabilidades = "" Then
            ContaX = "-1|"
        Else
            ContaX = Contabilidades
        End If
        Set RT = New ADODB.Recordset
        CalculaImporteCtas = 0
        vI2 = 0
        Contador = 0
        'para cada contbilida
        While ContaX <> ""
            
            'Vemos cual ContaX
            X = InStr(1, ContaX, "|")
            Y = CInt(Mid(ContaX, 1, X - 1))
            ContaX = Mid(ContaX, X + 1)
            
            Contabilidad = Y
            
            'Fecha del utlimo en hco
            VFecha3 = CDate(Mid(VarConsolidado(0), (11 * Contador) + 1, 10))
                
            'Quitar1
            If Año1_o2 Then
                QuitarUno = (Mid(VarConsolidado(1), (Contador * 2) + 1, 1) = 1)
            Else
                QuitarUno = Mid(VarConsolidado(2), (Contador * 2) + 1, 1)
            End If
            
            EjerciciosCerrados = (CDate("15/" & mess1 & "/" & anyos1) < VFecha3)
            vI1 = CalculaImporteCtas1Contabilidad(Pasivo, Codigo, mess1, anyos1, QuitarUno)
        
'        '------------------------------------------------
'        'Si hay k quitar saldos para segundo
'        'Comprobamos si hay que quitar el pyg y el cierre
'        QUitarDos = False
'        'Si hay k hacer los dos
'        If Mess2 > 0 Then
'            EjerciciosCerrados = (CDate("15/" & Mes2 & "/" & Anyo2) < VFecha3)
'            'Si el mes contiene el cierre, entonces adelante
'            If Month(vParam.fechafin) = Mes2 Then
'                'Si estamos en ejerccicios cerrados seguro que hay asiento de cierre y p y g
'                If EjerciciosCerrados Then
'                    QUitarDos = True
'                Else
'                    'Si no lo comprobamos. Concepto=960 y 980
'                    QUitarDos = HayAsientoCierreBalances(CByte(Mes2), Anyo2)
'                End If
'            End If
'        End If
        
        
        vI2 = vI2 + vI1

        Contador = Contador + 1
    Wend
    Set RT = Nothing
    Contabilidad = -1
    CalculaImporteCtas = vI2
End Function






Private Function CalculaImporteCtas1Contabilidad(Pasivo As String, Codigo As Integer, ByRef mess1 As Integer, ByRef anyos1 As Integer, QuitarSaldos As Boolean) As Currency
Dim RC As ADODB.Recordset
Dim F1 As Date
Dim F2 As Date
Dim I1 As Currency
Dim B1 As Byte

    Set RC = New ADODB.Recordset
        
    'Vamos a calcular el importe para cada cuenta, para cada contbiliadad
        
    vCta = "SELECT * from "
   ' If Contabilidad > 0 Then vCta = vCta & "Conta" & Contabilidad & "."
    vCta = vCta & "sperdi2 WHERE pasivo ='" & Pasivo & "' AND codigo = " & Codigo & " AND numbalan = " & NumAsiento
    RC.Open vCta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
  
    'Primer mes
    If mess1 < Month(vParam.fechaini) Then
        F1 = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & anyos1 - 1)
    Else
        F1 = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & anyos1)
    End If

    F2 = CDate(DiasMes(CInt(mess1), anyos1) & "/" & mess1 & "/" & anyos1)
    I1 = 0
    EjerciciosCerrados = F1 < VFecha3
    If QuitarSaldos Then
        B1 = 1  'Ambos, pyg y cierre
    Else
        B1 = 0  'Si no tengo que quitar saldos pòngo un cero
    End If
    While Not RC.EOF
                                                                                                                            'QUITAMOS si fecha incio menor k ultima fecha traspasada
        CargaBalanceNuevo RC!codmacta, "", False, Month(F1), mess1, Year(F1), anyos1, False, F1, F2, EjerciciosCerrados, B1, Contabilidad, True, True, False

        'Si el balance es el de situacion numbalan=2
        ' y la cuenta es la de perdidas y ganacias
        
        'If (NumAsiento = 2 Or NumAsiento = 4) And RC!codmacta = Mid(vParam.ctaperga, 1, 3) Then
        If Not EsBalancePerdidas_y_ganancias And RC!codmacta = Mid(vParam.ctaperga, 1, 3) Then
        
                ImCierrH = 0
                ImCierrD = 0

                'Este trozo es manual. Solo, SOLO, sirve para la cta pyg del 129
                ObtenerPerdidasyGanancias EjerciciosCerrados, F1, F2, 3    'SOLO LE QUITO EL CIERRE


                'Como viene de un apunte de regularizacion entonces lo del haber va al debe
                'En impcierr tenemos los importes
                ImpH = ImpH - ImCierrH
                ImpD = ImpD - ImCierrD
                
                
'                ImpH = ImCierrD
'                ImpD = ImCierrH
                
        End If

        
        'NUEVO NUEVO
        '-----------
        
        'If RC!Codmacta = "640" Then Stop
        If vParam.NuevoPlanContable Then
        
            If EsBalancePerdidas_y_ganancias Then
                'If Mid(RC!Codmacta, 1, 1) = "6" Then
                '    ImpH = ImpD - ImpH
                'Else
                    ImpH = ImpH - ImpD
                'End If
            Else
                ImpH = ImpD - ImpH   'Como estaba
                If Pasivo = "B" Then ImpH = -1 * ImpH
            End If
        Else
            ImpH = ImpD - ImpH
            If Pasivo = "B" Then ImpH = -1 * ImpH
        End If


        If RC!tipsaldo <> "S" Then
            Debug.Print RC!codmacta
            'Stop
        End If
        
        Select Case RC!tipsaldo
        Case "D"
            'Y la cuenta es de haber pongo a 0
            If ImpH < 0 Then ImpH = 0
        Case "H"
           If ImpH < 0 Then ImpH = 0
        Case "S"

 
        End Select

        
        If RC!Resta = 1 Then ImpH = ImpH * -1
        I1 = I1 + ImpH
        
        
        'Siguiente
        RC.MoveNext
    Wend
    RC.Close
    CalculaImporteCtas1Contabilidad = I1
    Set RC = Nothing
End Function

Private Function PonerCuentasBalances(Pasivo As String, Codigo As Integer) As String
Dim RC As ADODB.Recordset


    Set RC = New ADODB.Recordset
        
    'Vamos a calcular el importe para cada cuenta, para cada contbiliadad
        
    vCta = "SELECT * from "
   ' If Contabilidad > 0 Then vCta = vCta & "Conta" & Contabilidad & "."
    vCta = vCta & "sperdi2 WHERE pasivo ='" & Pasivo & "' AND codigo = " & Codigo & " AND numbalan = " & NumAsiento
    RC.Open vCta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
  
    vCta = ""
    While Not RC.EOF
        If vCta <> "" Then vCta = vCta & ","
        If RC!Resta = 1 Then vCta = vCta & "-"
        vCta = vCta & RC!codmacta & " "

        'Siguiente
        RC.MoveNext
    Wend
    RC.Close
    PonerCuentasBalances = CStr(vCta)
    Set RC = Nothing
End Function





Private Sub CalculaSuma(CadenaSuma As String, A_Cero As Boolean)
Dim RA As ADODB.Recordset
    CadenaSuma = Trim(CadenaSuma)
    ImpD = 0
    ImpH = 0
    If CadenaSuma = "" Then Exit Sub
        
    
    'Quitamos todos los blancos
    Do
        A1 = InStr(1, CadenaSuma, " ")
        If A1 > 0 Then CadenaSuma = Mid(CadenaSuma, 1, A1 - 1) & Mid(CadenaSuma, A1 + 1)
    Loop Until A1 = 0
                  
    Aux = Mid(CadenaSuma, 1, 1)
    If Aux <> "+" And Aux <> "-" Then CadenaSuma = "+" & CadenaSuma
    
    Set RA = New ADODB.Recordset
    'Dejamos medio montado el sql
    ' "INSERT INTO tmpimpbalance (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,orden) VALUES (" & vUsu.Codigo & ","
    SQL = "Select importe1,importe2 from "
    If Contabilidad > 0 Then SQL = SQL & "Conta" & Contabilidad & "."
    SQL = SQL & "tmpimpbalance where codusu =" & vUsu.Codigo

    
    'Iremos deglosando cadenasuma
    ImAcD = 0
    ImAcH = 0
    Do
        'Empezamos en dos, pq lo primero es siempre un mas o un menos
        A1 = InStr(2, CadenaSuma, "+")
        A2 = InStr(2, CadenaSuma, "-")
        If A1 = 0 And A2 = 0 Then
            'Ya no hay mas para procesar
            A3 = 0
            Aux = CadenaSuma
            CadenaSuma = ""
            M1 = 1
        Else
            If A1 = 0 Then A1 = 32000
            If A2 = 0 Then A2 = 32000
            If A1 > A2 Then
                A3 = A2
            Else
                A3 = A1
            End If
            Aux = Mid(CadenaSuma, 1, A3 - 1)
            CadenaSuma = Mid(CadenaSuma, A3)
        End If
        
        'El signo
        If Mid(Aux, 1, 1) = "-" Then
            M1 = -1
        Else
            M1 = 1
        End If
        
        'La letra del pasivo / activo
        vCta = " AND Pasivo = '" & Mid(Aux, 2, 1)
        
        'El codigo del campo
        d = Mid(Aux, 3)
        Codigo = "-2"
        If d <> "" Then
            If IsNumeric(d) Then Codigo = d
        End If
        
        'SQL para la BD
        vCta = vCta & "' AND codigo =" & d
        
        RA.Open SQL & vCta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ImCierrD = 0
        ImCierrH = 0
        If Not RA.EOF Then
            If Not IsNull(RA.Fields(0)) Then ImCierrD = RA.Fields(0)
            If Not IsNull(RA.Fields(1)) Then ImCierrH = RA.Fields(1)
        End If
        RA.Close
        ImAcD = ImAcD + (M1 * ImCierrD)
        ImAcH = ImAcH + (M1 * ImCierrH)
    Loop Until CadenaSuma = ""
    'Ponemos a cero si asi lo dice la funcion
    If A_Cero Then
        If ImAcD < 0 Then ImAcD = 0
        If ImAcH < 0 Then ImAcH = 0
    End If
    ImpD = ImAcD
    ImpH = ImAcH
    Set RA = Nothing
    
End Sub



Private Sub CargaArbolImpresion(Padre As Integer, Pasivo As String, Nivel As Byte, vLibroCD As Boolean, EsBalancePyGNOabreviado As Boolean)
Dim RS As ADODB.Recordset
Dim MiAux As String
Dim TieneHijos As Boolean
Dim QueCuentas As String

    
    If Padre < 0 Then
        MiAux = " is null" 'NODO RAIZ
    Else
        MiAux = " = " & Padre & " AND Pasivo = '" & Pasivo & "'"
    End If
    MiAux = SQL & MiAux & " ORDER By Pasivo, Orden"
  
    Set RS = New ADODB.Recordset
    RS.Open MiAux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        
        
        If vParam.NuevoPlanContable Then
            'Nueva contabilidad
            TieneHijos = True   'Existe la posibilidad (de echo lo hace, que teniendo hijos sea una FORMULA
        
        
        
        Else
            'Antiguo plan
            TieneHijos = False
            If RS!Tipo = 0 Then
                If RS!tienenctas = 0 Then TieneHijos = True
            End If
        End If
        
        'Obtenmos el importe
        MiAux = Codigo & RS!Pasivo & "' AND codigo =" & RS!Codigo
        RT.Open MiAux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ImpD = 0: ImpH = 0: QueCuentas = ""
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpD = RT.Fields(0)
            If Not IsNull(RT.Fields(1)) Then ImpH = RT.Fields(1)
            QueCuentas = DBLet(RT.Fields(2))
        End If
        RT.Close
        
        If vParam.NuevoPlanContable And Padre < 0 And RS!Tipo = 0 Then
            'Marzo 2011
            'Para los pyG NO abreviados
            If EsBalancePyGNOabreviado Then
            'If RS!NumBalan = 2 Then
                'FALTA### comprobar que es para todos.    De momento para PyG normal
                'Aqui no pintaremos el resultado de la suma de subnodos si el padre es null
                ' y NO es una formula
                ImpD = 0: ImpH = 0:
            End If
        End If
        
        
        
        'Si no se pinta si el resultado es negativo entonces entonces
        If RS!Pintar = 0 Then   'PINTAR: SI siempre NO. Ngativos no
            If ImpD < 0 Then ImpD = 0
            If ImpH < 0 Then ImpH = 0
        End If
        
        
            'Insertamos la linea
            'AUX tiene:
            'INSERT INTO usuari.ztmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita) VALUES (" & vUsu.Codigo
            MiAux = Aux & ",'" & RS!Pasivo & "'," & M1 & ",'" & DBLet(RS!texlinea) & "','"
            'El sangrado para el texto
            MiAux = MiAux & Space((Val(Nivel) - 1) * 4)  'SANGRIA TABULADO PARRAFO
            MiAux = MiAux & RS!deslinea & "'" & ImporteASQL(ImpD) & ImporteASQL(ImpH)
            MiAux = MiAux & "," & RS!negrita & ","
            If vLibroCD Then
                MiAux = MiAux & "'" & DBLet(RS!LibroCD) & "'"
            Else
                MiAux = MiAux & "NULL"
            End If
            MiAux = MiAux & ",'" & QueCuentas & "')"
            'MiAux = MiAux & ")"
            Conn.Execute MiAux
    
    
        M1 = M1 + 1
        
        'Ahora,si tiene hijos cargamos el subarbol
        
        If TieneHijos Then CargaArbolImpresion RS!Codigo, RS!Pasivo, Nivel + 1, vLibroCD, EsBalancePyGNOabreviado
        
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub




Private Function HayAsientoCierreBalances(Mes As Byte, Anyo As Integer) As Boolean
Dim C As String
Dim RS As Recordset
    HayAsientoCierreBalances = False
    'C = "01/" & CStr(Me.cmbFecha(1).ListIndex + 1) & "/" & txtAno(1).Text
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierreBalances = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            Set RS = New ADODB.Recordset
            C = "Select count(*) from "
            If Contabilidad > 0 Then C = C & "Conta" & Contabilidad & "."
            C = C & "hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then
                    If RS.Fields(0) > 0 Then HayAsientoCierreBalances = True
                End If
            End If
            RS.Close
            Set RS = Nothing
        End If
    End If
End Function






'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'           GENERACION DATOS MEMORIA
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Public Function CalcularImporteCta(FechaI As Date, Cuenta As String, EnCerrados As Boolean, Opcion As Byte) As Currency
Dim SQL As String
Dim fechafin As Date
    On Error GoTo ECalcularImporteCta
    CalcularImporteCta = 0
    
    'Opciones
    ' 1.- Actual   2.- Anterior  3.- Asiento apertura
    Select Case Opcion
    Case 3
        'ASto apertura
        SQL = "Select sum(timported),sum(timporteH) from hlinapu"
        If EnCerrados Then SQL = SQL & "1"
        SQL = SQL & " WHERE fechaent='" & Format(FechaI, FormatoFecha) & "' AND codconce=970"
        SQL = SQL & " AND codmacta like '" & Cuenta & "%';"
        ImCierrD = 0
        ImCierrH = 0
    Case 1, 2
        'Lo k cambiara sera la fechaI y encerrados
        SQL = "Select sum(impmesde),sum(impmesHa) from hsaldos"
        If EnCerrados Then SQL = SQL & "1"
        SQL = SQL & " WHERE codmacta = " & Cuenta
        SQL = SQL & " AND anopsald = " & Year(FechaI)
        
        fechafin = DateAdd("yyyy", 1, FechaI)
        fechafin = DateAdd("d", -1, fechafin)
        vCta = Cuenta
        Contabilidad = vEmpresa.codempre
        ObtenerPerdidasyGanancias EnCerrados, FechaI, fechafin, 1   'El 1 Indica  los dos pyg y cerrados
    End Select
        
        
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        'Debe
        If Not IsNull(RT.Fields(0)) Then
            ImpD = RT.Fields(0)
        Else
            ImpD = 0
        End If
        'Haber
        If Not IsNull(RT.Fields(1)) Then
            ImpH = RT.Fields(1)
        Else
            ImpH = 0
        End If
        ImpD = ImpD - ImCierrD
        ImpH = ImpH - ImCierrH
        CalcularImporteCta = ImpD - ImpH
    End If
    RT.Close
    Exit Function
ECalcularImporteCta:
    MuestraError Err.Number, "Calcular Importe Cta. " & vbCrLf & Err.Description
End Function





'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
'
'
'           Traspasos.  ACE
'
'
'
'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////


Public Sub FijarValoresACE(BMCode As String, Mes As Integer, fechaini As Date, fechafin As Date, Mensual As Boolean, Digitos As Integer)

    vFecha1 = fechaini
    vFecha2 = fechafin
    'Mes calculo
    M1 = Mes
    'Mensual o acumulado
    M3 = Abs(Mensual)
    'Digitos
    A3 = Digitos
    'BMCode
    H = BMCode
    
    'La primeravez todo seran inserts
    EjerciciosCerrados = True
    'Contador
    A2 = 1
End Sub



Public Function GenerarACE2(Empresa As Integer) As Boolean

    On Error GoTo EGenerarACE
    GenerarACE2 = False
    
    'SOLO para años naturales.
    '_______________________
    
    
    'Deberiamos comprobar k para esta empresa existe el  nivel de digitpos marcado en A3
    Set RT = New ADODB.Recordset
    
    
    '--------------------------------------------------------------------------------------------
    'Para el mes ACTUAL
    SQL = "select sum(impmesde) as debe,sum(impmesha)as haber,codmacta from conta" & CStr(Empresa)
    SQL = SQL & ".hsaldos WHERE"
    SQL = SQL & " codmacta like '" & Mid("__________", 1, A3) & "' AND "
    SQL = SQL & " anopsald =" & Year(vFecha1) & " AND mespsald "
    
    If M3 > 0 Then
        'SOLO MENSUAL
        Aux = ""
    Else
        'ACUMULADO
        Aux = "<"
    End If
    SQL = SQL & Aux & "= " & M1
    Aux = " GROUP BY codmacta"
        
    DatosAce
    GenerarACE2 = True
    
    
    
    
    
    
    
    
EGenerarACE:
    If Err.Number <> 0 Then MuestraError Err.Number, "Genera ACE"
    Set RT = Nothing
End Function



Private Sub DatosAce()
Dim RS As ADODB.Recordset
    
    Set RS = New ADODB.Recordset
    RS.Open SQL & Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        vCta = RS!codmacta
        ImpD = RS!Debe
        ImpH = RS!Haber
        ImpD = ImpD - ImpH
        'Si es true significa k es la primera
        If Not EjerciciosCerrados Then
            NumAsiento = ExisteEntradaAce
        Else
            NumAsiento = 0
            ImCierrD = 0
        End If
        ImpD = ImpD + ImCierrD
        d = TransformaComasPuntos(CStr(ImpD))
        If NumAsiento > 0 Then
            SQL = "UPDATE Usuarios.ztmppresu1 SET importe=" & d
            SQL = SQL & " WHERE codusu = " & vUsu.Codigo & " AND Codigo =" & NumAsiento
            
        Else
            SQL = "INSERT INTO Usuarios.ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (" & vUsu.Codigo & ","
            SQL = SQL & A2 & ",'" & vCta & "',NULL,0,0," & d & ")"
            A2 = A2 + 1
        End If
        Conn.Execute SQL
        
        'Sig
        RS.MoveNext
    Wend
    'Ya hemos pasado una vez por aqui. Ponemos a false la señal
    EjerciciosCerrados = False
    RS.Close
    Set RS = Nothing
End Sub


Private Function ExisteEntradaAce() As Long
Dim C As String
    
    ExisteEntradaAce = -1
    ImCierrD = 0
    C = "Select importe,codigo from Usuarios.ztmppresu1 where codusu= " & vUsu.Codigo
    C = C & " AND cta = '" & vCta & "';"
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        ExisteEntradaAce = RT!Codigo
        If Not IsNull(RT!Importe) Then ImCierrD = RT!Importe
    End If
    RT.Close
End Function




Public Function GeneraFicheroAce() As Boolean

    On Error GoTo EGeneraFicheroAce
    GeneraFicheroAce = False
    Aux = App.path & "\trasace.txt"
    If Dir(Aux) <> "" Then Kill Aux
    
    SQL = "Select cta,importe from Usuarios.ztmppresu1 where codusu =" & vUsu.Codigo & "  ORDER BY cta"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
    
        A1 = FreeFile
        Open Aux For Output As #A1
        
        'Primer registro
        SQL = "AA" & Mid(H & "      ", 1, 7) 'BMCODE
        SQL = SQL & Format(M1, "00") & Format(vFecha1, "yy")
        SQL = SQL & " ," & Format(A3, "00") & "EUR"
        Print #A1, SQL
        
        Aux = Space(20)
        A2 = 0
        'Registros
        While Not RT.EOF
            
            'La cuenta   NUNCA PODRAN PONER 10 digitos
            SQL = RT!Cta
            If A2 = 0 Then
                A2 = Len(SQL)
                If A2 = 10 Then MsgBox "Podrian solaparse cifras.", vbExclamation
                A2 = 20 - A2
            End If
            
            d = Format(RT!Importe, "0.00")
            d = Aux & d
            d = Right(d, A2)
            SQL = SQL & d
            If RT!Importe <> 0 Then Print #A1, SQL
            
            
            'Sig
            RT.MoveNext
        Wend
        
        'Ultimo registro
        SQL = "XX" & Mid(H & "      ", 1, 7) 'BMCODE
        SQL = SQL & Format(M1, "00") & Format(vFecha1, "yy")
        SQL = SQL & "       "
        Print #A1, SQL
        Close (A1)
        
        GeneraFicheroAce = True
        
        
    End If
    RT.Close
    Exit Function
EGeneraFicheroAce:
    MuestraError Err.Number, "Generando fichero ACE"
End Function





Public Function ImpirmirListadoInmovilizados(ByRef miSQL As String) As Boolean
    
On Error GoTo EImpirmirListadoInmovilizados
    ImpirmirListadoInmovilizados = False
    Conn.Execute "Delete from Usuarios.zfichainmo WHERE codusu =" & vUsu.Codigo
    
    Set RT = New ADODB.Recordset
    
    d = "Select * from sinmov "
    Aux = ""
    If miSQL <> "" Then
        M1 = InStr(1, miSQL, "WHERE")
        If M1 > 0 Then
            Aux = Mid(miSQL, M1)
            M1 = InStr(1, Aux, "ORDER BY")
            If M1 > 0 Then Aux = Mid(Aux, 1, M1 - 1)
            d = d & Aux
        End If
    End If
    d = d & " order by codinmov"
    M1 = 0
    RT.Open d, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Insertamos los elementos
    Aux = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, fechaamor, Importe, porcenta) VALUES (" & vUsu.Codigo & ","
    While Not RT.EOF
        M1 = M1 + 1
        d = M1 & "," & RT!Codinmov & ",'" & DevNombreSQL(RT!nominmov) & "','" & Format(RT!fechaadq, "dd/mm/yyyy")
        
        ImpD = RT!valoradq
        d = d & "'," & TransformaComasPuntos(CStr(ImpD)) & ",'" & RT!codmact1
                
        ImpD = RT!amortacu
        d = d & "'," & TransformaComasPuntos(CStr(ImpD))
        
        ImpD = RT!coeficie
        d = d & "," & TransformaComasPuntos(CStr(ImpD)) & ")"
        
        d = Aux & d
        Conn.Execute d
        RT.MoveNext
    Wend
    RT.Close
    Set RT = Nothing
    If M1 = 0 Then
        MsgBox "Ningun elemento de inmovilizado", vbExclamation
    Else
        ImpirmirListadoInmovilizados = True
    End If
    Exit Function
EImpirmirListadoInmovilizados:
    MuestraError Err.Number, Err.Description
End Function


'--------------------------------------------------------
'Listado evolucion de saldos, menusal

Public Sub FijarValoresEvolucionMensualSaldos(fec1 As Date, fec2 As Date)
    vFecha1 = fec1
    vFecha2 = fec2
    Set RT = Nothing
    Aux = "INSERT INTO Usuarios.ztmpconext (codusu, cta,  Pos, fechaent, timporteD, timporteH, saldo) VALUES (" & vUsu.Codigo & ",'"
    Contabilidad = -1
    A3 = Year(vFecha2)
    
End Sub

Public Function DatosEvolucionMensualSaldos2(ByRef Cuenta As String, ByRef DescCuenta As String, vSQL As String, MostrarTodosMeses As Boolean, EsEnHlinapu1 As Boolean, QuitarCierre As Boolean) As Byte
Dim NuloApertura As Boolean
Dim HacerAñoAnterior As Boolean
Dim importeCierreD As Currency
Dim importeCierreH As Currency

    vCta = Cuenta
    ObtenerApertura False, vFecha1, vFecha2, NuloApertura
        
    If QuitarCierre Then
        ObtenerPerdidasyGanancias False, vFecha1, vFecha2, 1
        importeCierreD = ImCierrD 'los guardo aqui, pq luego estas variables las reutilizo
        importeCierreH = ImCierrH
    End If
    SQL = "INSERT INTO usuarios.ztmpconextcab (codusu, cuenta, cta, acumantD, acumantH, acumantT ) VALUES ("
    SQL = SQL & vUsu.Codigo & ",'" & DevNombreSQL(DescCuenta) & "','" & vCta & "',"
    If NuloApertura Then
        SQL = SQL & "0,0,0)"
        ImAcD = 0: ImAcH = 0: ImPerD = 0
    Else
        ImAcD = ImpD
        ImAcH = ImpH
        SQL = SQL & TransformaComasPuntos(CStr(ImpD)) & "," & TransformaComasPuntos(CStr(ImpH)) & ","
        ImPerD = ImpD - ImpH
        SQL = SQL & TransformaComasPuntos(CStr(ImPerD)) & ")"
        
    End If
    Conn.Execute SQL
    
    SQL = "Select * from hsaldos where codmacta = '" & vCta & "'" & vSQL
    SQL = SQL & " ORDER by anopsald,mespsald"
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    M1 = Month(vFecha1)
    A1 = Year(vFecha1)
    ImCierrD = 0: ImCierrH = 0
    NumAsiento = 0
    While Not RT.EOF
    
        A2 = RT!anopsald
        M2 = RT!mespsald
                    
        If MostrarTodosMeses Then
            If M2 <> M1 Then
                ImpD = 0
                ImpH = 0
                  
            
                If Year(vFecha1) = Year(vFecha2) Then
                    'Se ha saltado algun(os) mes(es)
                    'Los rellenaremos ?????
        
                    
                        
                    For M3 = M1 To M2 - 1
                        VFecha3 = CDate("01/" & M3 & "/" & Year(vFecha1))
                        InsertaLineaEvolucion
                    Next M3
                    M1 = M2
                                
                
                
                Else
                
                    If A1 = A2 Then
                                       
                                         'El ultimo mes en meterse fue el anterior
                        For M3 = M1 To M2 - 1
                            VFecha3 = CDate("01/" & M3 & "/" & A1)
                            InsertaLineaEvolucion
                        Next M3
                        
                        
                    Else
                        For M3 = M1 To 12
                            VFecha3 = CDate("01/" & M3 & "/" & A1)
                            InsertaLineaEvolucion
                        Next M3
                        
                        A1 = A2 'cambiamos año
                        For M3 = 1 To M2 - 1
                            VFecha3 = CDate("01/" & M3 & "/" & A1)
                            InsertaLineaEvolucion
                        Next M3
                        
                    End If
                    M1 = M2
                
                End If
            End If
        End If
        ImpD = RT!impmesde
        ImpH = RT!impmesha
        
        VFecha3 = CDate("01/" & RT!mespsald & "/" & RT!anopsald)
        If VFecha3 = vFecha1 Then
            'Hay que quitar los saldos de apertura
            ImpD = ImpD - ImAcD
            ImpH = ImpH - ImAcH
            If ImpD <> 0 Or ImpH <> 0 Then NumAsiento = 1
        Else
            If QuitarCierre Then
                If Format(VFecha3, "mmyyyy") = Format(vFecha2, "mmyyyy") Then
                    'Cierre
                    ImpD = ImpD - importeCierreD
                    ImpH = ImpH - importeCierreH
                    If ImpD <> 0 Or ImpH <> 0 Then NumAsiento = 1
                Else
                    NumAsiento = 1
                End If
            Else
                NumAsiento = 1
            End If
        End If
        
        ImCierrD = ImCierrD + ImpD
        ImCierrH = ImCierrH + ImpH
        
        If MostrarTodosMeses Then
            M3 = 1
        Else
            If ImpD = 0 And ImpH = 0 Then
                M3 = 0
            Else
                M3 = 1
            End If
        End If
        
        If M3 = 1 Then InsertaLineaEvolucion
        
        M1 = M1 + 1
        If Year(vFecha1) <> Year(vFecha2) Then   'Si años partidos
            If M1 > 12 Then
                M1 = 1   'Ponemos el mes a 1 otra vez
                A1 = Year(vFecha2)
            End If
        End If
                
        
                
        RT.MoveNext
    Wend
    RT.Close
  
    If NumAsiento = 0 Then
        SQL = " WHERE codusu =" & vUsu.Codigo & " AND cta='" & vCta & "'"
        Conn.Execute "DELETE FROM Usuarios.ztmpconext" & SQL
        Conn.Execute "DELETE FROM Usuarios.ztmpconextcab" & SQL
        Exit Function
    End If
    
    
    If MostrarTodosMeses Then
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        
        
        
            If M1 <= 12 Then
                'Se ha saltado algun(os) mes(es)
                'Los rellenaremos ?????
    
                ImpD = 0
                ImpH = 0
            
                'Año natural
                For M3 = M1 To 12
                    VFecha3 = CDate("01/" & M3 & "/" & A2)
                    InsertaLineaEvolucion
                Next M3
            End If
        Else
        
            'Años tipo cooperativas
            
            ImpD = 0
            ImpH = 0
            
            'Veremos donde se ha quedado, si en la mitad del año primero o en el segundo
            If Year(vFecha2) <> A1 Then
            
                'Se ha quedado en la primera parte de los años
                'Si el mes donde se ha quedado es el ultimo
             
                    For M3 = M1 To 12
                        VFecha3 = CDate("01/" & M3 & "/" & A2)
                        InsertaLineaEvolucion
                    Next M3
                    M1 = 1
                    M2 = Month(vFecha2)
            Else
                'Stop
                M2 = Month(vFecha2)
            End If
            
            'OK Hay que rellenar
            If M1 <= M2 Then
                'Rellenamos primero el año1
                
                For M3 = M1 To M2
                    VFecha3 = CDate("01/" & M3 & "/" & Year(vFecha2))
                    InsertaLineaEvolucion
                Next M3

            End If
            
            

        End If
    End If
    
    'Updateo el total
    'ImPerD = ImPerD + (ImCierrD - ImCierrH)
    ImCierrD = ImCierrD + ImAcD
    ImCierrH = ImCierrH + ImAcH
    SQL = "UPDATE Usuarios.ztmpconextcab SET acumtotD=" & TransformaComasPuntos(CStr(ImCierrD))
    SQL = SQL & " , acumtotH=" & TransformaComasPuntos(CStr(ImCierrH))
    SQL = SQL & " , acumtotT=" & TransformaComasPuntos(CStr(ImPerD))
    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cta='" & vCta & "'"
    Conn.Execute SQL
End Function


Private Sub InsertaLineaEvolucion()
    'Aux = "INSERT INTO ztmpconext (codusu, cta,  Pos,numdiari, fechaent,
    'timporteD, timporteH, saldo) VALUES ("
    Codigo = Aux & vCta & "'," & 1 & ",'" & Format(VFecha3, FormatoFecha) & "',"
    Codigo = Codigo & TransformaComasPuntos(CStr(ImpD)) & "," & TransformaComasPuntos(CStr(ImpH)) & ","
    ImPerD = ImPerD + (ImpD - ImpH)
    Codigo = Codigo & TransformaComasPuntos(CStr(ImPerD)) & ")"
    Conn.Execute Codigo
    
End Sub


'-------------------------------------------------------------
Public Function BorrarCuenta(Cuenta As String, ByRef L1 As Label) As String
On Error GoTo Salida
Dim SQL As String
Dim RS As ADODB.Recordset

'Con ls tablas declarads sin el ON DELETE , no dejara borrar
BorrarCuenta = "Error procesando datos"
Set RS = New ADODB.Recordset



'Nuevo 15 Noviembre 2006.
'Comprobare a mano hlinapu tanto en cta como en contrapr
'Lo hare com TieneDatosSQLCount que utiliza el count
L1.Caption = "Historicos"
L1.Refresh
SQL = "SELECT count(*) from hlinapu where codmacta = '" & Cuenta & "'"
If TieneDatosSQLCount(RS, SQL, 0) Then
    BorrarCuenta = "Cuenta en historico de apuntes"
    GoTo Salida
End If


SQL = "SELECT count(*) from hlinapu where ctacontr ='" & Cuenta & "'"
If TieneDatosSQLCount(RS, SQL, 0) Then
    BorrarCuenta = "Contrapartida en historico de apuntes"
    GoTo Salida
End If



'lineas de apuntes, contrapartidads   -->1
SQL = "Select * from linasipre where ctacontr ='" & Cuenta & "'"
If TieneDatosSQL(RS, SQL) Then
    BorrarCuenta = "Contrapartida en asientos predefinidos"
    GoTo Salida
End If

SQL = "Select * from linapu where codmacta ='" & Cuenta & "'"
If TieneDatosSQL(RS, SQL) Then
    BorrarCuenta = "cuenta en introduccion de asientos"
    GoTo Salida
End If

'-->2
SQL = "Select * from linapu where ctacontr ='" & Cuenta & "'"
If TieneDatosSQL(RS, SQL) Then
    BorrarCuenta = "Contrapartida en introduccion de asientos"
    GoTo Salida
End If


'Cerrados
'----------------------------------




L1.Caption = "Otras tablas"
L1.Refresh

'-->3
'Otras tablas
'Reparto de gastos para inmovilizado
SQL = "Select codmacta2 from sbasin where codmacta2='" & Cuenta & "'"
If TieneDatosSQL(RS, SQL) Then
    BorrarCuenta = "Reparto de gastos para inmovilizado"
    GoTo Salida
End If

'-->4
SQL = "Select * from presupuestos where codmacta ='" & Cuenta & "'"
If TieneDatosSQL(RS, SQL) Then
    BorrarCuenta = "Presupuestos"
    GoTo Salida
End If




'-->5    Referencias a ctas desde eltos de inmovilizado
SQL = "select codinmov from sinmov where codmact1='" & Cuenta & "'"
SQL = SQL & " or codmact2='" & Cuenta & "'"
SQL = SQL & " or codmact3='" & Cuenta & "'"
SQL = SQL & " or codprove='" & Cuenta & "'"
If TieneDatosSQL(RS, SQL) Then
    BorrarCuenta = "Elementos de inmovilizado"
    GoTo Salida
End If





'-->6    Referencias a ctas desde eltos de inmovilizado
SQL = "select codiva from samort where codiva='" & Cuenta & "'"
If TieneDatosSQL(RS, SQL) Then
    BorrarCuenta = "IVA en elmentos de inmovilizado"
    GoTo Salida
End If

'Cta bancaria
    SQL = "select codmacta from ctabancaria where codmacta='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "Asociado a cuenta bancaria."
        GoTo Salida
    End If
    
DoEvents
If vEmpresa.TieneTesoreria Then
    L1.Caption = "Tesoreria"
    L1.Refresh
    
    'Habra k buscar en las tablas de tesoreria, k no esten enlazadas
    ' con FOREING KEY
'    SQL = "Select codmacta from scaja where codmacta = '" & Cuenta & "'"
'    If TieneDatosSQL(RS, SQL) Then
'        BorrarCuenta = "TESORERIA: Cuenta de caja."
'        GoTo Salida
'    End If

        
    SQL = "select Ctacaja from susucaja where ctacaja='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "TESORERIA: Usuarios - caja."
        GoTo Salida
    End If

    SQL = "select codmacta from Departamentos where codmacta='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "TESORERIA: Departamentos."
        GoTo Salida
    End If
            
    SQL = "select codmacta from scobro where codmacta='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "TESORERIA: Cobros."
        GoTo Salida
    End If
    
    SQL = "select ctaprove  from spagop where ctaprove ='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "TESORERIA: Pagos."
        GoTo Salida
    End If
    
    
    SQL = "select ctaingreso from ctabancaria where ctaingreso='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "TESORERIA: Pagos. ctaingreso"
        GoTo Salida
    End If
    
        
            
    'Contrapartida de gastosfijos
    SQL = "select contrapar from sgastfij where contrapar='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "TESORERIA: Gastos fijos. contrapar"
        GoTo Salida
    End If
    
    
    SQL = "select ctagastos from ctabancaria where ctagastos='" & Cuenta & "'"
    If TieneDatosSQL(RS, SQL) Then
        BorrarCuenta = "TESORERIA: Cuenta bancaria. Cta gastos."
        GoTo Salida
    End If
    
End If


'------------------------------------------------------------------------
L1.Caption = "Historicos"
L1.Refresh
SQL = "SELECT count(*) from hlinapu1 where codmacta = '" & Cuenta & "'"
If TieneDatosSQLCount(RS, SQL, 0) Then
    BorrarCuenta = "Cuenta en historico de apuntes (cerrados)"
    GoTo Salida
End If


SQL = "SELECT count(*) from hlinapu1 where ctacontr ='" & Cuenta & "'"
If TieneDatosSQLCount(RS, SQL, 0) Then
    BorrarCuenta = "Contrapartida en historico de apuntes (cerrados)"
    GoTo Salida
End If





'------------------------------------------------------------------------
L1.Caption = "SALDOS"
L1.Refresh



SQL = ValoresEnHsaldos("hsaldos", Cuenta, RS)
If SQL <> "" Then
    BorrarCuenta = SQL
    GoTo Salida
End If

SQL = ValoresEnHsaldos("hsaldos1", Cuenta, RS)
If SQL <> "" Then
    BorrarCuenta = SQL
    GoTo Salida
End If

SQL = ValoresEnHsaldos("hsaldosanal", Cuenta, RS)
If SQL <> "" Then
    BorrarCuenta = SQL
    GoTo Salida
End If
SQL = ValoresEnHsaldos("hsaldosanal1", Cuenta, RS)
If SQL <> "" Then
    BorrarCuenta = SQL
    GoTo Salida
End If






'SI kkega aqui es k ha ido bien
BorrarCuenta = ""
Salida:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar ctas." & Err.Description
    Set RS = Nothing
 
End Function


Private Function ValoresEnHsaldos(Tabla As String, codmacta As String, ByRef Rss As ADODB.Recordset) As String
Dim C As String
Dim SaldoDistintoCero As Boolean
Dim TieneSaldos As Boolean

    On Error GoTo EValoresEnHsaldos
    If Tabla = "hsaldos" Or Tabla = "hsaldos1" Then
        'impmesde impmesha
        C = "Select sum(impmesde) debe,sum(impmesha) haber from " & Tabla & " WHERE codmacta = '" & codmacta & "'"
    Else
        'debccost habccost
        C = "Select sum(debccost) debe,sum(habccost) haber from " & Tabla & " WHERE codmacta = '" & codmacta & "'"

    End If
    Rss.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SaldoDistintoCero = False
    TieneSaldos = False
    If Not Rss.EOF Then
        If Not IsNull(Rss!Debe) Then
            If Rss!Debe <> 0 Then SaldoDistintoCero = True
            TieneSaldos = True
        End If
        If Not IsNull(Rss!Haber) Then
            If Rss!Haber <> 0 Then SaldoDistintoCero = True
            TieneSaldos = True
        End If
    End If
    Rss.Close
    
    If Not TieneSaldos Then
        ValoresEnHsaldos = "" 'Todo perfecto
    
    Else
        If SaldoDistintoCero Then
            ValoresEnHsaldos = "Registros en tabla:" & Tabla
        Else
            'Tiene registros, pero son CERO
            C = "DELETE from " & Tabla & " WHERE codmacta = '" & codmacta & "'"
            Conn.Execute C
            ValoresEnHsaldos = "" 'TOdo bien
        End If
    End If
    Exit Function
EValoresEnHsaldos:
    
    ValoresEnHsaldos = "Error: " & Err.Description
    Err.Clear
End Function
'le pasamos el SQL y vemos si tiene algun dato
Private Function TieneDatosSQL(ByRef RS As ADODB.Recordset, vSQL As String) As Boolean
    TieneDatosSQL = False
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then TieneDatosSQL = True
    RS.Close

End Function


Private Function TieneDatosSQLCount(ByRef RS As ADODB.Recordset, vSQL As String, IndexdelCount As Integer) As Boolean
    TieneDatosSQLCount = False
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(IndexdelCount)) Then If RS.Fields(IndexdelCount) > 0 Then TieneDatosSQLCount = True
    End If
        
    RS.Close

End Function

'-----------------------------------------------------------------------
'
'   I N F O R M E S        C R I S T A L
Public Function DevNombreInformeCrystal(QueInforme As Integer) As String

    DevNombreInformeCrystal = DevuelveDesdeBD("informe", "scryst", "codigo", CStr(QueInforme), "N")
    If DevNombreInformeCrystal = "" Then
        MsgBox "Opcion NO encontrada: " & QueInforme, vbExclamation
        DevNombreInformeCrystal = "ERROR"
    End If

End Function






'------------------------------------------------------------------
' BALANCE INICIO EJERCICIO
'
'   Es un balance que tiene todos los saldos de las cuentas a fecha
' inicio de ejerecicio. Es decir, tiene la apertura mas todos los apuntes que
' se hayan introducido con esa fecha
'
' Niveles:  Sera un string con los niveles del balance
Public Function CargaBalanceInicioEjercicio(Niveles As String) As Boolean


On Error GoTo ECargaBalanceInicioEjercicio
    CargaBalanceInicioEjercicio = False
    
    vCta = "INSERT INTO Usuarios.ztmpbalancesumas (codusu,cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) "
    SQL = "select " & vUsu.Codigo & ",hlinapu.codmacta,nommacta,sum(timported) debe,"
    SQL = SQL & " sum(timporteH) haber from hlinapu,cuentas "
    SQL = SQL & " where cuentas.codmacta = hlinapu.codmacta and fechaent='" & Format(vParam.fechaini, FormatoFecha)
    SQL = SQL & "' and codconce = 970" 'APERTURA
    SQL = SQL & " group by 1,2 order by 2"
    
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RT.EOF
        ImpD = DBLet(RT!Debe, "N")
        ImpH = DBLet(RT!Haber, "N")
        'Apertura. La carga sobre los valores Imcierrd y H
        BuscarValorEnPrecargado RT!codmacta
        ImPerD = ImpD - ImCierrD  'Para obtener los valores del periodo reales
        ImPerH = ImpH - ImCierrH
        'Finalmente INsert mostraremos
        SQL = SQL & ",(" & vUsu.Codigo & ",'" & RT!codmacta & "','" & DevNombreSQL(RT!nommacta) & "',"
        SQL = SQL & TransformaComasPuntos(CStr(ImCierrD)) & "," & TransformaComasPuntos(CStr(ImCierrH))
        'ANterior
        SQL = SQL & ",0,0,"
        'Periodo
        SQL = SQL & TransformaComasPuntos(CStr(ImPerD)) & "," & TransformaComasPuntos(CStr(ImPerH)) & ","
        'Total
        If ImpD >= ImpH Then
            ImpD = ImpD - ImpH
            SQL = SQL & TransformaComasPuntos(CStr(ImpD)) & ",0)"
        Else
            ImpH = ImpH - ImpD
            SQL = SQL & "0," & TransformaComasPuntos(CStr(ImpH)) & ")"
        End If
        RT.MoveNext
        
        If Len(SQL) > 100000 Then
            SQL = Mid(SQL, 2) 'kito la primera coma
            SQL = vCta & " VALUES " & SQL
            Conn.Execute SQL
            SQL = ""
        End If
    Wend
    RT.Close
    
    
    If SQL <> "" Then
        SQL = Mid(SQL, 2) 'kito la primera coma
        SQL = vCta & " VALUES " & SQL
        Conn.Execute SQL
        SQL = ""
    End If




    'Ya estan cargados a ultimo nivel. AHora cogere y segun los niveles se vean o no
    'Hare un insert into group by
    
    For M1 = 1 To 9
        If Mid(Niveles, M1, 1) = "1" Then

            '---------------------------------------------------------------
                    
            SQL = "select " & vUsu.Codigo & ",substring(cta,1," & M1 & "), nomcta, sum(aperturaD), sum(aperturaH), sum(acumAntD), sum(acumAntH),"
            SQL = SQL & "sum(acumPerD), sum(acumPerH), sum(TotalD), sum(TotalH) from usuarios.ztmpbalancesumas where codusu = " & vUsu.Codigo & " and cta like '" & String(vEmpresa.DigitosUltimoNivel, "_") & "' GROUP by 1,2"
            SQL = vCta & SQL
            Conn.Execute SQL
            
            'Updateo las nommactas
            SQL = "Select codmacta,nommacta from cuentas where codmacta like '" & String(M1, "_") & "'"
            RsBalPerGan.Close
            RsBalPerGan.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            SQL = "Select cta from Usuarios.ztmpbalancesumas where codusu = " & vUsu.Codigo & " and cta like '" & String(M1, "_") & "' GROUP BY 1"
            RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                RsBalPerGan.Find "codmacta = '" & RT!Cta & "'", , adSearchForward, 1
                If RsBalPerGan.EOF Then
                    SQL = "###"
                Else
                    SQL = DevNombreSQL(RsBalPerGan!nommacta)
                End If
                SQL = "UPDATE Usuarios.ztmpbalancesumas set nomcta = '" & SQL & "' WHERE cta = '" & RT!Cta & "' and codusu = " & vUsu.Codigo
                Conn.Execute SQL
                RT.MoveNext
            Wend
            RT.Close
        End If
    Next
    'Si no quiere a ultimo nivel me cargo a ultimo nivel
    If Mid(Niveles, 10, 1) = "0" Then
        SQL = "DELETE FROM Usuarios.ztmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta like '" & String(vEmpresa.DigitosUltimoNivel, "_") & "'"
        Conn.Execute SQL
    End If
        
    
    CargaBalanceInicioEjercicio = True
ECargaBalanceInicioEjercicio:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RT = Nothing
End Function



