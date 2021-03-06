VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarfechaini As Date 'copia local
Private mvarfechafin As Date 'copia local
Private mvarautocoste As Boolean 'copia local
Private mvaremitedia As Boolean 'copia local
Private mvarlistahco As Boolean 'copia local
Private mvarnumdiapr As Integer 'copia local
Private mvarconcefpr As Integer 'copia local
Private mvarconceapr As Integer 'copia local
Private mvarnumdiacl As Integer 'copia local
Private mvarconcefcl As Integer 'copia local
Private mvarconceacl As Integer 'copia local
Private mvarlimimpcl As Currency 'copia local
'Private mvarconpresu As Boolean 'copia local
Private mvarperiodos As String 'copia local
Private mvargrupogto As String 'copia local
Private mvargrupovta As String 'copia local
Private mvarctaperga As String 'copia local
Private mvarabononeg As Boolean 'copia local
Private mvargrupoord As String 'copia local
Private mvartinumfac As String 'copia local
Private mvarmodhcofa As Boolean 'copia local
Private mvaranofactu As Integer 'copia local
Private mvarperfactu As Integer 'copia local
Private mvarnctafact As String 'copia local
Private mvarCodiNume As Byte
Private mvarAsienActAuto As Boolean 'copia local
Private mvarConstructoras As Boolean

'Modificacion del 28 de Octubre de 2003
Private mvarContabilizaFactura As Boolean

'Private mvardiremail As String 'copia local
'Private mvarSmtpHost As String 'copia local
'Private mvarSmtpUser As String 'copia local
'Private mvarSmtpPass As String 'copia local


'Modificacion 19 Diciembre


'Modificacion 2 Agosto 2004
Private mvarCCenFacturas  As Boolean
Private mvarSubgrupo1  As String 'copia local
Private mvarSubgrupo2 As String 'copia local


'Texto para la ampliacion de facturas.
' Descripcion concepto normal| ABONO
Private mvarAmpliacionFacurasCli As String
Private mvarAmpliacionFacurasPro As String



Private mvarEnlazaCtasMultibase As String

'27 Marzo 2007
Private mvarAgenciaViajes As Boolean

'Nuevo plan contable
Private mvarNuevoPlanContable As Boolean
Private mvarFechaActiva As Date 'copia local

'Bloqueo de cuentas a partir de una fecha. EN la introduccion (facturas,astos,inmov, presup,banco....)
'Para ello tendremos
'                      - mvarCuentasBloq    codmacta:fecbloq|codmacta:....
'                   Habra una funcion privada que rellenara esta variable (DE SOLO LECTURA)
Private mvarCuentasBloqueadas As String



'Para autoocion. Ellos el grupo de analitica 9 VAN a seguir llevandolo
Private mvarAutomocion As String

Private mvarTicketsEn340LetraSerie As Boolean
Private mvarPresentacion349Mensual As Boolean


Private mvarEnviarPorOutlook As Boolean



Private mvarGranEmpresa As Boolean   'Por si utiliza las cuentas 8 y 9 con otro fin ... para que no
                                     'se de gran empresa

'ALZIRA  Ene 2013
' Solo lleva GET.  Se cambia en BD
' Sera el numero contabilidad destino para los trasapaso de la cta bancos 555
Private mvarTraspasCtasBanco As Integer


'DIC 2014
'Importacion facturas
'   0.NO
'   1.- Navarres
Private mvarImportarFacturasProv As Byte


Private mvarLetraSerieAutofactura As String


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As Recordset
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo Cparametros
'------------------------------------------------

Public Property Let Constructoras(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarConstructoras = vData
End Property

Public Property Get Constructoras() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     Constructoras = mvarConstructoras
End Property


Public Property Let AgenciaViajes(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarAgenciaViajes = vData
End Property

Public Property Get AgenciaViajes() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     AgenciaViajes = mvarAgenciaViajes
End Property


'''Public Property Let diario43(ByVal vData As Integer)
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado izquierdo de la asignacion
''''Syntax: X.numdiapr = 5
'''     mvardiario43 = vData
'''End Property
'''
'''Public Property Get diario43() As Integer
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado derecho de la asignacion
''''Syntax: Debug.Print X.numdiapr
'''     diario43 = mvardiario43
'''End Property
'''
'''
'''Public Property Let conce43(ByVal vData As Integer)
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado izquierdo de la asignacion
''''Syntax: X.concefpr = 5
'''     mvarconce43 = vData
'''End Property
'''
'''Public Property Get conce43() As Integer
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado derecho de la asignacion
''''Syntax: Debug.Print X.concefpr
'''     conce43 = mvarconce43
'''End Property



Public Property Let CCenFacturas(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarCCenFacturas = vData
End Property

Public Property Get CCenFacturas() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     CCenFacturas = mvarCCenFacturas
End Property

Public Property Let Subgrupo1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarSubgrupo1 = vData
End Property

Public Property Get Subgrupo1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     Subgrupo1 = mvarSubgrupo1
End Property

Public Property Let Subgrupo2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarSubgrupo2 = vData
End Property

Public Property Get Subgrupo2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     Subgrupo2 = mvarSubgrupo2
End Property

'Private mvarAmpliacionFacurasCli As String
'Private mvarAmpliacionFacurasPro As String

Public Property Get AmpliacionFacurasCli() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     AmpliacionFacurasCli = mvarAmpliacionFacurasCli
End Property

Public Property Get AmpliacionFacurasPro() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     AmpliacionFacurasPro = mvarAmpliacionFacurasPro
End Property



'Solo GET

Public Property Get CuentasBloqueadas() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     CuentasBloqueadas = mvarCuentasBloqueadas
End Property




Public Property Let ContabilizaFactura(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarContabilizaFactura = vData
End Property

Public Property Get ContabilizaFactura() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     ContabilizaFactura = mvarContabilizaFactura
End Property


''Direccion E-mail.
'Public Property Let diremail(ByVal vData As String)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.periodos = 5
'     mvardiremail = vData
'End Property
'
'Public Property Get diremail() As String
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.periodos
'     diremail = mvardiremail
'End Property
'
'
''Host -----------------------------------------------
'Public Property Let SmtpHOST(ByVal vData As String)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.periodos = 5
'     mvarSmtpHost = vData
'End Property
'
'Public Property Get SmtpHOST() As String
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.periodos
'     SmtpHOST = mvarSmtpHost
'End Property
'
'
'
'
''Usuarios SMTP ------------------------------------------
'Public Property Let SmtpUser(ByVal vData As String)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.periodos = 5
'     mvarSmtpUser = vData
'End Property
'
'Public Property Get SmtpUser() As String
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.periodos
'     SmtpUser = mvarSmtpUser
'End Property
'
'
''Password SMTP------------------------------
'Public Property Let SmtpPass(ByVal vData As String)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.periodos = 5
'     mvarSmtpPass = vData
'End Property
'
'Public Property Get SmtpPass() As String
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.periodos
'     SmtpPass = mvarSmtpPass
'End Property






Public Property Let fechaini(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarfechaini = vData
End Property

Public Property Get fechaini() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     fechaini = mvarfechaini
End Property


Public Property Let fechafin(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarfechafin = vData
End Property

Public Property Get fechafin() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     fechafin = mvarfechafin
End Property


Public Property Let autocoste(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.autocoste = 5
     mvarautocoste = vData
End Property

Public Property Get autocoste() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.autocoste
     autocoste = mvarautocoste
End Property


Public Property Let emitedia(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.emitedia = 5
     mvaremitedia = vData
End Property

Public Property Get emitedia() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.emitedia
     emitedia = mvaremitedia
End Property


Public Property Let listahco(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.listahco = 5
     mvarlistahco = vData
End Property

Public Property Get listahco() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.listahco
     listahco = mvarlistahco
End Property


Public Property Let numdiapr(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdiapr = 5
     mvarnumdiapr = vData
End Property

Public Property Get numdiapr() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdiapr
     numdiapr = mvarnumdiapr
End Property


Public Property Let concefpr(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.concefpr = 5
     mvarconcefpr = vData
End Property

Public Property Get concefpr() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.concefpr
     concefpr = mvarconcefpr
End Property


Public Property Let conceapr(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.conceapr = 5
     mvarconceapr = vData
End Property

Public Property Get conceapr() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.conceapr
     conceapr = mvarconceapr
End Property


Public Property Let numdiacl(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdiacl = 5
     mvarnumdiacl = vData
End Property

Public Property Get numdiacl() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdiacl
     numdiacl = mvarnumdiacl
End Property


Public Property Let concefcl(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.concefcl = 5
     mvarconcefcl = vData
End Property

Public Property Get concefcl() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.concefcl
     concefcl = mvarconcefcl
End Property


Public Property Let conceacl(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.conceacl = 5
     mvarconceacl = vData
End Property

Public Property Get conceacl() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.conceacl
     conceacl = mvarconceacl
End Property
Public Property Let EnlazaCtasMultibase(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.limimpcl = 5
     mvarEnlazaCtasMultibase = vData
End Property

Public Property Get EnlazaCtasMultibase() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.limimpcl
     EnlazaCtasMultibase = mvarEnlazaCtasMultibase
End Property



Public Property Let limimpcl(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.limimpcl = 5
     mvarlimimpcl = vData
End Property

Public Property Get limimpcl() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.limimpcl
     limimpcl = mvarlimimpcl
End Property


'Public Property Let conpresu(ByVal vData As Boolean)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.conpresu = 5
'     mvarconpresu = vData
'End Property
'
'Public Property Get conpresu() As Boolean
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.conpresu
'     conpresu = mvarconpresu
'End Property


Public Property Let periodos(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarperiodos = vData
End Property

Public Property Get periodos() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     periodos = mvarperiodos
End Property


Public Property Let grupogto(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupogto = 5
     mvargrupogto = vData
End Property

Public Property Get grupogto() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupogto
     grupogto = mvargrupogto
End Property


Public Property Let grupovta(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupovta = 5
     mvargrupovta = vData
End Property

Public Property Get grupovta() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupovta
     grupovta = mvargrupovta
End Property


Public Property Let ctaperga(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ctaperga = 5
     mvarctaperga = vData
End Property

Public Property Get ctaperga() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ctaperga
     ctaperga = mvarctaperga
End Property


Public Property Let abononeg(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarabononeg = vData
End Property

Public Property Get abononeg() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     abononeg = mvarabononeg
End Property



Public Property Let NuevoPlanContable(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarNuevoPlanContable = vData
End Property

Public Property Get NuevoPlanContable() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     NuevoPlanContable = mvarNuevoPlanContable
End Property




Public Property Let FechaActiva(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarFechaActiva = vData
End Property

Public Property Get FechaActiva() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     FechaActiva = mvarFechaActiva
End Property



''''Public Property Let intfaccl(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intfaccl = 5
''''     mvarintfaccl = vData
''''End Property
''''
''''Public Property Get intfaccl() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intfaccl
''''     intfaccl = mvarintfaccl
''''End Property
''''
''''
''''Public Property Let intfacpr(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intfacpr = 5
''''     mvarintfacpr = vData
''''End Property
''''
''''Public Property Get intfacpr() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intfacpr
''''     intfacpr = mvarintfacpr
''''End Property
''''
''''
''''Public Property Let intliapu(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intliapu = 5
''''     mvarintliapu = vData
''''End Property
''''
''''Public Property Get intliapu() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intliapu
''''     intliapu = mvarintliapu
''''End Property
''''
''''
''''Public Property Let intlihap(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intlihap = 5
''''     mvarintlihap = vData
''''End Property
''''
''''Public Property Get intlihap() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intlihap
''''     intlihap = mvarintlihap
''''End Property


Public Property Let grupoord(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupoord = 5
     mvargrupoord = vData
End Property

Public Property Get grupoord() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupoord
     grupoord = mvargrupoord
End Property


Public Property Let tinumfac(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tinumfac = 5
     mvartinumfac = vData
End Property

Public Property Get tinumfac() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tinumfac
     tinumfac = mvartinumfac
End Property


Public Property Let modhcofa(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.modhcofa = 5
     mvarmodhcofa = vData
End Property

Public Property Get modhcofa() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.modhcofa
     modhcofa = mvarmodhcofa
End Property



Public Property Let anofactu(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.anofactu = 5
     mvaranofactu = vData
End Property

Public Property Get anofactu() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.anofactu
     anofactu = mvaranofactu
End Property


Public Property Let perfactu(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.perfactu = 5
     mvarperfactu = vData
End Property

Public Property Get perfactu() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.perfactu
     perfactu = mvarperfactu
End Property


Public Property Let nctafact(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nctafact = 5
     mvarnctafact = vData
End Property

Public Property Get nctafact() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nctafact
     nctafact = mvarnctafact
End Property


Public Property Let AsienActAuto(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.AsienActAuto = 5
     mvarAsienActAuto = vData
End Property

Public Property Get AsienActAuto() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.AsienActAuto
     AsienActAuto = mvarAsienActAuto
End Property


Public Property Let CodiNume(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarCodiNume = vData
End Property

Public Property Get CodiNume() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     CodiNume = mvarCodiNume
End Property




Public Property Let Automocion(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarAutomocion = vData
End Property

Public Property Get Automocion() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     Automocion = mvarAutomocion
End Property




Public Property Let TicketsEn340LetraSerie(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarTicketsEn340LetraSerie = vData
End Property

Public Property Get TicketsEn340LetraSerie() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     TicketsEn340LetraSerie = mvarTicketsEn340LetraSerie
End Property





Public Property Let Presentacion349Mensual(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarPresentacion349Mensual = vData
End Property

Public Property Get Presentacion349Mensual() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     Presentacion349Mensual = mvarPresentacion349Mensual
End Property



Public Property Get GranEmpresa() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     GranEmpresa = mvarGranEmpresa
End Property



Public Property Let GranEmpresa(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi10 = 5
     mvarGranEmpresa = vData
End Property


Public Property Let EnviarPorOutlook(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarEnviarPorOutlook = vData
End Property

Public Property Get EnviarPorOutlook() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     EnviarPorOutlook = mvarEnviarPorOutlook
End Property




Public Property Let LetraSerieAutofactura(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarLetraSerieAutofactura = vData
End Property

Public Property Get LetraSerieAutofactura() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     LetraSerieAutofactura = mvarLetraSerieAutofactura
End Property


'Solo GET
Public Property Get TraspasCtasBanco() As Integer
     TraspasCtasBanco = mvarTraspasCtasBanco
End Property
Public Property Get ImportarFacturasProv() As Byte
     ImportarFacturasProv = mvarImportarFacturasProv
End Property


'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------



''****   AGREGAR  ***
'Public Function Agregar() As Byte
'    On Error GoTo Err1
'    SQL = "SELECT * FROM parametros"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " fechaini = " & mvarfechaini
'   ' Set Rs = DB.OpenRecordSet(Sql, dbOpenDynaset)
'    If Rs.RecordCount > 0 Then
'        Mens = "Ya existe un Registro en la Tabla parametros con esos valores."
'        MsgBox Mens, vbExclamation
'        Agregar = 1
'        'ELSE
'        Else
'            Rs.AddNew
'            Rs!fechaini = mvarfechaini
'            Rs!fechafin = mvarfechafin
'            Rs!autocoste = mvarautocoste
'            Rs!emitedia = mvaremitedia
'            Rs!contabil = mvarcontabil
'            Rs!listahco = mvarlistahco
'            Rs!numdiapr = mvarnumdiapr
'            Rs!concefpr = mvarconcefpr
'            Rs!conceapr = mvarconceapr
'            Rs!numdiacl = mvarnumdiacl
'            Rs!concefcl = mvarconcefcl
'            Rs!conceacl = mvarconceacl
'            Rs!limimpcl = mvarlimimpcl
'            'RS!conpresu = mvarconpresu
'            Rs!periodos = mvarperiodos
'            Rs!grupogto = mvargrupogto
'            Rs!grupovta = mvargrupovta
'            Rs!ctaperga = mvarctaperga
'            Rs!abononeg = mvarabononeg
'            Rs!amortiza = mvaramortiza
'            Rs!intfaccl = mvarintfaccl
'            Rs!intfacpr = mvarintfacpr
'            Rs!intliapu = mvarintliapu
'            Rs!intlihap = mvarintlihap
'            Rs!grupoord = mvargrupoord
'            Rs!tinumfac = mvartinumfac
'            Rs!modhcofa = mvarmodhcofa
'            Rs!anofactu = mvaranofactu
'            Rs!perfactu = mvarperfactu
'            Rs!nctafact = mvarnctafact
'            Rs!AsienActAuto = mvarAsienActAuto
'            Rs.Update
'            Agregar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err1:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripci�n: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Rs.Close
'    Agregar = 1
'End Function


''****   ELIMINAR  ***
'Public Function Eliminar() As Byte
'    On Error GoTo Err2
'    SQL = "SELECT * FROM parametros"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " fechaini = " & mvarfechaini
' '   Set Rs = DB.OpenRecordSet(Sql, dbOpenDynaset)
'    If Rs.RecordCount = 0 Then
'        Mens = "No existe un Registro en la Tabla parametros con esos valores."
'        MsgBox Mens, vbExclamation
'        Eliminar = 1
'        'ELSE
'        Else
'            Rs.Delete
'            Rs.MoveNext
'            If Rs.EOF Then Rs.MoveFirst
'            Eliminar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err2:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripci�n: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Rs.Close
'    Eliminar = 1
'End Function


'****   LEER  ***
Public Function Leer() As Byte
    On Error GoTo Err3
    SQL = "SELECT * FROM parametros"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic
    If RS.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarfechaini = RS!fechaini
            mvarfechafin = RS!fechafin
            mvarautocoste = RS!autocoste
            mvaremitedia = RS!emitedia
            mvarlistahco = RS!listahco
            mvarnumdiapr = RS!numdiapr
            mvarconcefpr = DBLet(RS!concefpr, "N")
            mvarconceapr = DBLet(RS!conceapr, "N")
            mvarnumdiacl = DBLet(RS!numdiacl, "N")
            mvarconcefcl = DBLet(RS!concefcl, "N")
            mvarconceacl = DBLet(RS!conceacl, "N")
            mvarlimimpcl = RS!limimpcl
            'mvarconpresu = RS!conpresu
            mvarperiodos = RS!periodos
            mvargrupogto = DBLet(RS!grupogto)
            mvargrupovta = DBLet(RS!grupovta)
            mvarctaperga = DBLet(RS!ctaperga)
            mvarabononeg = RS!abononeg
            mvargrupoord = DBLet(RS!grupoord)
            mvartinumfac = RS!tinumfac
            mvarmodhcofa = RS!modhcofa
            mvaranofactu = RS!anofactu
            mvarperfactu = RS!perfactu
            mvarnctafact = RS!nctafact
            mvarAsienActAuto = RS!AsienActAuto
            mvarCodiNume = RS!CodiNume
            
            'Modificacion del 28 de Octubre de 2003
            mvarContabilizaFactura = DBLet(RS!ContabilizaFact, "N")
'            mvardiremail = DBLet(RS!diremail)
'            mvarSmtpHost = DBLet(RS!SmtpHOST)
'            mvarSmtpUser = DBLet(RS!SmtpUser)
'            mvarSmtpPass = DBLet(RS!SmtpPass)
            
            
            'Modificacion del 19 Diciembre 2003
'            mvarintfaccl = Rs!intfaccl
'            mvarintfacpr = Rs!intfacpr
'            mvarintliapu = Rs!intliapu
'            mvarintlihap = Rs!intlihap
          '  mvarconce43 = DBLet(RS!conce43, "N")
          '  mvardiario43 = DBLet(RS!diario43, "N")
            mvarConstructoras = (RS!Constructoras = 1)
            
            
            
            
            
            ParametrosNuevos

            If ASignarConceptosFacturas Then Leer = 0
    End If

    RS.Close
    
    
                    'Para que leea si hay bloqueo de cuentas
    If Leer = 0 Then ObtenerCuentasBloqueadas
    
    
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
    Mens = Mens & "Descripci�n: " & Err.Description
    MsgBox Mens, vbExclamation
    On Error Resume Next
    CerrarRs RS
    Leer = 1
    Err.Clear
End Function



Private Sub ParametrosNuevos()
Dim Nuevos As String

                '" agosto 2004
            mvarCCenFacturas = False
            mvarSubgrupo1 = ""
            mvarSubgrupo2 = ""
            mvarEnlazaCtasMultibase = ""
            mvarAgenciaViajes = False
            mvarNuevoPlanContable = False
            mvarFechaActiva = mvarfechaini   'Por fecto se carga con la fecha inicio
            mvarAutomocion = ""
            mvarEnviarPorOutlook = False
            
            Nuevos = ""
            
            
            On Error Resume Next
            
            mvarCCenFacturas = DBLet(RS!CCenFacturas, "B")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "CCenFacturas"
            
            mvarSubgrupo1 = DBLet(RS!Subgrupo1, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Subgrupo1"
            
            mvarSubgrupo2 = DBLet(RS!Subgrupo2, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Subgrupo2"
            
            If Not IsNull(RS!FechaActiva) Then mvarFechaActiva = RS!FechaActiva
            If Err.Number <> 0 Then AgregarMsg Nuevos, "FechaActiva"
            
            
            mvarEnlazaCtasMultibase = DBLet(RS!enlaza_cta, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "enlaza_cta"
            
            mvarAgenciaViajes = (DBLet(RS!AgenciaViajes, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "AgenciaViajes"
            
            mvarNuevoPlanContable = (DBLet(RS!NuevaPlan, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "NuevaPlan"
            
            mvarAutomocion = DBLet(RS!Automocion, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Automocion"
            
            
            
            mvarTicketsEn340LetraSerie = (DBLet(RS!TicketsEn340LetraSerie, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "TicketsEn340LetraSerie"
            
            mvarPresentacion349Mensual = (DBLet(RS!Presentacion349Mensual, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Presentacion349Mensual"
            
            
            mvarEnviarPorOutlook = (DBLet(RS!EnvioDesdeOutlook, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "EnvioDesdeOutlook"
            
            mvarGranEmpresa = (DBLet(RS!GranEmpresa, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "GranEmpresa"
            
            mvarTraspasCtasBanco = DBLet(RS!TraspasCtasBanco, "N")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Trapaso bancos entre secciones"
            
            mvarImportarFacturasProv = DBLet(RS!ImportacionFrasProveedor, "N")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Importar facturas tipo"
            
            mvarLetraSerieAutofactura = DBLet(RS!LetraSerieAutofactura, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Letra serie autofactura"
                        
        If Err.Number <> 0 Then Err.Clear
        If Nuevos <> "" Then MsgBox "Faltan parametros. El programa continuara. Avise soporte tecnico" & vbCrLf & vbCrLf & Nuevos, vbExclamation
    
End Sub

Private Sub AgregarMsg(CADENA As String, Trozo As String)
    
        CADENA = CADENA & vbCrLf & Trozo
        Err.Clear
End Sub
Private Function ASignarConceptosFacturas() As Boolean
Dim Aux As String
    
    On Error GoTo EASignarConceptosFacturas
    'Private mvarconcefpr As Integer
    'Private mvarconceapr As Integer
    '
    'Private mvarconcefcl As Integer
    'Private mvarconceacl As Integer
    
    ASignarConceptosFacturas = False
    
    'Private mvarAmpliacionFacurasCli As String
    'Private mvarAmpliacionFacurasPro As String
    
    'FACTURAS PROVEEDORES
    Aux = "tipoconce"
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconcefpr), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasPro = SQL & "|"
        
    Aux = "tipoconce"
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconceapr), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasPro = mvarAmpliacionFacurasPro & SQL & "|"
    
    'FACTURAS CLIENTES
    Aux = "tipoconce"
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconcefcl), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasCli = SQL & "|"
        
    Aux = "tipoconce"
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconceacl), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasCli = mvarAmpliacionFacurasCli & SQL & "|"
    
    
    
    
    ASignarConceptosFacturas = True
    Exit Function
EASignarConceptosFacturas:
    MuestraError Err.Number, "ASignar Conceptos Facturas"
End Function





'Metodo publico para la obtencion de cuentas bloquedas
Public Sub ObtenerCuentasBloqueadas()

    
    On Error GoTo EmvarCuentasBloqueadas
    
    mvarCuentasBloqueadas = ""
    Set RS = Nothing
    Set RS = New ADODB.Recordset
    SQL = "Select codmacta,fecbloq from cuentas where apudirec = 'S' and not (fecbloq is null)"
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        mvarCuentasBloqueadas = mvarCuentasBloqueadas & RS!codmacta & ":" & RS!fecbloq & "|"
        RS.MoveNext
    Wend
    RS.Close
    
EmvarCuentasBloqueadas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener CuentasBloqueadas"


End Sub
