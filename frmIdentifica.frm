VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00765341&
   BorderStyle     =   0  'None
   Caption         =   "Identificacion"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2040
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single
Dim CodPC As Long

Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        
        espera 0.5
        Me.Refresh
        'MsgBox "Antes config"
        'Vemos datos de configconta.ini
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then
             vConfig.SERVER = InputBox("Servidor: ")
             vConfig.User = InputBox("Usuario: ")
             vConfig.password = InputBox("Password: ")
             vConfig.Integraciones = InputBox("Path integraciones: ")
             vConfig.Grabar
             MsgBox "Reinicie la contabilidad", vbCritical
             End
             Exit Sub
        End If
        
         
         'Abrimos conexion para comprobar el usuario
         'Luego, en funcion del nivel de usuario que tenga cerraremos la conexion
         'y la abriremos con usuario-codigo ajustado a su nivel
         If AbrirConexionUsuarios() = False Then
             MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
             End
         End If
         
         
        'Gestionar el nombredel PC para la asignacion de PC en el entorno de red
        'MsgBox "Antes gestiona pC"
        CodPC = GestionaPC2
        CadenaDesdeOtroForm = ""
         
         'La llave
         'MsgBox "Antes llave"
'         Load frmLLave
'         If Not frmLLave.ActiveLock1.RegisteredUser Then
'             'No ESTA REGISTRADO
'             frmLLave.Show vbModal
'         Else
'             Unload frmLLave
'         End If
        
         'Leemos el ultimo usuario conectado
         NumeroEmpresaMemorizar True
         
         
         'MsgBox "Antes codpc>0"
         
         If CodPC > 0 Then
            If ActualizarVersion Then
                Set Conn = Nothing
               Unload Me
               End
               Exit Sub
            End If
        End If
         
         
         
         
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then espera T1

         
         PonerVisible True
         
         If Text1(0).Text <> "" Then
            Text1(1).SetFocus
        Else
            Text1(0).SetFocus
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    PonerVisible False
    T1 = Timer
    Text1(0).Text = ""
    Text1(1).Text = ""
    
    
    Label1(3).Caption = "Vers. " & App.Major & "." & App.Minor & "." & App.Revision
    Label1(2).Caption = "Cargando ..."
    
    PrimeraVez = True
    CargaImagen
    Me.Height = 5520
    Me.Width = 7965
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.path & "\arifon.dat")
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set Conn = Nothing
        End
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub









Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        'Probar conexion usuario
        Validar
    End If
        
    
End Sub



Private Sub Validar()
Dim NuevoUsu As Usuario
Dim OK As Byte

    'Validaremos el usuario y despues el password
    Set vUsu = New Usuario
    
    
    
    If vUsu.Leer(Text1(0).Text) = 0 Then
        If vUsu.Nivel < 0 Then
            'NO tiene autorizacion de ningun nivel. Es menos 1
            OK = 3
        Else
            'Con exito
            If vUsu.PasswdPROPIO = Text1(1).Text Then
                OK = 0
            Else
                OK = 1
            End If
        End If
    Else
        OK = 2
    End If
    
    If OK <> 0 Then
        If OK = 3 Then
            MsgBox "Usuario sin autorizacion.", vbExclamation
        Else
            MsgBox "Usuario-Clave Incorrecto", vbExclamation
        End If
        
        Text1(1).Text = ""
        If OK = 2 Then
            Text1(0).SetFocus
        Else
            Text1(1).SetFocus
        End If
    Else
        'OK
        Screen.MousePointer = vbHourglass
        CadenaDesdeOtroForm = "OK"
        Label1(2).Caption = "Leyendo ."  'Si tarda pondremos texto aquin
        PonerVisible False
        Me.Refresh
        espera 0.1
        Me.Refresh
        espera 0.2
        Screen.MousePointer = vbHourglass
        HacerAccionesBD
        Unload Me
    End If

End Sub

Private Sub HacerAccionesBD()
Dim SQL As String
    
    T1 = Timer
    
    'Limpiamos datos blanace
    SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    Label1(2).Caption = Label1(2).Caption & "."
    Label1(2).Refresh

    SQL = "DELETE from Usuarios.ztmpconextcab where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    Label1(2).Caption = Label1(2).Caption & "."
    Label1(2).Refresh

    SQL = "DELETE from usuarios.ztmpconext where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    Label1(2).Caption = Label1(2).Caption & "."
    Me.Refresh
    
    SQL = "DELETE from Usuarios.zcuentas where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    Label1(2).Caption = Label1(2).Caption & "."
    Label1(2).Refresh

    SQL = "DELETE from usuarios.ztmplibrodiario where codusu= " & vUsu.Codigo
    Label1(2).Caption = Label1(2).Caption & "."
    Label1(2).Refresh
    Conn.Execute SQL
    
    SQL = "DELETE from usuarios.zdirioresum where codusu= " & vUsu.Codigo
    Label1(2).Caption = Label1(2).Caption & "."
    Conn.Execute SQL
    
    
    SQL = "DELETE from usuarios.zhistoapu where codusu= " & vUsu.Codigo
    Label1(2).Caption = Label1(2).Caption & "."
    Label1(2).Refresh
    Conn.Execute SQL
    
    SQL = "DELETE from usuarios.ztmpctaexplotacion where codusu= " & vUsu.Codigo
    Label1(2).Caption = Label1(2).Caption & "."
    Label1(2).Refresh
    Conn.Execute SQL
    
    
    Me.Refresh
    T1 = Timer - T1
    If T1 < 1 Then espera 0.7
    
    Label1(2).Visible = False
    Me.Refresh
    espera 0.2
End Sub


Private Sub PonerVisible(Visible As Boolean)
    Label1(2).Visible = Not Visible  'Cargando
    Text1(0).Visible = Visible
    Text1(1).Visible = Visible
    Label1(0).Visible = Visible
    Label1(1).Visible = Visible
End Sub




'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim Cad As String
On Error GoTo ENumeroEmpresaMemorizar


        
    Cad = App.path & "\ultusu.dat"
    If Leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            
                'El primer pipe es el usuario
                Text1(0).Text = Cad
    
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad = Text1(0).Text
        Print #NF, Cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub


Private Function ActualizarVersion() As Boolean
Dim Version As Integer
    ActualizarVersion = 0
    If Dir(App.path & "\Actualizar.exe", vbArchive) <> "" Then
        Set miRsAux = New ADODB.Recordset
        Version = HayQueActualizar
        If Version > 0 Then
            CadenaDesdeOtroForm = "Estan disponibles actualizaciones para instalarse en esta maquina. ¿Desea continuar?"
            If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbYes Then
                'LANZAMOS EL actualizador
                CadenaDesdeOtroForm = App.path & "\Actualizar.exe "
                '       Parametros
                '       applicacion    version   PC
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " CONTA " & Version & " " & CodPC
                Shell CadenaDesdeOtroForm, vbNormalNoFocus
                ActualizarVersion = True
            End If
        End If
        Set miRsAux = Nothing
        CadenaDesdeOtroForm = ""
    End If
End Function


Private Function HayQueActualizar() As Integer
Dim V As Integer
    On Error GoTo EA
    HayQueActualizar = 0
    
    CadenaDesdeOtroForm = "Select max(ver) from yVersion where app='CONTA'"
    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    V = 0
    If Not miRsAux.EOF Then V = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    If V = 0 Then Exit Function
    
    
    'YA TENGO LA ULTIMA VERSION disponible. Voy a ver cual tengo
    CadenaDesdeOtroForm = DevuelveDesdeBD("Conta", "PCs", "codpc", CStr(CodPC), "N")
    If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = 0
    NumRegElim = Val(CadenaDesdeOtroForm)
    If V > NumRegElim Then
        'OK esta desactualizado.
        'Veo cual es la version qe hay que lanzar.
        HayQueActualizar = NumRegElim + 1
    End If
        
    
    Exit Function
EA:
    Err.Clear
    Err.Clear
    Set miRsAux = Nothing
End Function

