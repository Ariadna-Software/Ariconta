VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmBackUP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmBackUP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl2.Animation Animation1 
      Height          =   915
      Left            =   300
      TabIndex        =   5
      Top             =   1500
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1614
      _Version        =   327681
      FullWidth       =   301
      FullHeight      =   61
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   1740
      TabIndex        =   3
      Top             =   3120
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   3795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "sobre ficheros locales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   4860
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Copia de seguridad :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   0
      Width           =   3480
   End
End
Attribute VB_Name = "frmBackUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const TamanyoBULK = 300000

Private Tablas() As String
Private NumTablas As Integer

Dim RS As Recordset
Dim NF As Integer
Dim Archivo As String
Dim Izquierda As String
Dim Derecha As String


'En todas las futuros backups, se trata de cargar el array tablas con las tablas(-1) a copiar


Private Sub cmdAceptar_Click()
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de copia", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Select Case Combo1.ListIndex
    Case 0
        CopiaTodo
    Case 1
        RenumeracionAsientos
    Case 2
        Cierre
    Case 3
        Amorizacion
    End Select
    'Ahora hacemos las copias
    HacerBackUp
    MsgBox "Copia finalizada en: " & Archivo, vbInformation
    cmdAceptar.Enabled = False
    Label1.Caption = ""
    PonerVideo False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = ""
    Label2.Caption = "Empresa: " & vEmpresa.nomempre
    Caption = "Backup para " & UCase(vEmpresa.nomresum)
    
    CargaCombo
End Sub


Private Sub CargaCombo()
Combo1.Clear
Combo1.AddItem "Copia todo "
Combo1.AddItem "Renumeración asientos"
Combo1.AddItem "Cierre ejercicio"
Combo1.AddItem "Amortización"
End Sub

Private Sub CopiaTodo()


    Set RS = New ADODB.Recordset
    RS.Open "SHOW TABLES", Conn, adOpenKeyset, adLockOptimistic, adCmdText
    NumTablas = 0
    While Not RS.EOF
        If LCase(Mid(RS.Fields(0), 1, 3)) = "tmp" Then
            'Las temporales no hacemos nada
        Else
            NumTablas = NumTablas + 1
        End If
        RS.MoveNext
    Wend
    
    RS.MoveFirst
    
    ReDim Tablas(NumTablas - 1)
    NumTablas = 0
    While Not RS.EOF
        If LCase(Mid(RS.Fields(0), 1, 3)) = "tmp" Then
            'Las temporales no hacemos nada
        Else
            Tablas(NumTablas) = RS.Fields(0)
            NumTablas = NumTablas + 1
        End If
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing

End Sub


Private Sub PonerVideo(Encender As Boolean)
If Encender Then
    Me.Animation1.Open App.path & "\actua.avi"
    Me.Animation1.Play
    Me.Animation1.Visible = True
Else
    Me.Animation1.Stop
    Me.Animation1.Visible = False
End If
End Sub



Private Sub HacerBackUp()
Dim I As Integer

    If NumTablas > 3 Then PonerVideo True


    Archivo = FijarCarpeta
    If Archivo = "" Then
        MsgBox "no se ha creado correctamente la carpeta de copia.", vbExclamation
        Exit Sub
    End If
        
    
    For I = 0 To NumTablas - 1
        Label1.Caption = Tablas(I) & "     (" & I + 1 & " de " & NumTablas & ")"
        Me.Refresh
        DoEvents
        BKTablas (Tablas(I))
    Next I
End Sub



Private Function FijarCarpeta() As String
Dim FE As String
Dim I As Integer

On Error GoTo EFijarCarpeta
    FijarCarpeta = ""
    If Dir(App.path & "\BACKUP", vbDirectory) = "" Then MkDir App.path & "\BACKUP"
    
    Derecha = App.path & "\BACKUP\"
    Izquierda = Format(Now, "yymmdd")
    I = -1
    Do
        I = I + 1
        FE = Format(I, "00")
        FE = Derecha & Izquierda & FE
        If Dir(FE, vbDirectory) = "" Then
            'OK
            MkDir FE
            FijarCarpeta = FE
            I = 100
        End If
    Loop Until I > 99
    Exit Function
EFijarCarpeta:
    MuestraError Err.Number
End Function



Private Sub BKTablas(Tabla As String)


    Set RS = New ADODB.Recordset
    RS.Open Tabla, Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If RS.EOF Then
        'No hace falta hacer back up
    
    Else
        NF = FreeFile
        Open Archivo & "\" & Tabla & ".sql" For Output As #NF
        CreandoLineasBackUp Tabla
        
        Close #NF
    End If
    RS.Close
    Set RS = Nothing
End Sub

Private Sub CreandoLineasBackUp(Tabla As String)
Dim Cad As String
Dim C2 As String
Dim T1 As Single
Dim B As Boolean
'

        BACKUP_TablaIzquierda RS, Izquierda
        C2 = ""
        T1 = Timer
        While Not RS.EOF
            
            BACKUP_Tabla RS, Derecha
            C2 = C2 & "," & Derecha
            If Len(C2) > TamanyoBULK Then
                C2 = Mid(C2, 2)
                Cad = "INSERT INTO " & Tabla & " " & Izquierda & " VALUES " & C2 & ";"
                Print #NF, Cad
                C2 = ""
            End If
            RS.MoveNext
            
            If Timer - T1 > 4 Then
                Me.Refresh
                T1 = Timer
                If B Then DoEvents
                B = Not B
            End If
        Wend
        
        If Len(C2) > 0 Then
            C2 = Mid(C2, 2)
            Cad = "INSERT INTO " & Tabla & " " & Izquierda & " VALUES " & C2 & ";"
            Print #NF, Cad
        End If
End Sub

Private Sub RenumeracionAsientos()
    NumTablas = 6
    ReDim Tablas(NumTablas - 1)
    Tablas(0) = "hcabapu"
    Tablas(1) = "hlinapu"
    Tablas(2) = "cabfact"
    Tablas(3) = "cabfactprov"
    Tablas(4) = "linfact"
    Tablas(5) = "linfactprov"
End Sub



Private Sub Cierre()
    NumTablas = 5
    ReDim Tablas(NumTablas - 1)
    Tablas(0) = "hcabapu"
    Tablas(1) = "hlinapu"
    Tablas(2) = "contadores"
    Tablas(3) = "hsaldos"
    Tablas(4) = "hsaldosanal"

End Sub

Private Sub Amorizacion()
    NumTablas = 3
    ReDim Tablas(NumTablas - 1)
    Tablas(0) = "samort"
    Tablas(1) = "sinmov"
    Tablas(2) = "shisin"
End Sub
