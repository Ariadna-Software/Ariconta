VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMantenusu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de usuarios"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   Icon            =   "frmMantenusu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":9004
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":9A16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameEditorMenus 
      Height          =   6855
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   9255
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6015
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10610
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   35
         Top             =   6360
         Width           =   975
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   34
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   6360
         Width           =   5055
      End
   End
   Begin VB.Frame FrameUsuario 
      Height          =   5415
      Left            =   1200
      TabIndex        =   15
      Top             =   720
      Width           =   5655
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   3840
         MaxLength       =   17
         PasswordChar    =   "*"
         TabIndex        =   24
         Text            =   "123456789012345"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   3600
         Width           =   5295
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2880
         Width           =   5295
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdFrameUsu 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   26
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdFrameUsu 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   25
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMantenusu.frx":10278
         Left            =   120
         List            =   "frmMantenusu.frx":10285
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "mail-password"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   42
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "mail-user"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Servidor SMTP"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "e-mail"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   2280
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Confirma Pass."
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   31
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre completo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Login"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame FrameNormal 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdModBloqueo 
         Height          =   255
         Left            =   4440
         Picture         =   "frmMantenusu.frx":102AA
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Modificar empresas bloquedas"
         Top             =   1770
         Width           =   255
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMantenusu.frx":15A8C
         Left            =   6480
         List            =   "frmMantenusu.frx":15A99
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMantenusu.frx":15ABC
         Left            =   4200
         List            =   "frmMantenusu.frx":15AC9
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   3
         Left            =   1800
         Picture         =   "frmMantenusu.frx":15AE6
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Prohibir acceso"
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdConfigMenu 
         Caption         =   "Configurar menu"
         Height          =   375
         Left            =   7320
         TabIndex        =   37
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   5655
         Begin VB.TextBox Text4 
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   240
            Width           =   4215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMantenusu.frx":1C338
            Left            =   960
            List            =   "frmMantenusu.frx":1C348
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Nivel"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmMantenusu.frx":1C37B
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo usuario"
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "frmMantenusu.frx":22BCD
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Modificar usuario"
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   2
         Left            =   1080
         Picture         =   "frmMantenusu.frx":2941F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar usuario"
         Top             =   6360
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   3480
         TabIndex        =   5
         Tag             =   $"frmMantenusu.frx":2FC71
         Top             =   2040
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resum."
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5895
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Login"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Acceso"
         Height          =   255
         Index           =   1
         Left            =   5760
         TabIndex        =   47
         Top             =   6360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "P.G.C."
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   46
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   13
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Empresas "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMantenusu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim SQL As String
Dim I As Integer


Private Sub cmdConfigMenu_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    CargarListEditorMenu
    Label7.Caption = ListView1.SelectedItem.SubItems(1)
    Me.FrameEditorMenus.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEditorMenus_Click(Index As Integer)
    If Index = 0 Then
        GuardarMenuUsuario
    End If
    Me.FrameEditorMenus.Visible = False
    
    
End Sub



Private Sub cmdFrameUsu_Click(Index As Integer)



    If Index = 0 Then
        For I = 0 To Text2.Count - 1
            Text2(I).Text = Trim(Text2(I).Text)
            If I < 4 Then
                If Text2(I).Text = "" Then
                    MsgBox Label4(I).Caption & " requerido.", vbExclamation
                    Exit Sub
                End If
            End If
        Next I
        
        If Combo2.ListIndex < 0 Then
            MsgBox "Seleccione un nivel de acceso", vbExclamation
            Exit Sub
        End If
    
        'Password
        If Text2(2).Text <> Text2(3).Text Then
            MsgBox "Password y confirmacion de password no coinciden", vbExclamation
            Exit Sub
        End If
        
        
        'Ahora vamos con los campos de e-mail
        CadenaDesdeOtroForm = ""
        For I = 4 To 7
            If Text2(I).Text <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
        Next I
        
        If Len(CadenaDesdeOtroForm) > 0 And Len(CadenaDesdeOtroForm) <> 4 Then
            MsgBox "Falta por rellenar correctamente los datos del e-mail.", vbExclamation
            CadenaDesdeOtroForm = ""
            Exit Sub
        End If
        
        
        
        
        
        
        'Compruebo que el login es unico
        I = 0
        If UCase(Label6.Caption) = "NUEVO" Then
            Set miRsAux = New ADODB.Recordset
            SQL = "Select login from Usuarios.Usuarios where login='" & Text2(0).Text & "'"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not miRsAux.EOF Then SQL = "Ya existe en la tabla usuarios uno con el login: " & miRsAux.Fields(0)
            miRsAux.Close
            Set miRsAux = Nothing
            If SQL <> "" Then
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
        Else
            'MODIFICAR
            If FrameUsuario.Tag = 0 Then
                'Estoy modificando un dato normal
                I = CInt(ListView1.SelectedItem.Text)
            Else
                'Estoy agregando un usuario que ya existia en contabiñlidad
                'es decir, le estoy asignando su NIVELUSU de contabilidad
                I = CInt(FrameUsuario.Tag)
            End If
        End If
        
        InsertarModificar I
        
        
    End If
    'Cargar usuarios
    If UCase(Label6.Caption) = "NUEVO" Then
        'CargaUsuarios
        CadenaDesdeOtroForm = ""
    Else
        'Pero cargamos el tag como coresponde
        'ListView1.SelectedItem.Tag = Combo2.ItemData(Combo2.ListIndex) & "|" & Text2(1).Text & "|"
        
        If Me.FrameUsuario.Tag <> 0 Then
            CadenaDesdeOtroForm = FrameUsuario.Tag
        Else
            CadenaDesdeOtroForm = ListView1.SelectedItem.Text
        End If
        
  
    End If
    
    CargaUsuarios
    If CadenaDesdeOtroForm <> "" Then
        For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).Text = CadenaDesdeOtroForm Then
                    Set ListView1.SelectedItem = ListView1.ListItems(I)
                    Exit For
                End If
        Next I
    End If
    DatosUsusario
    CadenaDesdeOtroForm = ""
    'Para ambos casos
    Me.FrameUsuario.Visible = False
    Me.FrameNormal.Enabled = True
    
End Sub


Private Sub InsertarModificar(ByVal CodigoUsuario As Integer)
Dim Ant As Integer
Dim Fin As Boolean

On Error GoTo EInsertarModificar

    Set miRsAux = New ADODB.Recordset
    If UCase(Label6.Caption) = "NUEVO" Then
        
        'Nuevo
        SQL = "Select codusu from Usuarios.Usuarios where codusu > 0"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Ant = 1
        I = 1
        If miRsAux.EOF Then Fin = True
        While Not Fin
            If miRsAux!codusu - Ant > 0 Then
                'Hay un salto
                I = Ant
                Fin = True
            Else
                Ant = Ant + 1
            End If
            If Not Fin Then
                miRsAux.MoveNext
                If miRsAux.EOF Then
                    Fin = True
                    I = Ant
                End If
            End If
        Wend
        miRsAux.Close

        
        SQL = "INSERT INTO Usuarios.usuarios (codusu, nomusu,  nivelusu, login, passwordpropio,dirfich) VALUES ("
        SQL = SQL & I
        SQL = SQL & ",'" & Text2(1).Text & "',"
        'Combo
        SQL = SQL & Combo2.ItemData(Combo2.ListIndex) & ",'"
        SQL = SQL & Text2(0).Text & "','"
        SQL = SQL & Text2(3).Text & "',"
        'DIR FICH tiene
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For I = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(I).Text & "|"
            Next I
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        SQL = SQL & CadenaDesdeOtroForm & ")"
        
    Else
        SQL = "UPDATE Usuarios.Usuarios Set nomusu='" & Text2(1).Text
        
        'Si el combo es administrador compruebo que no fuera en un principio SUPERUSUARIO
        If Combo2.ListIndex = 2 Then
            'Si el combo1 es 3 entonces es super
            If Combo1.ListIndex = 3 Then
                I = 0
            Else
                I = 1
            End If
        Else
            I = Combo2.ItemData(Combo2.ListIndex)
        End If
        SQL = SQL & "' , nivelusu =" & I
        'SQL = SQL & "  , login = '" & Text2(2).Text
        SQL = SQL & "  , passwordpropio = '" & Text2(3).Text & "'"
        
        
        'El e-mail
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For I = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(I).Text & "|"
            Next I
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        SQL = SQL & " ,dirfich = " & CadenaDesdeOtroForm
        
        
        
        
        'aqui, en lugar del selecteditem tengo k pasarle el codigo de usuario
        'ya que cuando es nuevo usario y cojo los datos desde otra aplicacion entonces
        'no lo tengo selected y enonces peta
        
        SQL = SQL & " WHERE codusu = " & CodigoUsuario
    End If
    Conn.Execute SQL
    CadenaDesdeOtroForm = ""
    Exit Sub
EInsertarModificar:
    MuestraError Err.Number, "EInsertarModificar"
End Sub



Private Sub cmdModBloqueo_Click()

    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un usuario", vbExclamation
        Exit Sub
    End If


    frmMensajes.Opcion = 22
    frmMensajes.Parametros = ListView1.SelectedItem.Text
    frmMensajes.Show vbModal
    DatosUsusario
End Sub

Private Sub cmdUsu_Click(Index As Integer)
    
    
    Select Case Index
    Case 0, 1
        Limpiar Me
        If Index = 0 Then
            'Nuevo usuario
            
            Label6.Caption = "NUEVO"
            I = 0 'Para el foco
        Else
            'Modificar
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            Label6.Caption = "MODIFICAR"
            Set miRsAux = New ADODB.Recordset
            SQL = "Select * from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "Error inesperado: Leer datos usuarios", vbExclamation
            Else
                'LimpiarCamposUsuario
                PonerDatosUsuario
            End If
            I = 1 'Para el foco
            FrameUsuario.Tag = 0  'Marcamos que es una modificacion desde un usuario existente
        End If
        Text2(0).Enabled = (Index = 0)
        Me.FrameNormal.Enabled = False
        Me.FrameUsuario.Visible = True
        Text2(I).SetFocus
    Case 2, 3
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        I = vUsu.Codigo Mod 1000
        If I = CInt(ListView1.SelectedItem.Text) Then
            MsgBox "El usuario es el mismo con el que esta trabajando actualmente", vbInformation
            Exit Sub
        End If
        
        If Index = 2 Then
            
            SQL = "El usuario " & ListView1.SelectedItem.SubItems(1) & " será eliminado y no tendra acceso a los programas de Ariadna (Ariconta, ariges....) ." & vbCrLf
            SQL = SQL & vbCrLf & "                              ¿Desea continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            SQL = "DELETE from Usuarios.Usuarios where codusu = " & ListView1.SelectedItem.Text
            
        Else
            SQL = "Al usuario " & ListView1.SelectedItem.SubItems(1) & " no le estará permitido el acceso al programas Ariconta." & vbCrLf
            SQL = SQL & vbCrLf & "                              ¿Desea continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            SQL = "UPDATE Usuarios.usuarios SET nivelusu = -1 WHERE codusu = " & ListView1.SelectedItem.Text
        End If
        Screen.MousePointer = vbHourglass
            Conn.Execute SQL
        
            '//El codigo siguiente seria mas logico meterlo en el modulo de usuario
            '   pero de momento un saco de cemento
            If Index = 2 Then EliminarAuxiliaresUsuario CInt(ListView1.SelectedItem.Text)
        
            CargaUsuarios
        Screen.MousePointer = vbDefault
    
    End Select

End Sub
Private Sub EliminarAuxiliaresUsuario(codusu As Long)

    On Error GoTo EEliminarAuxiliaresUsuario
    SQL = "DELETE FROM usuarios.usuarioempresa where codusu =" & codusu
    Conn.Execute SQL
    
    SQL = "DELETE FROM usuarios.appmenususuario where  codusu =" & codusu
    Conn.Execute SQL
    
    Exit Sub
EEliminarAuxiliaresUsuario:
    MuestraError Err.Number, "Eliminar Auxiliares Usuario"

End Sub
Private Sub PonerDatosUsuario()
            Text2(0).Text = miRsAux!Login
            Text2(1).Text = miRsAux!nomusu
            Text2(2).Text = miRsAux!passwordpropio
            Text2(3).Text = miRsAux!passwordpropio
            I = miRsAux!nivelusu
            If I = -1 Then I = 3
            If I < 2 Then
                Combo2.ListIndex = 2
            Else
                If I = 2 Then
                    Combo2.ListIndex = 1
                Else
                    Combo2.ListIndex = 0
                End If
            End If
       
        
        'Cargamos los datos del correo e-mail
        SQL = Trim(DBLet(miRsAux!Dirfich, "T"))
        If SQL <> "" Then
            For I = 1 To 4
                Text2(3 + I).Text = RecuperaValor(SQL, I)
            Next I
        End If

End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Combo3_Click(Index As Integer)
    If Not PrimeraVez Then DatosUsusario
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaUsuarios
    End If
    FrameEditorMenus.Visible = False
    LeerEditorMenus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.ListView1.SmallIcons = ImageList1
    Me.ListView2.SmallIcons = ImageList1
    Me.FrameUsuario.Visible = False
    Me.FrameNormal.Enabled = True
    LeerDatosCombo True
End Sub


Private Sub LeerDatosCombo(Leer As Boolean)
    On Error GoTo ELe
    SQL = App.path & "\mantenusu.dat"
    If Leer Then
        Combo3(0).ListIndex = 0
        Combo3(1).ListIndex = 0
        If Dir(SQL, vbArchive) <> "" Then
            I = FreeFile
            Open SQL For Input As #I
            Line Input #I, SQL
            Close #I
            
            I = RecuperaValor(SQL, 1)
            Combo3(0).ListIndex = I
            I = RecuperaValor(SQL, 2)
            Combo3(1).ListIndex = I
        End If
    Else
        'GUARDAR
        SQL = App.path & "\mantenusu.dat"
        If Dir(SQL, vbArchive) <> "" Then Kill SQL
        I = FreeFile
        Open SQL For Output As #I
        SQL = Combo3(0).ListIndex & "|" & Combo3(1).ListIndex & "|"
        Print #I, SQL
        Close #I
            
    End If
    Exit Sub
ELe:
    Err.Clear
End Sub

Private Sub CargaUsuarios()
Dim Itm As ListItem

    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    '                               Aquellos usuarios k tengan nivel usu -1 NO son de conta
    '  QUitamos codusu=0 pq es el usuario ROOT
    SQL = "Select * from Usuarios.Usuarios where nivelusu >=0 and codusu > 0 order by codusu"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set Itm = ListView1.ListItems.Add
        Itm.Text = miRsAux!codusu
        Itm.SubItems(1) = miRsAux!Login
        Itm.SmallIcon = 3
        'Nombre y nivel de usuario
        SQL = miRsAux!nivelusu & "|" & miRsAux!nomusu & "|"
        Itm.Tag = SQL
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ListView1.ListItems.Count > 0 Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
        DatosUsusario
    End If

End Sub



Private Sub DatosUsusario()
Dim ItmX As ListItem
On Error GoTo EDatosUsu

    ListView2.ListItems.Clear
    If ListView1.SelectedItem Is Nothing Then
        Text4.Text = ""
        Combo1.ListIndex = -1
        Exit Sub
    End If
    
    
        Text4.Text = RecuperaValor(ListView1.SelectedItem.Tag, 2)
        'NIVEL
        SQL = RecuperaValor(ListView1.SelectedItem.Tag, 1)
        '                           COMBO                      en Bd
        '                       0.- Consulta                     3
        '                       1.- Normal                       2
        '                       2.- Administrador                1
        '                       3.- SuperUsuario (root)          0
        If Not IsNumeric(SQL) Then SQL = 3
        Select Case Val(SQL)
        Case 2
            Combo1.ListIndex = 1
        Case 1
            Combo1.ListIndex = 2
        Case 0
            Combo1.ListIndex = 3
        Case Else
            Combo1.ListIndex = 0
        End Select
        
        
        
    
    
    SQL = "select empresas.codempre,nomempre,nomresum,usuarioempresa.codempre bloqueada from usuarios.empresas left join usuarios.usuarioempresa on "
    SQL = SQL & " empresas.codempre = usuarioempresa.codempre And (usuarioempresa.codusu = " & ListView1.SelectedItem.Text & " Or codusu Is Null)"
    
    
    
    'Segun sea el combo del PGC
    If Combo3(0).ListIndex > 0 Then
        SQL = SQL & " WHERE empresas.codempre "
        If Combo3(0).ListIndex = 1 Then
            SQL = SQL & " > "   'plan del 2008
        Else
            SQL = SQL & " < "   'plan del 90
        End If
        SQL = SQL & " 100 "
    End If
    
    If Combo3(1).ListIndex > 0 Then
        'BLOQUEDAS O NO
        If Combo3(0).ListIndex > 0 Then
            SQL = SQL & " AND "
        Else
            SQL = SQL & " WHERE "
        End If
        If Combo3(1).ListIndex = 1 Then SQL = SQL & " NOT "
        SQL = SQL & " (usuarioempresa.codempre is null) "
    End If
    
    'EL orden
    SQL = SQL & " ORDER BY empresas.codempre"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set ItmX = ListView2.ListItems.Add
        ItmX.Text = miRsAux.Fields(0)
        ItmX.SubItems(1) = miRsAux!nomempre
        ItmX.SubItems(2) = miRsAux!nomresum
        If miRsAux.Fields(0) > 100 Then
            ItmX.ForeColor = &H808080
            ItmX.ListSubItems(1).ForeColor = &H808080
            ItmX.ListSubItems(2).ForeColor = &H808080
        End If
        
        If IsNull(miRsAux!bloqueada) Then
            ItmX.SmallIcon = 1
        Else
            ItmX.SmallIcon = 2
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Exit Sub
EDatosUsu:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LeerDatosCombo False
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim AsignarDatos As Boolean

    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    If Index = 0 Then
        If UCase(Label6.Caption) = "NUEVO" Then
        
            'Si es nuevo entonces, primero compruebo que no existe el login
            'Si existe, y el usuario tiene nivel conta >=0 entonces
            ' existe en la conta. Si existe pero el nivel conta es -1 entonces
            'lo que hacemos es ponerle los datos y que cambie la opcion de nivel usu
            SQL = "Select * from usuarios.usuarios where login='" & Text2(0).Text & "'"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                'Tiene nivel usu
                If miRsAux!nivelusu > 0 Then
                    MsgBox "El usuario ya existe para la contabilidad", vbExclamation
                    LimpiarCamposUsuario
                    Text2(0).SetFocus
                    
                Else
                    If miRsAux!codusu = 0 Then
                        MsgBox "Esta intentando modificar datos del usuario ADMINISTRADOR", vbCritical
                        AsignarDatos = False
                    Else
                        SQL = "El usuario existe para otras aplicaciones de Ariadna Software." & vbCrLf
                        SQL = SQL & "¿Desea agragarlo como usuario a la contabilidad?"
                        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then AsignarDatos = True
                    End If
                    If AsignarDatos Then
                        PonerDatosUsuario
                        'Ahora pongo el label y el campo a disbled
                        Text2(1).SetFocus
                        Label6.Caption = "MODIFICAR"
                        Text2(0).Enabled = False
                        FrameUsuario.Tag = miRsAux!codusu 'Pongo el frame al codigo ndel usuario
                    Else
                        LimpiarCamposUsuario
                        Text2(0).SetFocus
                    End If
                End If
            End If
            miRsAux.Close
        End If
    End If
    
End Sub

Private Sub LimpiarCamposUsuario()
    For I = 0 To 7
        Text2(I).Text = ""
    Next I
End Sub
Private Sub LeerEditorMenus()
    On Error GoTo ELeerEditorMenus
    cmdConfigMenu.Visible = False
    SQL = "Select count(*) from usuarios.appmenus where aplicacion='conta'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then cmdConfigMenu.Visible = True
        End If
    End If
    miRsAux.Close
        

    
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargarListEditorMenu()
Dim Nod As Node
Dim J As Integer

    TreeView1.Nodes.Clear
    SQL = "Select * from usuarios.appmenus where aplicacion='conta'"
    SQL = SQL & " ORDER BY padre ,orden"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If DBLet(miRsAux!Padre, "N") = 0 Then
            Set Nod = TreeView1.Nodes.Add(, , "C" & miRsAux!Contador)
        Else
            SQL = "C" & miRsAux!Padre
            Set Nod = TreeView1.Nodes.Add(SQL, tvwChild, "C" & miRsAux!Contador)
        End If
        SQL = miRsAux!Name & "|"
        If Not IsNull(miRsAux!Indice) Then SQL = SQL & miRsAux!Indice
        Nod.Tag = SQL
   
        Nod.Text = miRsAux!Caption
        Nod.Checked = True
        Nod.EnsureVisible
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If TreeView1.Nodes.Count > 1 Then TreeView1.Nodes(1).EnsureVisible
    
    'AHora ire nodo a nodo buscando los k deshabilitamos de la aplicacion
    SQL = "Select * from usuarios.appmenusUsuario where aplicacion='Conta' AND codusu =" & ListView1.SelectedItem.Text
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        For I = 1 To TreeView1.Nodes.Count
            SQL = miRsAux!Tag
            If TreeView1.Nodes(I).Tag = SQL Then
                TreeView1.Nodes(I).Checked = False
                If TreeView1.Nodes(I).Children > 0 Then Recursivo2 TreeView1.Nodes(I).Child, TreeView1.Nodes(I).Checked
                Exit For
            End If
        Next I
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    
    Set miRsAux = Nothing
End Sub



Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
If Node.Children > 0 Then Recursivo2 Node.Child, Node.Checked
End Sub


Private Sub CheckarNodo(N As Node, Valor As Boolean)
Dim NO As Node
    Set NO = N.LastSibling
    Do
        N.Checked = Valor
        If N.Children > 0 Then CheckarNodo N, Valor
        If N.Next <> NO.LastSibling Then Set N = N.Next
    Loop Until NO = N
End Sub

Private Sub Recursivo2(ByVal Nod As Node, Valor As Boolean)
Dim nx As Node
Dim Aux

    
    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
        'aux = nx.Root
        'aux = nx.Parent
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
      End If
    Set nx = Nothing
End Sub


Private Sub GuardarMenuUsuario()
    SQL = "DELETE from usuarios.appmenusUsuario where aplicacion='Conta' AND codusu =" & ListView1.SelectedItem.Text
    Conn.Execute SQL
    
    I = 0
    SQL = "INSERT INTO usuarios.appmenususuario (aplicacion, codusu, codigo, tag) VALUES ('Conta'," & ListView1.SelectedItem.Text & ","
    RecursivoBD TreeView1.Nodes(1)
End Sub

Private Sub InsertaBD(vTag As String)
Dim C As String
    I = I + 1
    'SQL = "INSERT INTO appmenususuario (aplicacion, codusu, codigo, tag)
    C = SQL & I & ",'" & vTag & "')"
    Conn.Execute C
End Sub


Private Sub RecursivoBD(ByVal Nod As Node)
Dim nx As Node
Dim Aux

    
    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
        'aux = nx.Root
        'aux = nx.Parent
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
      End If
    Set nx = Nothing
End Sub


