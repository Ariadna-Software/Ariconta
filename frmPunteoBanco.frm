VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPunteoBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punteo bancario"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmPunteoBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameGenera 
      Height          =   4935
      Left            =   3000
      TabIndex        =   25
      Top             =   1380
      Width           =   5415
      Begin VB.TextBox Text11 
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdAtoCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3900
         TabIndex        =   34
         Top             =   4440
         Width           =   1155
      End
      Begin VB.CommandButton cmdAstoAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   33
         Top             =   4440
         Width           =   1155
      End
      Begin VB.TextBox txtFec 
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Text            =   "99/99/9999"
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   1440
         Width           =   3915
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   1440
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   2160
         Width           =   3915
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Documento"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   2640
         Width           =   945
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1260
         Picture         =   "frmPunteoBanco.frx":030A
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   960
         Picture         =   "frmPunteoBanco.frx":6B5C
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   720
         Picture         =   "frmPunteoBanco.frx":D3AE
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Si no pone contrapartida podra añadir mas de una linea en el asiento"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   3960
         Width           =   4995
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   42
         Top             =   540
         Width           =   495
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmPunteoBanco.frx":13C00
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Ampliación"
         Height          =   195
         Left            =   1440
         TabIndex        =   41
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Diario"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "Contrapartida"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   180
         Width           =   5175
      End
   End
   Begin VB.Frame FrameIntro 
      Height          =   2715
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5355
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar punteados"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   2100
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4020
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtFec 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox txtFec 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Text            =   "99/99/9999"
         Top             =   540
         Width           =   1035
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   3120
         Picture         =   "frmPunteoBanco.frx":13C8B
         Top             =   630
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   780
         Picture         =   "frmPunteoBanco.frx":13D16
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Left            =   1260
         Picture         =   "frmPunteoBanco.frx":13DA1
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   8
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta banco"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   1110
         Width           =   1095
      End
   End
   Begin VB.Frame FrameDatos 
      Height          =   7515
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11835
      Begin VB.CommandButton Command1 
         Height          =   435
         Left            =   9240
         Picture         =   "frmPunteoBanco.frx":1A5F3
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Buscar importe "
         Top             =   6960
         Width           =   495
      End
      Begin VB.CommandButton cmdPedirDatos 
         Height          =   435
         Left            =   9840
         Picture         =   "frmPunteoBanco.frx":20E45
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Pedir datos"
         Top             =   6960
         Width           =   495
      End
      Begin VB.CommandButton cmdsalir 
         Height          =   435
         Left            =   11160
         Picture         =   "frmPunteoBanco.frx":21847
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salir"
         Top             =   6960
         Width           =   495
      End
      Begin VB.CommandButton cmdCrearAsiento 
         Height          =   435
         Left            =   10380
         Picture         =   "frmPunteoBanco.frx":28099
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Crear asiento"
         Top             =   6960
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3060
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   7080
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "D/H"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Concepto"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6255
         Left            =   5940
         TabIndex        =   13
         Top             =   540
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "D/H"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ampliacion"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "Doble click busca el importe en el otro lado"
         Height          =   255
         Left            =   4620
         TabIndex        =   46
         Top             =   7080
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "Diferencia"
         Height          =   255
         Left            =   3060
         TabIndex        =   21
         Top             =   6840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Apuntes"
         Height          =   255
         Left            =   1620
         TabIndex        =   20
         Top             =   6840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Extractos"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   6840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Extracto bancario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   4
         Left            =   5880
         TabIndex        =   14
         Top             =   180
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmPunteoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmCuentasBancarias
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmCo As frmConceptos
Attribute frmCo.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCC As frmColCtas
Attribute frmCC.VB_VarHelpID = -1



Dim SQL As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim Importe As Currency
Dim I As Integer
Dim PrimeraSeleccion As Boolean
Dim ClickAnterior As Byte '0 Empezar 1.-Debe 2.-Haber
    
'Con estas dos variables
Dim ContadorBus As Integer
Dim Checkear As Boolean
Dim De As Currency
Dim Ha As Currency
Dim EstaLW1 As Boolean




Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    If Text1.Text <> "" Then
        'Tiene cta.
        'Veamos si la cuenta esta definida en ctas bancarias o no
        SQL = DevuelveDesdeBD("codmacta", "ctabancaria", "codmacta", Text1.Text, "T")
        If SQL <> "" Then
            'Bloqueamos manualamente la tabla, con esa cuenta
            If Not BloqueoManual(True, "PUNTEOB", Text1.Text) Then
                MsgBox "Imposible acceder a puntear la cuenta. Esta bloqueada"
            Else
                Text3(0).Text = "": Text3(1).Text = "": Text3(2).Text = ""
                'Datos ok. Vamos a ver los resultados
                Label1(4).Caption = Text1.Text & " - " & Text2.Text
                PonerTamanyo True
                Me.Refresh
                CargarDatosLw True
            End If
        Else
            MsgBox "La cuenta no esta asociada a una cuenta bancaria.", vbExclamation
        End If
    Else
        MsgBox "Introduzca la cuenta ", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargarDatosLw(BorrarImportes As Boolean)
       'Resetamos importes punteados
       If BorrarImportes Then
            De = 0
            Ha = 0
            Text3(0).Text = "": Text3(1).Text = "": Text3(2).Text = ""
        End If
        PrimeraSeleccion = True
                    
        'Cargamos los datos
        SQL = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
        Conn.Execute SQL
            
        SQL = "DELETE from tmpconext where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        
        SQL = "fechaent >= '" & Format(txtFec(0).Text, FormatoFecha)
        SQL = SQL & "' AND fechaent <= '" & Format(txtFec(1).Text, FormatoFecha) & "'"
        
        CargaDatosConExt Text1.Text, txtFec(0).Text, txtFec(1).Text, SQL, Text2.Text

                    
                    
                    
        Me.Refresh
        CargaBancario
        Me.Refresh
        CargaLineaApuntes
        Me.Refresh
End Sub



Private Sub cmdAstoAceptar_Click()
Dim NA As Long


    If txtFec(2).Text = "" Or Text9.Text = "" Or Text7.Text = "" Then
        MsgBox "Todos los campos, excepto la contrapartida, son obligados", vbExclamation
        Exit Sub
    End If
    
    'Generamos el asiento en errores
    If Not IsDate(txtFec(2).Text) Then
        MsgBox "Fecha incorrecta", vbExclamation
        Exit Sub
    End If
    
    varFecOk = FechaCorrecta2(CDate(txtFec(2).Text))
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            SQL = varTxtFec
        Else
            SQL = "Fechas fuera de ejercicio actual/siguiente"
        End If
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    Set RS = New ADODB.Recordset
    NA = 0
    SQL = "Select max(numasien) from cabapue"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then NA = DBLet(RS.Fields(0), "N")
    RS.Close
    
    'Ya tenemos el numero de asiento en errores
    NA = NA + 1
    
    'Ahora generemos la cabecera de apunte
    Screen.MousePointer = vbHourglass
    If GenerarCabecera(NA) Then
        CadenaDesdeOtroForm = ""
        If Text4.Text <> "" Then
            frmAsientosErr.DesdeNorma43 = 2
        Else
            frmAsientosErr.DesdeNorma43 = 1
        End If
        frmAsientosErr.Datos = NA & "|"
        frmAsientosErr.Show vbModal
    End If
    
    'Borramos las lineas del apunte
    Screen.MousePointer = vbHourglass
    SQL = "Delete from linapue where numasien = " & NA
    Conn.Execute SQL
    SQL = "Delete from cabapue where numasien = " & NA
    Conn.Execute SQL
    If CadenaDesdeOtroForm = "OK" Then
    
        'Aumentamos los importes punteados
        Importe = CCur(ListView1.SelectedItem.SubItems(1))
        De = De + Importe
        Ha = Ha + Importe
        PonerImportes
    
        'Puntemos el extracto
        SQL = "UPDATE norma43 SET punteada= 1 WHERE codigo=" & ListView1.SelectedItem.Tag
        Conn.Execute SQL
    
        'Para buscarlo
        NA = ListView1.SelectedItem.Tag
        'Volvemos a cargar todo
        CargarDatosLw False
        'Volvemos a siutar el select item
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Tag = NA Then
                Set ListView1.SelectedItem = ListView1.ListItems(I)
                ListView1.SelectedItem.EnsureVisible
                ListView1_DblClick
                Exit For
            End If
        Next I
    End If
    Me.FrameGenera.visible = False
    Me.FrameDatos.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAtoCancelar_Click()
    Me.FrameGenera.visible = False
    Me.FrameDatos.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub PonerTamanyo(Punteo As Boolean)
    Me.FrameDatos.visible = Punteo
    Me.FrameIntro.visible = Not Punteo
    If Punteo Then
        Me.Height = FrameDatos.Height + 400
        Me.Width = FrameDatos.Width + 100
        If Screen.Width > 12300 Then
            Me.Top = 800
            Me.Left = 800
        Else
            Me.Top = 0
            Me.Left = 0
        End If
    
    Else
        Me.Height = FrameIntro.Height + 400
        Me.Width = FrameIntro.Width + 100
        If Screen.Width > 12300 Then
            Me.Top = 4000
            Me.Left = 4000
        Else
            Me.Top = 1000
            Me.Left = 1000
        End If
    End If
          
End Sub


Private Sub cmdCrearAsiento_Click()
    'Crear asiento
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem.Checked Then
        MsgBox "Extracto ya esta punteado", vbExclamation
        Exit Sub
    End If
    
    'Deshabilitamos
    Me.FrameDatos.Enabled = False
    'Limpiamos y ponemos datos
    Me.txtFec(2).Text = Format(ListView1.SelectedItem.Text, "dd/mm/yyyy")
    
    'dIARIO POR DEFECTO DE PARAMETROS
    'Veremos si hay parametros
    SQL = DevuelveDesdeBD("diario43", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    Text9.Text = SQL
    If Text9.Text <> "" Then SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text9.Text, "N")
    Text10.Text = SQL
    
    'Concepto por defecto desde parametros
    SQL = DevuelveDesdeBD("conce43", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    Text6.Text = SQL
    If Text6.Text <> "" Then SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text6.Text, "N")
    Text7.Text = SQL
    
    'La ampliacion del concepto viene del extracto bancario
    Text8.Text = ListView1.SelectedItem.SubItems(4)
    
    Text4.Text = "": Text5.Text = ""
    Text11.Text = ""
    Label1(5).Caption = Label1(4).Caption
    'Ponemos visible
    Me.FrameGenera.visible = True
    'Ponemos el foco en doc
    Text11.SetFocus
    
End Sub

Private Sub cmdPedirDatos_Click()
    'Desbloqueamos
    Label1(5).Caption = Label1(4).Caption
    BloqueoManual False, "PUNTEOB", Text1.Text
    PonerTamanyo False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Command1_Click()
    If EstaLW1 Then
        ListView1_DblClick
    Else
        ListView2_DblClick
    End If
End Sub

Private Sub Form_Load()
    FrameGenera.visible = False
    Text1.Text = ""
    Text2.Text = ""
    txtFec(0).Text = Format(DateAdd("m", -1, Now), "dd/mm/yyyy")
    txtFec(1).Text = Format(Now, "dd/mm/yyyy")
    PonerTamanyo False
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'Desbloqueamos
    BloqueoManual False, "PUNTEOB", Text1.Text
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtFec(CInt(txtFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    Text4.Text = RecuperaValor(CadenaSeleccion, 1)
    Text5.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCo_DatoSeleccionado(CadenaSeleccion As String)
    Text6.Text = RecuperaValor(CadenaSeleccion, 1)
    Text7.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text1.Text = RecuperaValor(CadenaSeleccion, 1)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    Text9.Text = RecuperaValor(CadenaSeleccion, 1)
    Text10.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click()
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub

Private Sub Image2_Click()
    Set frmCo = New frmConceptos
    frmCo.DatosADevolverBusqueda = "0|1|"
    frmCo.Show vbModal
    Set frmCo = Nothing
End Sub

Private Sub Image3_Click()
    Set frmCC = New frmColCtas
    frmCC.DatosADevolverBusqueda = "0|1"
    frmCC.ConfigurarBalances = 3  'NUEVO
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub imgCuentas_Click()
    Set frmCta = New frmCuentasBancarias
    frmCta.DatosADevolverBusqueda = "0|1|"
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub

Private Sub imgppal_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    txtFec(0).Tag = Index
    If txtFec(Index).Text <> "" Then
        If IsDate(txtFec(Index).Text) Then frmC.Fecha = CDate(txtFec(0).Text)
    End If
    frmC.Show vbModal
    Set frmC = Nothing
End Sub


Private Sub ListView1_Click()
    EstaLW1 = True
End Sub

Private Sub ListView1_DblClick()
Dim J As Integer
Dim Find As Boolean
Dim Fin As Long

    EstaLW1 = True
    If ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    J = ListView2.SelectedItem.Index
    Find = False
    Fin = ListView2.ListItems.Count
    Do
        For I = J To Fin
            If ListView2.ListItems(I).SubItems(1) = ListView1.SelectedItem.SubItems(1) Then
                If ListView2.ListItems(I).SubItems(2) <> ListView1.SelectedItem.SubItems(2) Then
                    'Ha encontrado con el mismo importe y signos distintos D-H
                    Set ListView2.SelectedItem = ListView2.ListItems(I)
                    ListView2.SelectedItem.EnsureVisible
                    Find = True
                    Exit For
                End If
            End If
        Next I
        If Not Find Then
            If J > 1 Then
                Fin = J
                J = 1
            Else
                Find = True
            End If
        End If
                
    Loop Until Find
End Sub



'Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub

Private Sub ListView2_Click()
    EstaLW1 = False
End Sub

Private Sub ListView2_DblClick()
Dim J As Integer
Dim Find As Boolean
Dim Fin As Long

    EstaLW1 = False
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then
        J = 0
    Else
        J = ListView1.SelectedItem.Index + 1
    End If
    Find = False
    Fin = ListView1.ListItems.Count
    Do
        For I = J To Fin
            If ListView1.ListItems(I).SubItems(1) = ListView2.SelectedItem.SubItems(1) Then
                If ListView1.ListItems(I).SubItems(2) <> ListView2.SelectedItem.SubItems(2) Then
                    'Ha encontrado con el mismo importe y signos distintos D-H
                    Set ListView1.SelectedItem = ListView1.ListItems(I)
                    ListView1.SelectedItem.EnsureVisible
                    Find = True
                    Exit For
                End If
            End If
        Next I
        If Not Find Then
            If J > 1 Then
                Fin = J
                J = 1
            Else
                Find = True
            End If
        End If
                
    Loop Until Find

End Sub

Private Sub Text1_GotFocus()
    PonFoco Text1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Or KeyCode = 187 Then
        KeyCode = 0
        Text1.Text = ""
        imgCuentas_Click
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus()
Dim RC As String
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "+" Then Text1.Text = ""
    If Text1.Text = "" Then
        Text2.Text = ""
        Exit Sub
    Else
         RC = Text1.Text
         If CuentaCorrectaUltimoNivel(RC, SQL) Then
             Text1.Text = RC
             Text2.Text = SQL
         Else
             MsgBox SQL, vbExclamation
             Text1.Text = ""
             Text2.Text = ""
         End If
         If Text1.Text = "" Then PonerFoco Text1
         
    End If
             
End Sub


Private Sub PonerFoco(Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub Text11_GotFocus()
    PonFoco Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text4_GotFocus()
    PonFoco Text4
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Or KeyCode = 187 Then
        KeyCode = 0
        Text1.Text = ""
        Image3_Click
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text4_LostFocus()
Dim RC As String

    Text4.Text = Trim(Text4.Text)
    If Text4.Text = "+" Then Text4.Text = ""
    If Text4.Text = "" Then
        Text5.Text = ""
    Else
        RC = Text4.Text
        If CuentaCorrectaUltimoNivel(RC, SQL) Then
            Text4.Text = RC
            Text5.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text5.Text = ""
            Text4.Text = ""
            Text4.SetFocus
        End If
    End If
End Sub

Private Sub Text6_GotFocus()
    PonFoco Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text6_LostFocus()
   With Text6
        .Text = Trim(.Text)
        I = 1
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "El valor debe ser numérico: " & .Text, vbExclamation
            Else
                 If Val(.Text) >= 900 Then
                    MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
                Else
                    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", .Text, "N")
                    If SQL = "" Then
                        MsgBox "Concepto NO encontrado: " & .Text, vbExclamation
                    Else
                        Text7.Text = SQL
                        I = 0
                    End If
                End If
            End If
        Else
            'Igual a "" luego pasamos a otro campo en la tabulacion
            I = 2
        End If
        If I > 0 Then
            .Text = ""
            Text7.Text = ""
            If I = 1 Then Text6.SetFocus
        End If
    End With
End Sub

Private Sub Text8_GotFocus()
    PonFoco Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text9_GotFocus()
    PonFoco Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text9_LostFocus()
    With Text9
        .Text = Trim(.Text)
        I = 1
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "El valor debe ser numérico: " & .Text, vbExclamation
            Else
                SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", .Text, "N")
                If SQL = "" Then
                    MsgBox "Concepto NO encontrado: " & .Text, vbExclamation
                Else
                    Text10.Text = SQL
                    I = 0
                End If
            End If
        Else
            'Igual a "" luego pasamos a otro campo
            I = 2
        End If
        If I > 0 Then
            .Text = ""
            Text10.Text = ""
            If I = 1 Then Text9.SetFocus
        End If
    End With
End Sub

Private Sub txtfec_GotFocus(Index As Integer)
    PonFoco txtFec(Index)
End Sub

Private Sub txtfec_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfec_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtFec(Index).Text = Trim(txtFec(Index).Text)
    Mal = True
    If txtFec(Index).Text = "" Then
        MsgBox "Escriba una fecha correcta", vbExclamation
    Else
        If Not EsFechaOK(txtFec(Index)) Then
            MsgBox "No es una fecha correcta", vbExclamation
        Else
            Mal = False
        End If
    End If
    If Mal Then txtFec(Index).SetFocus
End Sub



Private Sub CargaBancario()

    ListView1.ListItems.Clear
    SQL = "Select * from norma43 where"
    SQL = SQL & " codmacta ='" & Text1.Text & "'"
    SQL = SQL & " AND fecopera >='" & Format(txtFec(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecopera <='" & Format(txtFec(1).Text, FormatoFecha) & "'"
    'OCultar/mostrar punteados
    If Check1.Value = 0 Then
        'Ocultar los ya puntedos
        SQL = SQL & " AND Punteada = 0 "
    End If
    SQL = SQL & " ORDER BY fecopera,codigo"

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = Format(RS!fecopera, "dd/mm/yy")
        'Importe Debe
        If Not IsNull(RS!ImporteD) Then
            Importe = RS!ImporteD
            SQL = "D"
        Else
            'Importe HABER
            If Not IsNull(RS!ImporteH) Then
                Importe = RS!ImporteH
                SQL = "H"
            Else
                SQL = "XX"
            End If
        End If
        ItmX.SubItems(1) = Format(Importe, FormatoImporte)
        ItmX.SubItems(2) = SQL
        ItmX.SubItems(3) = Format(RS!Saldo, FormatoImporte)
        ItmX.SubItems(4) = RS!Concepto
        ItmX.Tag = RS!Codigo
        ItmX.Checked = (RS!punteada = 1)
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
End Sub


Private Sub CargaLineaApuntes()

    ListView2.ListItems.Clear
    SQL = "Select numasien,fechaent,numdiari,linliapu,ampconce,timported,timporteh,punteada,saldo FROM tmpconext"
    SQL = SQL & " WHERE codusu = " & vUsu.Codigo
    
    If Check1.Value = 0 Then
        'Ocultar los ya puntedos
        SQL = SQL & " AND Punteada = '' "
    End If
    SQL = SQL & " ORDER BY pos"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Set ItmX = ListView2.ListItems.Add()
        ItmX.Text = Format(RS!fechaent, "dd/mm/yy")
        'Importe Debe
        SQL = " "
        If Not IsNull(RS!timported) Then
            Importe = Format(RS!timported, FormatoImporte)
            SQL = "D"
        Else
            'Importe HABER
            If Not IsNull(RS!timporteH) Then
                Importe = RS!timporteH
                SQL = "H"
            Else
                Importe = 0
                SQL = "XX"
            End If
        End If
        ItmX.SubItems(1) = Format(Importe, FormatoImporte)
        ItmX.SubItems(2) = SQL
        ItmX.SubItems(3) = Format(RS!Saldo, FormatoImporte)
        ItmX.SubItems(4) = RS!ampconce
        ItmX.Tag = RS!Numasien & "|" & RS!NumDiari & "|" & RS!Linliapu & "|"
        ItmX.Checked = (RS!punteada <> "")
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub




'----------------- PUNTEOS

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
EstaLW1 = True
Screen.MousePointer = vbHourglass
    Set ListView1.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, True
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 1 Then
            If PrimeraSeleccion Then
                BusquedaEnHaber
                PrimeraSeleccion = False
                ClickAnterior = 1
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault

End Sub


Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Screen.MousePointer = vbHourglass
    EstaLW1 = False
    Set ListView2.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, False
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 2 Then
            If PrimeraSeleccion Then
                BusquedaEnDebe
                PrimeraSeleccion = False
                ClickAnterior = 2
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub BusquedaEnHaber()
    ContadorBus = 1
    Checkear = False
    Do
        I = 1
        While I <= ListView2.ListItems.Count
            'Comprobamos k no esta chekeado
            If Not ListView2.ListItems(I).Checked Then
                'K tiene el mismo importe
                If ListView1.SelectedItem.SubItems(1) = ListView2.ListItems(I).SubItems(1) Then
                    'K no sean DEBE o HABER los dos
                    Checkear = (ListView1.SelectedItem.SubItems(2) <> ListView2.ListItems(I).SubItems(2))

                    If Checkear Then
                        'Tiene el mismo importe y no esta chequeado
                        Set ListView2.SelectedItem = ListView2.ListItems(I)
                        ListView2.SelectedItem.EnsureVisible
                        ListView2.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            I = I + 1
        Wend
        ContadorBus = ContadorBus + 1
        Loop Until ContadorBus > 2
End Sub



Private Sub BusquedaEnDebe()
    ContadorBus = 1
    Checkear = False
    Do
        I = 1
        While I <= ListView1.ListItems.Count
            If ListView2.SelectedItem.SubItems(1) = ListView1.ListItems(I).SubItems(1) Then
                'Lo hemos encontrado. Comprobamos que no esta chequeado
                If Not ListView1.ListItems(I).Checked Then
                    'Tiene el mismo importe y no son debe o haber
                    Checkear = (ListView2.SelectedItem.SubItems(2) <> ListView1.ListItems(I).SubItems(2))

                    If Checkear Then
                        Set ListView1.SelectedItem = ListView1.ListItems(I)
                        ListView1.SelectedItem.EnsureVisible
                        ListView1.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            I = I + 1
        Wend
        ContadorBus = ContadorBus + 1
    Loop Until ContadorBus > 2
End Sub



Private Sub PunteaEnBD(ByRef IT As ListItem, EnDEBE As Boolean)
Dim RC As String
On Error GoTo EPuntea
    
    
    If Not EnDEBE Then
        'ASientos
        'Actualizamos en DOS tablas, en la tmp y en la hcoapuntes
        SQL = "UPDATE hlinapu SET "
        If IT.Checked Then
            RC = "1"
            Importe = 1
            Else
            RC = "0"
            Importe = -1
        End If
        Importe = Importe * CSng(IT.SubItems(1))
        If EnDEBE Then
            De = De + Importe
        Else
            Ha = Ha + Importe
        End If
        SQL = SQL & " punteada = " & RC
        SQL = SQL & " WHERE fechaent='" & Format(IT.Text, FormatoFecha) & "'"
        SQL = SQL & " AND numasien="
        RC = RecuperaValor(IT.Tag, 1)
        SQL = SQL & RC & " AND numdiari ="
        RC = RecuperaValor(IT.Tag, 2)
        SQL = SQL & RC & " AND linliapu ="
        RC = RecuperaValor(IT.Tag, 3)
        SQL = SQL & RC
        
        
        
        
    Else
        'En Norma 43
        
        If IT.Checked Then
            RC = "1"
            Importe = 1
            Else
            RC = "0"
            Importe = -1
        End If
        Importe = Importe * CSng(IT.SubItems(1))
        If EnDEBE Then
            De = De + Importe
        Else
            Ha = Ha + Importe
        End If
        SQL = "UPDATE norma43 SET punteada= " & RC & " WHERE codigo=" & IT.Tag
        
    End If
    
    Conn.Execute SQL
    
    'Ponemos los importes
    PonerImportes

    
    Exit Sub
EPuntea:
    MuestraError Err.Number, "Accediendo BD para puntear", Err.Description
End Sub






Private Sub PonFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub



Private Function GenerarCabecera(NumAsi As Long) As Boolean
Dim Cad As String

    On Error GoTo EGenerarCabecera
    GenerarCabecera = False
    
    '-------------------------------------------------------------------------
    'Insertamos cabecera
    SQL = "INSERT INTO cabapue (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
    'Ejemplo
    ' 1, '2003-11-25', 1, 1, NULL, 'misobs')
    SQL = SQL & Text9.Text & ",'" & Format(CDate(txtFec(2).Text), FormatoFecha) & "'," & NumAsi & ",1,NULL,"
    'Observaciones
    SQL = SQL & "'Asiento generado desde punteo bancario por " & vUsu.Nombre & " el " & Format(Now, "dd/mm/yyyy") & "')"
    Conn.Execute SQL
    
    '-----------------------------------------------------------------------------
    'La linea del asiento
    'Hemos puesto linapu mas atras para poder cambiarla
    SQL = "INSERT INTO linapue (numdiari, fechaent, numasien, numdocum,"
    SQL = SQL & " ampconce, codconce, linliapu, codmacta, timporteD, timporteH, ctacontr, codccost, idcontab, punteada) VALUES ("
    
    'Ejemplo valores
    '1, '2001-01-20', 0, 0, '0', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0)"
    SQL = SQL & Text9.Text & ",'" & Format(CDate(txtFec(2).Text), FormatoFecha) & "'," & NumAsi & ",'"
    '          documento
    SQL = SQL & Text11.Text & "','"
    
    'Ampliacion concepto
    Cad = Mid(Text7.Text & " " & Text8.Text, 1, 30)
    SQL = SQL & DevNombreSQL(Cad) & "',"
    
    'Concepto
    SQL = SQL & Text6.Text & ","
    
    'El importe
    Importe = CCur(ListView1.SelectedItem.SubItems(1))
    Cad = "1,'" & Text1.Text & "',"
    If ListView1.SelectedItem.SubItems(2) = "H" Then
        'Va al debe
        Cad = Cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
    Else
        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Importe))
    End If
    
    'Contrapartida
    If Text4.Text <> "" Then
        Cad = Cad & ",'" & Text4.Text & "'"
    Else
        Cad = Cad & ",NULL"
    End If
    
    'y la punteamos
    Cad = SQL & Cad & ",NULL,'CONTAB',1)"
    Conn.Execute Cad
    
    'Si tiene contrapartida entonces genero la segunda linea de apuntes
    ' k sera la de la contrapartida, con el importe el mismo al lado contrario
    ' el mismo concepto
    If Text4.Text <> "" Then
        'SI TIENE
            Cad = "2,'" & Text4.Text & "',"
            'En la de arriba es igual a H
            If ListView1.SelectedItem.SubItems(2) = "D" Then
                'Va al debe
                Cad = Cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
            Else
                Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Importe))
            End If
            
            'Contrapartida es la del banco
            Cad = Cad & ",'" & Text1.Text & "'"
            
            'y NO la punteamos
            Cad = SQL & Cad & ",NULL,'CONTAB',0)"
            Conn.Execute Cad
    End If
    GenerarCabecera = True
    Exit Function
EGenerarCabecera:
    MuestraError Err.Number, Err.Description
End Function



Private Sub PonerImportes()

    If De <> 0 Then
        Text3(0).Text = Format(De, FormatoImporte)
        Else
        Text3(0).Text = ""
    End If
    If Ha <> 0 Then
        Text3(1).Text = Format(Ha, FormatoImporte)
        Else
        Text3(1).Text = ""
    End If
    Importe = De - Ha
    If Importe <> 0 Then
        Text3(2).Text = Format(Importe, FormatoImporte)
        Else
        Text3(2).Text = ""
    End If
End Sub
