VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBalances 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurador de balances"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "frmBalances2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameActivo 
      Height          =   6075
      Left            =   240
      TabIndex        =   6
      Top             =   540
      Width           =   10215
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   8640
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   5640
         Width           =   1395
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5115
         Left            =   7560
         TabIndex        =   12
         Top             =   300
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   9022
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   5640
         Width           =   4995
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   5595
         Width           =   735
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5115
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9022
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Con. Oficial"
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
         Left            =   7560
         TabIndex        =   36
         Top             =   5700
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Texto:"
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
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   5700
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo:"
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
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   5700
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9300
      TabIndex        =   20
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame FrameBotones 
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   6900
      Width           =   5775
      Begin VB.CommandButton cmdUpDown 
         Height          =   375
         Index           =   1
         Left            =   3120
         Picture         =   "frmBalances2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Bajar"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton cmdUpDown 
         Height          =   375
         Index           =   0
         Left            =   2640
         Picture         =   "frmBalances2.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Subir "
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   6
         Left            =   4920
         Picture         =   "frmBalances2.frx":0E1E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Agregar cuenta"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   5
         Left            =   5340
         Picture         =   "frmBalances2.frx":13A8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Eliminar cuenta"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   4
         Left            =   4500
         Picture         =   "frmBalances2.frx":1932
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Agregar cuenta"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   3
         Left            =   1740
         Picture         =   "frmBalances2.frx":1EBC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Eliminar nodo"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   2
         Left            =   1260
         Picture         =   "frmBalances2.frx":2446
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Modificar nodo"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   1
         Left            =   780
         Picture         =   "frmBalances2.frx":29D0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Agregar nodo hijo"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   0
         Left            =   300
         Picture         =   "frmBalances2.frx":2F5A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Agregar nodo"
         Top             =   180
         Width           =   375
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   11774
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmBalances2.frx":34E4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBal(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtBal(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtBal(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Check1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Check2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdGuardar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Activo"
      TabPicture(1)   =   "frmBalances2.frx":3500
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Pasivo"
      TabPicture(2)   =   "frmBalances2.frx":351C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Resumen Plan Contable"
      TabPicture(3)   =   "frmBalances2.frx":3538
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TreeView2"
      Tab(3).ControlCount=   1
      Begin MSComctlLib.TreeView TreeView2 
         Height          =   5895
         Left            =   -74700
         TabIndex        =   34
         Top             =   540
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   10398
         _Version        =   393217
         Indentation     =   1411
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Predeterminado"
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   4440
         Width           =   2355
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar cambios"
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   5220
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Balance de p�rdidas y ganancias"
         Height          =   255
         Left            =   6600
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Aparecen descripcion cuentas"
         Height          =   255
         Left            =   4020
         TabIndex        =   3
         Top             =   4440
         Width           =   2715
      End
      Begin VB.TextBox txtBal 
         Height          =   1815
         Index           =   2
         Left            =   3960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmBalances2.frx":3554
         Top             =   2400
         Width           =   5295
      End
      Begin VB.TextBox txtBal 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   1860
         Width           =   5295
      End
      Begin VB.TextBox txtBal 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text3"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   6
         Left            =   780
         TabIndex        =   31
         Top             =   5640
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   5
         Left            =   780
         TabIndex        =   30
         Top             =   4815
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   4
         Left            =   780
         TabIndex        =   29
         Top             =   3975
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   780
         TabIndex        =   28
         Top             =   3150
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   780
         TabIndex        =   27
         Top             =   2325
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   780
         TabIndex        =   26
         Top             =   1485
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   780
         TabIndex        =   25
         Top             =   660
         Width           =   435
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5895
         Left            =   120
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Descripci�n"
         Height          =   315
         Index           =   2
         Left            =   3060
         TabIndex        =   24
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   315
         Index           =   1
         Left            =   3060
         TabIndex        =   23
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "N�mero"
         Height          =   315
         Index           =   0
         Left            =   3060
         TabIndex        =   22
         Top             =   1380
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":355C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":97F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":A208
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":A522
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":AF34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":B4CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":10B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances2.frx":161B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuListview 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnInsertarNuevaCuenta 
         Caption         =   "Insertar cuenta"
         Begin VB.Menu mnInsertarSaldo 
            Caption         =   "Saldo"
         End
         Begin VB.Menu mnInsertarHaber 
            Caption         =   "Haber"
         End
         Begin VB.Menu mnInsertDebe 
            Caption         =   "Debe"
         End
      End
      Begin VB.Menu mnEliminarCuena 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu menuTree 
      Caption         =   "Tree"
      Visible         =   0   'False
      Begin VB.Menu mnNuevo 
         Caption         =   "Nuevo"
         Begin VB.Menu mnNuevoGrupo 
            Caption         =   "Elto. del grupo"
         End
         Begin VB.Menu mnInsertarSubGrupo 
            Caption         =   "Sub grupo"
         End
      End
      Begin VB.Menu mnEliminarGrupo 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mnDescLin 
         Caption         =   "Descripcion linea"
      End
      Begin VB.Menu mnbarra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnModificarTexto 
         Caption         =   "Modificar texto asociado"
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnArriba 
         Caption         =   "Hacia arriba"
      End
      Begin VB.Menu mnAbajo 
         Caption         =   "Hacia abajo"
      End
   End
End
Attribute VB_Name = "frmBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public numBalance As Integer


'Private WithEvents frmCta As frmColCtas


Dim PrimeraVez As Boolean
Dim SQL As String
Dim RS As Recordset
Dim NodoArbol As Node
Dim I As Integer
Dim Devolucion As String
Dim Clave As String



Private Sub OpcionesTAB()
Dim B As Boolean

    Me.frameActivo.Visible = Me.SSTab1.Tab = 1 Or Me.SSTab1.Tab = 2
    FrameBotones.Visible = frameActivo.Visible
    'Sin NumRegElim=0 entonces el balance ya existe
    'Si no es el numero k habra k asignarle
    B = (NumRegElim = 0)
    SSTab1.TabVisible(1) = B
    If B Then
        If Check2.Value = 1 Then
            'Si es perdidas i ganacias entonces , en el nuevo Plan solo hay una columna
            If vParam.NuevoPlanContable Then
                SSTab1.TabVisible(2) = False
            Else
                SSTab1.TabVisible(2) = True
            End If
        Else
            'Si es balance situacion
            SSTab1.TabVisible(2) = True
        End If
    Else
        SSTab1.TabVisible(2) = False
    End If
    'SSTab1.TabVisible(2) = (NumRegElim = 0)
End Sub


Private Sub Check2_Click()
    PonerTextosTab
End Sub

Private Sub cmdGuardar_Click()
On Error GoTo EGuarda

    If Val(txtBal(0).Text) > 50 Then
        MsgBox "No se puede guardar", vbExclamation
        Exit Sub
    End If

    txtBal(1).Text = Trim(txtBal(1).Text)
    txtBal(2).Text = Trim(txtBal(2).Text)
    If txtBal(1).Text = "" Then
        MsgBox "Nombre del balance no puede estar en blanco"
        Exit Sub
    End If
    
    If NumRegElim = 0 Then
    
        'Si esta como predeterminado
        
        If Check3.Value Then
            SQL = "UPDATE sbalan set predeterminado=0 where perdidas =" & Abs(CInt(Check2.Value))
            Conn.Execute SQL
        End If
        'Es modificar
        SQL = "UPDATE sbalan SET nombalan='" & txtBal(1).Text & "', Descripcion = "
        If txtBal(2).Text = "" Then
            Clave = "NULL"
        Else
            Clave = "'" & txtBal(2).Text & "'"
        End If
        SQL = SQL & Clave & ", Aparece =" & Abs(CInt(Check1.Value))
        SQL = SQL & ", perdidas =" & Abs(CInt(Check2.Value))
        SQL = SQL & ", predeterminado =" & Abs(CInt(Check3.Value))
        SQL = SQL & " WHERE numbalan =" & numBalance
        
        
        
    Else
        'Si no ha marcado predeterminado, y no existe ningun balance del tipo PREDETERMINADO, lo pongo a true
        If Check3.Value = 0 Then
            If Not ExistePredeterminado Then Check3.Value = 1
        End If
        'Es nuevo
        SQL = "INSERT INTO sbalan (numbalan, nombalan, Descripcion, Aparece, perdidas, Predeterminado) VALUES ("
        SQL = SQL & NumRegElim & ",'" & txtBal(1).Text & "',"
        If txtBal(2).Text = "" Then
            Clave = "NULL"
        Else
            Clave = "'" & txtBal(2).Text & "'"
        End If
        SQL = SQL & Clave & "," & Abs(CInt(Check1.Value)) & ","
        SQL = SQL & Abs(CInt(Check2.Value)) & "," & Abs(CInt(Check3.Value)) & ")"
    End If
    Conn.Execute SQL
    
    
    'AHora ponemos si se ha seleccionado como predetrminado, el resto lo ponemos como NO predeterminado
    If Check3.Value Then
        'Ha puesto perdeterminado. Entonces vamos a poner todos los demas
        'Balances del tipo de este
        SQL = "UPDATE sbalan SET predeterminado=0  WHERE Perdidas = " & Abs(CInt(Check2.Value))
        SQL = SQL & " AND numbalan <> "
        If NumRegElim > 0 Then
            SQL = SQL & NumRegElim
        Else
            SQL = SQL & numBalance
        End If
        Conn.Execute SQL
    End If
    
    
    'Si llega aqui es k ha sido con exito
    'entonces
    If NumRegElim > 0 Then
        numBalance = NumRegElim
        NumRegElim = 0
        OpcionesTAB
        Me.Refresh
    End If
    Exit Sub
EGuarda:
    MuestraError Err.Number, "Guardar encabezado balance"
End Sub

Private Sub cmdUpDown_Click(Index As Integer)
Dim LetraActivo As String
Dim N As Node
    
    If TreeView1.Nodes.Count = 0 Then Exit Sub
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    'Subir
    If Index = 0 Then
        'El primer hijo
        If TreeView1.SelectedItem.Previous Is Nothing Then
            MsgBox "Ya es el primero", vbExclamation
            Exit Sub
        End If
    Else
        
        If TreeView1.SelectedItem.Next Is Nothing Then
            MsgBox "Ya es el ultimo", vbExclamation
            Exit Sub
        End If
    End If

    Set RS = New ADODB.Recordset
    
    
    LetraActivo = SSTab1.Tag
    
    
    
    
    'CAMBIAR ORDEN
    '------------------------------------------
    
    SQL = "SELECT * from sperdid WHERE numbalan= " & numBalance
    SQL = SQL & " AND Pasivo = '" & LetraActivo & "' AND codigo = "
    Clave = ""
    If Index = 0 Then
        'Anterior
         Clave = Mid(TreeView1.SelectedItem.Previous.Key, 2)
    Else
         Clave = Mid(TreeView1.SelectedItem.Next.Key, 2)
    End If
    SQL = SQL & Clave
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Error leyendo NODO: " & Clave, vbExclamation
        RS.Close
        Exit Sub
    End If
    
    NumRegElim = RS!Orden  'TEngo el nodo destino
    RS.Close
    
    
    
    SQL = "SELECT * from sperdid WHERE numbalan= " & numBalance
    SQL = SQL & " AND Pasivo = '" & LetraActivo & "' AND codigo = "
    SQL = SQL & Mid(TreeView1.SelectedItem.Key, 2)
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Error leyendo NODO: " & TreeView1.SelectedItem.Key, vbExclamation
        RS.Close
        Exit Sub
    End If
    
    I = RS!Orden   'En I tengo el nodo que se mueve
    RS.Close
    
    If I = NumRegElim Then
        MsgBox "TIENEN EL MISMO ORDEN. Consulte a soporte tecnico", vbExclamation
    End If
    
    'Updateamos los nodos
    SQL = "UPDATE sperdid SeT orden = " & I
    SQL = SQL & " Where NumBalan = " & numBalance
    SQL = SQL & " AND Pasivo = '" & LetraActivo & "' AND codigo = "
    Clave = ""
    If Index = 0 Then
        'Anterior
         Clave = Mid(TreeView1.SelectedItem.Previous.Key, 2)
    Else
         Clave = Mid(TreeView1.SelectedItem.Next.Key, 2)
    End If
    SQL = SQL & Clave
    Conn.Execute SQL
    
    
    SQL = "UPDATE sperdid SeT orden = " & NumRegElim
    SQL = SQL & " Where NumBalan = " & numBalance
    SQL = SQL & " AND Pasivo = '" & LetraActivo & "' AND codigo = "
    SQL = SQL & Mid(TreeView1.SelectedItem.Key, 2)
    Conn.Execute SQL
    
    
    'Recargamos el nodo
    CadenaDesdeOtroForm = TreeView1.SelectedItem.Key
    
    If TreeView1.SelectedItem.Parent Is Nothing Then
        'Cargamos otravez todo
        CargaTree LetraActivo

        
        
    Else
        NumRegElim = Mid(TreeView1.SelectedItem.Parent.Key, 2)
        Set NodoArbol = TreeView1.SelectedItem.FirstSibling
        While Not (NodoArbol Is Nothing)
                
            Set N = NodoArbol.Next
            TreeView1.Nodes.Remove NodoArbol.Index
            Set NodoArbol = N
            
        Wend
                'En numregelim tenemos el codigo del padre de los que estamos moviendo
        Clave = NumRegElim & "|"
        SQL = "SELECT * from sperdid WHERE numbalan= " & numBalance
        SQL = SQL & " AND Pasivo = '" & LetraActivo & "' AND Padre "
        While Clave <> ""
            I = InStr(1, Clave, "|")
            If I = 0 Then Clave = ""
            If I > 0 Then
                Devolucion = Mid(Clave, 1, I - 1)
                Clave = Mid(Clave, I + 1)
                RS.Open SQL & " = " & Devolucion & " ORDER BY orden", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Devolucion = LetraActivo & Devolucion
                While Not RS.EOF
                    Clave = Clave & RS!Codigo & "|"
                    Set N = Me.TreeView1.Nodes.Add(Devolucion, tvwChild, LetraActivo & RS!Codigo, RS!deslinea)
                    If RS!Tipo = 1 Then
                         N.Tag = DBLet(RS!Formula) & "||"
                         N.Image = 3
                    Else
                         N.Tag = DBLet(RS!texlinea) & "|" & DBLet(RS!LibroCD) & "|"
                         N.Image = 2
                    End If
                    RS.MoveNext
                Wend
                RS.Close
            End If
            'Si llevamos muchos nodos refrescamos
            I = TreeView1.Nodes.Count Mod 25
            If I = 0 Then TreeView1.Refresh
        Wend
    
    
    
    
    
    
    
    
    



    
        
        
        
        
        
        
        
    
    End If
    
    
    'Dejamos seleccionado el que estaba
    For I = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(I).Key = CadenaDesdeOtroForm Then
            Set TreeView1.SelectedItem = TreeView1.Nodes(I)
            TreeView1.SelectedItem.EnsureVisible
            TreeView1.SelectedItem.Expanded = True
            Exit For
        End If
    Next I
    CargaListview
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    
    
    
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Tipo As String

    If Val(txtBal(0).Text) > 50 Then
        If Index < 4 Then
            MsgBox "No se puede editar", vbExclamation
            Exit Sub
        End If
    End If


    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0
        'Nuevo nodo desde el nivel del nodo seleccionado
        NuevoNodoNivel
    Case 1
        'Nodo hijo al nivel seleccionado
        'Nuevo nodo SUBNIVEL
        NuevoNodSub
    Case 2
        'Modificar valores del nodo
        ModificaNodo
    Case 3
        'Eliminar el nodo
        EliminarNODO
        
    Case 4, 6
        InsertarModificarCuenta (Index = 4)
    Case 5
        EliminarCta
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        'Cargamos el arbol de cuentas
        Screen.MousePointer = vbHourglass
        CargaPLAN
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    Me.SSTab1.Tab = 0
    Limpiar Me
    PonerDatosBalance
    OpcionesTAB
    
    'Alguna cosilla mas
    'Si es nuevo plan contable, las cuentas entran todas con saldo, con lo cual
    'el boton de modificar cuentas no estara. Estara a�adir y eliminar
    Command1(6).Visible = Not vParam.NuevoPlanContable
    
End Sub




Private Sub ListView1_DblClick()
    Command1_Click 6
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim LEtra As String

    OpcionesTAB
    Select Case SSTab1.Tab
    Case 1
        LEtra = "A"
    Case 2
        LEtra = "B"
    Case Else
        LEtra = ""
    End Select
    SSTab1.Tag = LEtra
    If LEtra <> "" Then
        If PreviousTab <> SSTab1.Tab Then CargaTree LEtra
    End If
    
End Sub


Private Sub CargaTree(LetraActivo As String)
Dim Nod As Node

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    SQL = "SELECT * from sperdid WHERE numbalan= " & numBalance
    SQL = SQL & " AND Pasivo = '" & LetraActivo & "' AND Padre "
    TreeView1.Nodes.Clear
    Set RS = New ADODB.Recordset
    RS.Open SQL & " is null ORDER BY Orden", Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If RS.EOF Then
        RS.Close
        GoTo Cerrar
    End If
    
    'Las raices
    
    Clave = ""
    While Not RS.EOF
        Clave = Clave & RS!Codigo & "|"
        Set Nod = Me.TreeView1.Nodes.Add(, , LetraActivo & RS!Codigo, RS!deslinea)
        Nod.Image = 1
        Nod.Tag = DBLet(RS!texlinea) & "|" & DBLet(RS!LibroCD) & "|"
        RS.MoveNext
    Wend
    RS.Close
    
    
    While Clave <> ""
        I = InStr(1, Clave, "|")
        If I = 0 Then Clave = ""
        If I > 0 Then
            Devolucion = Mid(Clave, 1, I - 1)
            Clave = Mid(Clave, I + 1)
            RS.Open SQL & " = " & Devolucion & " ORDER BY orden", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Devolucion = LetraActivo & Devolucion
            While Not RS.EOF
                Clave = Clave & RS!Codigo & "|"
                Set Nod = Me.TreeView1.Nodes.Add(Devolucion, tvwChild, LetraActivo & RS!Codigo, RS!deslinea)
                If RS!Tipo = 1 Then
                     Nod.Tag = DBLet(RS!Formula) & "||"
                     Nod.Image = 3
                Else
                     Nod.Tag = DBLet(RS!texlinea) & "|" & DBLet(RS!LibroCD) & "|"
                     Nod.Image = 2
                End If
                RS.MoveNext
            Wend
            RS.Close
        End If
        'Si llevamos muchos nodos refrescamos
        I = TreeView1.Nodes.Count Mod 25
        If I = 0 Then TreeView1.Refresh
    Wend
    
    
    
    If Not (Nod Is Nothing) Then
        Nod.EnsureVisible
    End If
    CargaListview
Cerrar:
    If Err.Number Then MuestraError Err.Number
    Set RS = Nothing
End Sub



Private Sub TreeView1_DblClick()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Children > 0 Then
       ' NO hacemos nada
    Else
        ModificaNodo
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Text1.Text = Node.Key
    Text2.Text = RecuperaValor(Node.Tag, 1)
    Text3.Text = RecuperaValor(Node.Tag, 2)
    CargaListview
End Sub

Private Sub NuevoNodoNivel()
Dim Nod As Node
Dim Sig As Integer
Dim Actual As String

    
    If TreeView1.SelectedItem Is Nothing Then
        Clave = ""
    Else
        Actual = Mid(TreeView1.SelectedItem.Key, 2)
        If TreeView1.SelectedItem.Parent Is Nothing Then
            Clave = ""
        Else
            Clave = Mid(TreeView1.SelectedItem.Parent.Key, 2)
        End If
    End If
    Sig = DevuelveSiguiente(Clave)  'Obtiene el siguiente y lo guarda en I

    'Ahora insertamos el nodo
    If Clave = "" Then
        Set Nod = TreeView1.Nodes.Add(, , SSTab1.Tag & Sig, "Nodo nuevo")
        Nod.Image = 1
    Else
        Set Nod = TreeView1.Nodes.Add(SSTab1.Tag & Actual, tvwLast, SSTab1.Tag & Sig, "Nodo nuevo")
        Nod.Image = 2
    End If
    Nod.EnsureVisible
    Set TreeView1.SelectedItem = Nod
    'Numbal,pasivo,codigo,padre,orden
    Clave = numBalance & "|" & SSTab1.Tag & "|" & Sig & "|" & Clave & "|" & I & "|"
    InsertarNodo Nod, False
    Me.Refresh
End Sub



Private Function DevuelveSiguiente(Padre As String) As Integer
    Set RS = New ADODB.Recordset
    SQL = ") From sperdid where numbalan=" & numBalance & " AND pasivo = '"
    If SSTab1.Tab = 1 Then
        SQL = SQL & "A'"
    Else
        SQL = SQL & "B'"
    End If
    RS.Open "Select max(codigo" & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    DevuelveSiguiente = I + 1
    
    'Ahora comprobamos dentro del bloque k orden tiene
    If Padre <> "" Then
        SQL = SQL & " AND Padre =" & Padre
    Else
        SQL = SQL & " AND Padre is null"
    End If
    RS.Open "Select max(orden" & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    I = I + 1
    Set RS = Nothing
    
    
    
End Function


Private Sub NuevoNodSub()
Dim Nod As Node
Dim Sig As Integer
Dim Actual As String

    If TreeView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un nodo donde insertar el subnivel", vbExclamation
        Exit Sub
    End If
    
    Clave = Mid(TreeView1.SelectedItem.Key, 2)
    Sig = DevuelveSiguiente(Clave)  'Obtiene el siguiente y lo guarda en I, Y el orden
    
    Set Nod = TreeView1.Nodes.Add(SSTab1.Tag & Clave, tvwChild, SSTab1.Tag & Sig, "Nodo nuevo")
    Nod.Image = 2
    Nod.EnsureVisible
    Set TreeView1.SelectedItem = Nod
    
    Clave = numBalance & "|" & SSTab1.Tag & "|" & Sig & "|" & Clave & "|" & I & "|"
    InsertarNodo Nod, False
    Me.Refresh
End Sub


Private Sub InsertarNodo(ByRef Nod As Node, Modificar As Boolean)
    CadenaDesdeOtroForm = ""
    If Not Modificar Then Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    ListView1.ListItems.Clear
    frmMensajes.Parametros = Clave
    frmMensajes.Opcion = 7 + Abs(CInt(Modificar))
    frmMensajes.Show vbModal
    If CadenaDesdeOtroForm = "" Then
        'Ha cancelado
        If Not Modificar Then TreeView1.Nodes.Remove Nod.Key
        If Not TreeView1.SelectedItem Is Nothing Then TreeView1_NodeClick TreeView1.SelectedItem
    Else
        'Ha insertado
        'Devuelve el texto, el texto auxiliar, y si es formula o no
        Nod.Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        Nod.Tag = RecuperaValor(CadenaDesdeOtroForm, 2) & "|" & RecuperaValor(CadenaDesdeOtroForm, 4) & "|"
        Text2.Text = RecuperaValor(Nod.Tag, 1)
        Text3.Text = RecuperaValor(Nod.Tag, 2)
        
        Clave = RecuperaValor(CadenaDesdeOtroForm, 3) 'Si es formula
        If Clave = "1" Then
            I = 3  'Icono suma
        Else
            'Si es padre habra k ver k icono le corresponde
            If Nod.Parent Is Nothing Then
                I = 1   'Icono raiz
            Else
                I = 2   'Icono noraml
            End If
        End If
        Nod.Image = I
    End If
End Sub
    

Private Sub EliminarNODO()
    If TreeView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un nodo para eliminar", vbExclamation
        Exit Sub
    End If
    
    
    If TreeView1.SelectedItem.Children > 0 Then
        MsgBox "Debe elmininar primero los nodos hijo", vbExclamation
        Exit Sub
    End If
    
    'Deberiamos comprobar si el nodo es parte de una formula
    
    
    'preguntamos
    SQL = "Seguro que desea eliminar el nodo: " & TreeView1.SelectedItem.Text & "?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    'Comun para borrar
    Clave = " numbalan = " & numBalance & " AND Pasivo ='" & SSTab1.Tag & "' AND Codigo = " & Mid(TreeView1.SelectedItem.Key, 2)
    SQL = "Delete from sperdi2 where " & Clave
    Conn.Execute SQL
    SQL = "Delete from sperdid where " & Clave
    Conn.Execute SQL
    
    'Updateamos los ordenes
    TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
    
End Sub




Private Sub ModificaNodo()

    If TreeView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un nodo para modificar", vbExclamation
        Exit Sub
    End If

    Set RS = New ADODB.Recordset
    Clave = " numbalan = " & numBalance & " AND Pasivo ='" & SSTab1.Tag & "' AND Codigo = " & Mid(TreeView1.SelectedItem.Key, 2)
    SQL = "Select * from sperdid where " & Clave
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Se ha generado un error leyendo el codigo : " & Clave, vbExclamation
        RS.Close
        Exit Sub
    End If
    
    Clave = ""
    For I = 0 To RS.Fields.Count - 1
        Clave = Clave & DBLet(RS.Fields(I)) & "|"
    Next I
    RS.Close
    Set RS = Nothing
    
    'Ya tengo todos los valores aqui, en Clave
    'Si tiene cuentas no va a poder generar formula
    InsertarNodo TreeView1.SelectedItem, True
               
End Sub



Private Sub InsertarModificarCuenta(Insertar As Boolean)
Dim Resta As Byte
    On Error GoTo EInsertarEliminarCuenta
    
    
    If TreeView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un nodo donde insertar la cuenta", vbExclamation
        Exit Sub
    End If
    
    'Si el nodo es de formula, NO, no puede inserart cuentas
    If TreeView1.SelectedItem.Image = 3 Then
        MsgBox "El nodo es un campo de formula.", vbExclamation
        Exit Sub
    End If
    
    If Not Insertar Then
        If ListView1.SelectedItem Is Nothing Then
            MsgBox "Seleccione una cuenta para modificar", vbExclamation
            Exit Sub
        End If
    End If
    
    
    CadenaDesdeOtroForm = ""
    SQL = TreeView1.SelectedItem.Text & "|"
    If Insertar Then
        I = 9
    Else
        Select Case ListView1.SelectedItem.Icon
        Case 5
            I = 1  'Debe en el siguiente form
        Case 6
            I = 2   'Haber
        Case Else
            I = 0   'Saldo
        End Select
                
        'Cuenta , debe haber
        SQL = SQL & ListView1.SelectedItem.Text & "|" & I & "|"
        
        
        'Si es resta
        SQL = SQL & ListView1.SelectedItem.Tag & "|"
        
        'Para la opcion del formulario
        I = 10
    End If
    frmMensajes.Opcion = I
    frmMensajes.Parametros = SQL
    frmMensajes.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Devolucion = RecuperaValor(CadenaDesdeOtroForm, 1)
        Clave = RecuperaValor(CadenaDesdeOtroForm, 2)
        SQL = RecuperaValor(CadenaDesdeOtroForm, 3)
        Resta = CByte(SQL)
        If Insertar Then
            If Not CompruebaCuenta(Devolucion, Clave, SQL = "1") Then Exit Sub
        End If
        InsertaModiCtaSQL Clave, Insertar, Resta
        CargaListview
    End If
    
    
    Exit Sub
EInsertarEliminarCuenta:
    MuestraError Err.Number, "Insertar eliminar cuenta"
End Sub


Private Sub InsertaModiCtaSQL(ByRef Tipo As String, Insertar As Boolean, Resta As Byte)
    On Error Resume Next
    
    If Insertar Then
        SQL = "INSERT INTO sperdi2 (NumBalan, Pasivo, codigo, codmacta, tipsaldo,Resta ) VALUES ("
        SQL = SQL & numBalance & ",'" & SSTab1.Tag & "'," & Mid(TreeView1.SelectedItem.Key, 2)
        SQL = SQL & ",'" & Devolucion & "','" & Tipo & "'," & Resta & ")"
        Conn.Execute SQL
        'Updateamos el registro para indicar k tiene cuentas
        SQL = "UPDATE sperdid Set TienenCtas=1 WHERE numbalan=" & numBalance & " AND Pasivo ='" & SSTab1.Tag
        SQL = SQL & "' AND codigo =" & Mid(TreeView1.SelectedItem.Key, 2)
        Conn.Execute SQL
        
    Else
        'MODIFICAR
        SQL = "UPDATE sperdi2 SET tipsaldo = '" & Tipo & "' ,Resta = " & Resta
        SQL = SQL & " WHERE numbalan ="
        SQL = SQL & numBalance & " AND pasivo = '" & SSTab1.Tag & "' AND codigo = " & Mid(TreeView1.SelectedItem.Key, 2)
        SQL = SQL & " AND codmacta = '" & Devolucion & "'"
        Conn.Execute SQL
    End If
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando cuenta"
    End If
End Sub


'LA cuenta viene empipamada
Private Function CompruebaCuenta(ByRef Cta As String, LEtra As String, Resta As Boolean) As Boolean
Dim C1 As String

    SQL = "Select * from sperdi2 WHERE numbalan= " & numBalance
    SQL = SQL & " AND codmacta='" & Cta & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RS.EOF
    
    
        If RS!Resta = Abs(Resta) Then
        
            If LEtra = RS!TipSaldo Then
                C1 = "pasivo = '" & RS!Pasivo & "' AND numbalan = " & numBalance & " AND codigo "
                C1 = DevuelveDesdeBD("deslinea", "sperdid", C1, RS!Codigo)
                C1 = RS!Pasivo & " - " & RS!Codigo & " - " & RS!TipSaldo & ": " & C1
                SQL = SQL & C1 & vbCrLf
            End If
            
            
        Else
            
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    If SQL <> "" Then
        SQL = "La cuenta ya est� en los registros: " & vbCrLf & SQL & vbCrLf & "Desea continuar igualmente?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
            CompruebaCuenta = True
        Else
            CompruebaCuenta = False
        End If
    Else
        CompruebaCuenta = True
    End If
        
End Function



Private Sub CargaListview()
Dim ItmX As ListItem
    ListView1.ListItems.Clear
    If TreeView1.SelectedItem Is Nothing Then Exit Sub

'Comun

    SQL = "Select * from sperdi2 WHERE numbalan= " & numBalance
    SQL = SQL & " AND Pasivo = '" & Me.SSTab1.Tag & "' AND codigo = " & Mid(TreeView1.SelectedItem.Key, 2)
    SQL = SQL & " ORDER BY codmacta"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add(, , RS!codmacta)   ' Autor.
        Select Case RS!TipSaldo
        Case "D"
             ItmX.Icon = 5
        Case "H"
             ItmX.Icon = 6
        Case Else
             ItmX.Icon = 7
        End Select
        ItmX.Tag = Abs(RS!Resta)
        If ItmX.Tag = 1 Then ItmX.ForeColor = vbRed
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
End Sub



Private Sub PonerDatosBalance()
    
    If NumRegElim = 0 Then
        'El balance ya existe
        'Ponemos los datos
        Set RS = New ADODB.Recordset
        RS.Open "Select * from Sbalan where numbalan=" & numBalance, Conn, adOpenDynamic, adLockPessimistic, adCmdText
        If RS.EOF Then
            MsgBox "Error leyendo los datos del balance: " & numBalance, vbExclamation
        Else
            txtBal(1).Text = RS!nombalan
            txtBal(2).Text = DBLet(RS!Descripcion)
            Check1.Value = RS!Aparece
            Check2.Value = RS!perdidas
            Check3.Value = RS!predeterminado
            Check3.Tag = RS!predeterminado
            txtBal(0).Text = numBalance
        End If
        RS.Close
        Set RS = Nothing
        Caption = "Configurador balances  - " & txtBal(1).Text
        
        
        'Los text
        
        
    Else
        'Balance nuevo
        txtBal(0).Text = NumRegElim
        txtBal(1).Text = ""
        txtBal(2).Text = ""
        Caption = Caption & "     (NUEVO)"
        Check3.Value = 0
    End If
    PonerTextosTab
End Sub

Private Sub PonerTextosTab()
    If Check2.Value Then
        If vParam.NuevoPlanContable Then
            Me.SSTab1.TabCaption(1) = "P y G"
        Else
            Me.SSTab1.TabCaption(1) = "Debe"
            Me.SSTab1.TabCaption(2) = "Haber"
        End If
    Else
        If vParam.NuevoPlanContable Then
            Me.SSTab1.TabCaption(2) = "Patrimonio neto y pasivo"
        Else
            Me.SSTab1.TabCaption(2) = "Pasivo"
        End If
        Me.SSTab1.TabCaption(1) = "Activo"
    End If
End Sub

Private Sub EliminarCta()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    SQL = "Desea eliminar de la linea la cuenta: " & ListView1.SelectedItem.Text & " ?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    Clave = " numbalan = " & numBalance & " AND Pasivo ='" & SSTab1.Tag & "' AND Codigo = " & Mid(TreeView1.SelectedItem.Key, 2)
    SQL = "DELETE FROM sperdi2 WHERE codmacta ='" & ListView1.SelectedItem.Text & "' AND "
    SQL = SQL & Clave
    Conn.Execute SQL
    
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
    'Si no kedan mas cuentas
    If ListView1.ListItems.Count = 0 Then
        'Updateamos el registro para indicar k tiene cuentas
        SQL = "UPDATE sperdid Set TienenCtas=0 WHERE numbalan=" & numBalance & " AND Pasivo = '" & SSTab1.Tag
        SQL = SQL & "' AND codigo =" & Mid(TreeView1.SelectedItem.Key, 2)
        Conn.Execute SQL
    End If
End Sub
    



Private Function ExistePredeterminado() As Boolean
    
On Error GoTo EExiste
    Set RS = New ADODB.Recordset
    ExistePredeterminado = False
    SQL = "Select * from sbalan where perdidas = " & Abs(Check2.Value)
    SQL = SQL & " AND predeterminado = 1"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then ExistePredeterminado = True
    RS.Close
    
EExiste:
    If Err.Number <> 0 Then Err.Clear
    Set RS = Nothing
End Function







Private Sub CargaPLAN()
Dim J As Integer
Dim N As Integer
    
    On Error GoTo ECargaPlan
    Set RS = New ADODB.Recordset
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        Devolucion = Mid("__________", 1, J)
        SQL = "Select codmacta,nommacta from Cuentas where codmacta like '" & Devolucion & "'"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If I = vEmpresa.numnivel - 1 Then
            N = 3
        Else
            N = 2
        End If
        While Not RS.EOF
            Clave = RS!codmacta & " - " & RS!nommacta
            If J > 1 Then
                N = DigitosNivel(I - 1)
                Devolucion = "C" & Mid(RS!codmacta, 1, N)
                Set NodoArbol = Me.TreeView2.Nodes.Add(Devolucion, tvwChild, "C" & RS!codmacta, Clave)
                If J < 4 Then
                    NodoArbol.Image = 2
                Else
                    NodoArbol.Image = 4
                End If
            Else
                Set NodoArbol = Me.TreeView2.Nodes.Add(, , "C" & RS!codmacta, Clave)
                NodoArbol.Image = 1
            End If
            'Todos los subnodos de segundo nivel los vamos a mostrar
            If I = 2 Then NodoArbol.EnsureVisible
            
            RS.MoveNext
        Wend
        RS.Close
    Next I
    Set RS = Nothing
    'Que muestre la primera cuenta
    If TreeView2.Nodes.Count > 0 Then TreeView2.Nodes(1).EnsureVisible
    Exit Sub
ECargaPlan:
    MuestraError Err.Number, "Cargando plan: " & RS!codmacta
End Sub
