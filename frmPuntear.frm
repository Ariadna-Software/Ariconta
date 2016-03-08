VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPuntear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punteo de extractos"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmPuntear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePorImportes 
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   7920
      Width           =   11655
      Begin VB.CommandButton cmdPorIMportes 
         Caption         =   "Aceptar cambios"
         Height          =   435
         Index           =   1
         Left            =   8760
         TabIndex        =   36
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdPorIMportes 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   435
         Index           =   0
         Left            =   10560
         TabIndex        =   35
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Leyendo datos"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   37
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "Punteo automatico por importes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame framePregunta 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton Option1 
         Caption         =   "Documento"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   6
         Top             =   4680
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ampliación"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   4680
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CheckBox chkSin 
         Caption         =   "Incluir sólo apuntes sin punteo"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CheckBox chkImp 
         Caption         =   "Ordenar apuntes por importe"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   4080
         Width           =   2415
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox txtDesCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1440
         Width           =   3915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1455
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         Height          =   615
         Left            =   240
         Top             =   3840
         Width           =   5415
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Left            =   1080
         Picture         =   "frmPuntear.frx":030A
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha inicio"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   2565
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmPuntear.frx":6B5C
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha fin"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "frmPuntear.frx":6BE7
         Top             =   2520
         Width           =   240
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   240
         Top             =   2280
         Width           =   5415
      End
      Begin VB.Shape Shape2 
         Height          =   1335
         Left            =   240
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label9 
         Caption         =   "Punteo de extractos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   7620
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "Text2"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   10380
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ListView lwD 
      Height          =   6555
      Left            =   120
      TabIndex        =   20
      Top             =   1500
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11562
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Asiento"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ampliación"
         Object.Width           =   3649
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Documento"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.ListView lwh 
      Height          =   6555
      Left            =   6000
      TabIndex        =   22
      Top             =   1500
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11562
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
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Asiento"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ampliación"
         Object.Width           =   3649
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Documento"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Seleccion datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cta anterior"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cta siguiente"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Saldos"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "IMPORTES del punteo"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar marcas de punteado"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Punteo automático por importes"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   3
      Left            =   7680
      Picture         =   "frmPuntear.frx":6C72
      ToolTipText     =   "Puntear al haber"
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   2
      Left            =   7320
      Picture         =   "frmPuntear.frx":6DBC
      ToolTipText     =   "Quitar al haber"
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   1800
      Picture         =   "frmPuntear.frx":6F06
      ToolTipText     =   "Puntear al Debe"
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   1440
      Picture         =   "frmPuntear.frx":7050
      ToolTipText     =   "Quitar al Debe"
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "DEBE"
      Height          =   255
      Index           =   2
      Left            =   7620
      TabIndex        =   31
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "HABER"
      Height          =   255
      Index           =   3
      Left            =   9000
      TabIndex        =   30
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "SALDO"
      Height          =   255
      Index           =   4
      Left            =   10380
      TabIndex        =   29
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Fechas"
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblHaber 
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
      Left            =   6000
      TabIndex        =   24
      Top             =   8160
      Width           =   5700
   End
   Begin VB.Label lblDebe 
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
      Left            =   165
      TabIndex        =   23
      Top             =   8160
      Width           =   5700
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "HABER"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DEBE"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Cuenta"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblFecha 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmPuntear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EjerciciosCerrados As Boolean

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim RS As Recordset

Dim PrimeraSeleccion As Boolean
Dim ClickAnterior As Byte '0 Empezar 1.-Debe 2.-Haber
Dim I As Integer
Dim De As Currency
Dim Ha As Currency
    
Dim ModoPunteo As Byte
    '0- Punteo normal. El de toda la vida
    '1- Punteo automatico por importes
    
Dim CtasQueHaPunteado As String
    
'Con estas dos variables
Dim ContadorBus As Integer
Dim Checkear As Boolean


Private Sub chkImp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub chkSin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmdAceptar_Click()
If Text1.Text = "" Then
    MsgBox "Introduzca una cuenta", vbExclamation
    PonleFoco Text1
    Exit Sub
End If
If Text3(0).Text = "" Or Text3(1).Text = "" Then
    MsgBox "Introduce las fechas de consulta de extractos", vbExclamation
    Exit Sub
End If
If Text3(0).Text <> "" And Text3(1).Text <> "" Then
    If CDate(Text3(0).Text) > CDate(Text3(1).Text) Then
        MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
        Exit Sub
    End If
End If

lblFecha.Caption = ""
SQL = ""
'Llegados aqui. Vemos la fecha y demas
If Text3(0).Text <> "" Then
    SQL = " fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    lblFecha.Caption = Text3(0).Text
End If

If Text3(1).Text <> "" Then
    If SQL <> "" Then SQL = SQL & " AND "
    SQL = SQL & " fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    If lblFecha.Caption = "" Then
        lblFecha.Caption = "hasta "
    Else
        lblFecha.Caption = lblFecha.Caption & " - "
    End If
    lblFecha.Caption = lblFecha.Caption & Text3(1).Text
End If
Text3(0).Tag = SQL  'Para las fechas

'Bloqueamos manualamente la tabla, con esa cuenta
If Not BloqueoManual(True, "PUNTEO", CStr(Abs(EjerciciosCerrados) & Text1.Text)) Then
    MsgBox "Imposible acceder a puntear la cuenta. Puede estar bloqueada", vbExclamation
    Exit Sub
End If
'Ponemos la cuenta
Text4.Text = Text1.Text
Text5.Text = txtDesCta.Text
Me.framePregunta.Visible = False
PonerTamaños True


espera 0.1
Me.Refresh
DoEvents
CargarDatos

End Sub


Private Sub cmdPorIMportes_Click(Index As Integer)

    If Index = 1 And AlgunNodoPunteado Then
        If MsgBox("¿Actualizar el punteo en la base de datos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
        

    

    If Index = 1 Then
        'Actualizamos la BD
        'Pongo ModoPunteo=0 para que ejecute el SQL
        ModoPunteo = 0
        
        'UPDATEAMOS EN LA BD
        'Y volveremos a cargar los datos
        For I = 1 To lwh.ListItems.Count
            If lwh.ListItems(I).Checked Then PunteaEnBD lwh.ListItems(I), False
        Next I
        
        For I = 1 To lwD.ListItems.Count
            If lwD.ListItems(I).Checked Then PunteaEnBD lwD.ListItems(I), True
        Next I
        
        
    Else
        'Quit la seleccion
        For I = 1 To lwD.ListItems.Count
            If lwD.ListItems(I).Checked Then lwD.ListItems(I).Checked = False
        Next I
        For I = 1 To lwh.ListItems.Count
            If lwh.ListItems(I).Checked Then lwh.ListItems(I).Checked = False
        Next I
    End If
    'Limpiamos campos
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    De = 0: Ha = 0
    


    'Quitamos las posibles marcas
    PonerModoPunteo False
    
    If Index = 1 Then CargarDatos
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub




Private Sub PonerTamaños(Grande As Boolean)
If Grande Then
    AjustarPantalla Me
    If Screen.Height < 9500 Then
        Me.Height = Me.Height - 420
        Me.lwD.Height = 6200 ' Me.lwD.Height - 150
        Me.lwh.Height = 6200 'Me.lwh.Height - 150
        Me.lblDebe.Top = 7800 'Me.lblDebe.Top - 150
        Me.lblHaber.Top = 7800 'Me.lblHaber.Top - 150
        
    End If
Else
    Me.Top = 1200
    Me.Left = 1200
    Me.Width = Me.framePregunta.Width + 30
    Me.Height = Me.framePregunta.Height + 90
End If
End Sub


Private Sub Form_Load()
    Me.framePregunta.Visible = True
    Limpiar Me
    PrimeraSeleccion = True
    Caption = "Punteo de extractos"
    If EjerciciosCerrados Then
        I = -1
        Caption = Caption & "  EJERCICIOS CERRADOS"
    Else
        I = 0
    End If
    'La toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 18
        .Buttons(4).Image = 7
        .Buttons(5).Image = 8
        .Buttons(7).Image = 22
        .Buttons(8).Image = 23
        .Buttons(9).Image = 14
        .Buttons(11).Image = 10
        .Buttons(13).Image = 15
    End With
    FramePorImportes.Visible = False
    
    Text3(0).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
    Text3(1).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
    PonerTamaños False
    CtasQueHaPunteado = ""   'Parar cuando haga el unload
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    BloqueoManual False, "PUNTEO", Text1.Text
    VerLogPunteado
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CByte(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text1.Text = RecuperaValor(CadenaSeleccion, 1)
    txtDesCta.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub




Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    Image1(0).Tag = Index
    If Text3(Index).Text <> "" Then
        frmC.Fecha = CDate(Text3(Index).Text)
    Else
        frmC.Fecha = Now
    End If
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub Image2_Click()
    BloqueoManual False, "PUNTEO", Text1.Text
    PonerTamaños False
    Me.framePregunta.Visible = True
    Me.Text1.Text = Text4.Text
    Me.txtDesCta.Text = Text5.Text
    PrimeraSeleccion = True
    PonleFoco Text1
End Sub



Private Sub Image4_Click()
    OtraCuenta True
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If (Index Mod 2) = 0 Then
        RC = "quitar punteos de lo apuntes"
    Else
        RC = "puntear los apuntes"
    End If
    
    SQL = "Seguro que desea " & RC
    If Index > 1 Then
        RC = "HABER"
    Else
        RC = "DEBE"
    End If
    SQL = SQL & " al " & RC & "?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    'HA DICHO SI
    
    If Index < 2 Then
        'PUNTEAMOS o DESPUNTEAMOS EL DEBE
        '---------------------------------
        Checkear = True
        If Index = 1 Then Checkear = False
        For I = 1 To lwD.ListItems.Count
            lblDebe.Caption = lwD.ListItems(I).Text
            lblDebe.Refresh
            If lwD.ListItems(I).Checked = Checkear Then
                    
                lwD.ListItems(I).Checked = Not lwD.ListItems(I).Checked
                PunteaEnBD lwD.ListItems(I), True
            End If
        Next I
        If Index = 0 Then De = 0
        If Not (lwD.SelectedItem Is Nothing) Then
            lblDebe.Caption = "Doc: " & lwD.SelectedItem.SubItems(4)
        Else
            lblDebe.Caption = ""
        End If
    Else
    
        'PUNTEAMOS o DESPUNTEAMOS EL HABER
        '---------------------------------
        Checkear = True
        If Index = 3 Then Checkear = False
        For I = 1 To lwh.ListItems.Count
            lblHaber.Caption = lwh.ListItems(I).Text
            lblHaber.Refresh
            If lwh.ListItems(I).Checked = Checkear Then
                lwh.ListItems(I).Checked = Not lwh.ListItems(I).Checked
                PunteaEnBD lwh.ListItems(I), False
            End If
        Next I
        If Index = 2 Then Ha = 0
        If Not (lwh.SelectedItem Is Nothing) Then
            lblHaber.Caption = "Doc: " & lwh.SelectedItem.SubItems(4)
        Else
            lblHaber.Caption = ""
        End If
    End If
    ContadorBus = 0
'    CargarDatos
    
    
    
    If De - Ha <> 0 Then
        Text2(2).Text = Format(De - Ha, FormatoImporte)
    Else
        Text2(2).Text = ""
    End If
    
    
    
End Sub

Private Sub imgCuentas_Click()
    Set frmCta = New frmColCtas
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3  'NUEVO
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub



Private Sub lwD_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Screen.MousePointer = vbHourglass
    Set lwD.SelectedItem = Item
    lblDebe.Caption = "Doc: " & Item.SubItems(4)
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


Private Sub lwD_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblDebe.Caption = "Doc: " & Item.SubItems(4)
End Sub

Private Sub lwh_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Screen.MousePointer = vbHourglass
    Set lwh.SelectedItem = Item
    lblHaber.Caption = "Doc: " & Item.SubItems(4)
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

Private Sub lwh_ItemClick(ByVal Item As MSComctlLib.ListItem)
lblHaber.Caption = "Doc: " & Item.SubItems(4)
End Sub



Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus()

    
    RC = Trim(Text1.Text)
    If RC = "" Then
        txtDesCta.Text = ""
        Exit Sub
    End If
    If CuentaCorrectaUltimoNivel(RC, SQL) Then
        Text1.Text = RC
        txtDesCta.Text = SQL
    Else
        MsgBox SQL, vbExclamation
        Text1.Text = ""
        txtDesCta.Text = ""
        'Text1.SetFocus
        PonleFoco Text1
    End If
End Sub



Private Sub Text3_GotFocus(Index As Integer)
With Text3(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub OtraCuenta(Mas As Boolean)
    Screen.MousePointer = vbHourglass
    Text5.Text = "OBTENER CUENTA"
    Text5.Refresh
    If ObtenerCuenta(Mas) Then
        BloqueoManual False, "PUNTEO", Text1.Text
        'Ponemos en text1 y 2 los valores de la nueva cuenta
        Text1.Text = Text4.Text
        txtDesCta.Text = 2

    
        'Ya tenemos la cuenta
        If Not BloqueoManual(True, "PUNTEO", CStr(Abs(EjerciciosCerrados) & Text1.Text)) Then
            MsgBox "Imposible acceder a puntear la cuenta. Puede que este bloqueada", vbExclamation
            Image2_Click
            Exit Sub
        End If
        
        CargarDatos
    Else
        Text5.Text = ""
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index).Text, vbExclamation
        Text3(Index).Text = ""
        PonleFoco Text3(Index)
    End If
End Sub





Private Sub CargarDatos()
        Label5.Caption = "CARGA"
        Label6.Caption = "CARGA"
        Me.Refresh
        Screen.MousePointer = vbHourglass
        CargarDatos2
        Screen.MousePointer = vbDefault
        Label5.Caption = "DEBE"
        Label6.Caption = "HABER"
        Label5.Refresh
        Label6.Refresh
End Sub

Private Sub CargarDatos2()
Dim ItmX As ListItem
On Error GoTo ECargarDatos

    'Deberiamos bloquear la cuenta en punteos, es decir
    'en alguna tabla poner que se esta punteando la cuenta X


    'Limpiamos listview
    lwD.ListItems.Clear
    lwh.ListItems.Clear
    DoEvents
    
    If Option1(0).Value Then
        RC = "Ampliación"
    Else
        RC = "Numdocum"
    End If
    lwD.ColumnHeaders(3).Text = RC
    lwh.ColumnHeaders(3).Text = RC
    
    'Y label
    lblDebe.Caption = ""
    lblHaber.Caption = ""
    
    'Resetamos importes punteados
    De = 0
    Ha = 0
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    
    Set RS = New ADODB.Recordset
    
    RC = "SELECT numdiari, linliapu, fechaent, numasien, ampconce"
    If Option1(0).Value Then RC = RC & " as C1"
        
    RC = RC & ",timporteD, timporteH,punteada,numdocum"
    If Not Option1(0).Value Then RC = RC & " as C1"
    RC = RC & " FROM hlinapu"
    If EjerciciosCerrados Then RC = RC & "1"
    
    RC = RC & " WHERE "
    RC = RC & " codmacta ='" & Me.Text4.Text & "' AND "
    RC = RC & Text3(0).Tag
    'Si solo mostramos los sin puntear
    If chkSin.Value = 1 Then RC = RC & " AND punteada =0 "
    'Si esta marcado ordenar por importe o no
    If Me.chkImp.Value = 1 Then
        RC = RC & " ORDER BY timported desc,timporteh desc "
    Else
        RC = RC & " ORDER BY fechaent"
    End If
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    While Not RS.EOF
        
        If IsNull(RS!timported) Then
            'Va al haber
            Set ItmX = lwh.ListItems.Add()
            ItmX.SubItems(3) = Format(RS!timporteH, FormatoImporte)
        Else
            'AL DEBE
            Set ItmX = lwD.ListItems.Add()
            ItmX.SubItems(3) = Format(RS!timported, FormatoImporte)
        End If
        ItmX.Text = Format(RS!fechaent, "dd/mm/yyyy")
        
        ItmX.SubItems(1) = RS!Numasien
        'ItmX.SubItems(2) = RS!ampconce
        ItmX.SubItems(2) = RS!C1
        ItmX.SubItems(4) = RS.Fields(8)
        
        'En el tag, para actualizaciones i demas pondremos
        'Separado por pipes los valores de numdiari y linliapu
        'claves de la tabla hlinapu
        ItmX.Tag = RS!NumDiari & "|" & RS!Linliapu & "|"
        
        'El check
        ItmX.Checked = (RS!punteada = 1)
        
        
        
        'Siguiente
        RS.MoveNext
        I = I + 1
        'Por si refrescamos
        If I > 3000 Then
            I = 0
            Me.Refresh
        End If
    Wend
    RS.Close

    If lwD.ListItems.Count > 0 Then lblDebe.Caption = "Doc: " & lwD.ListItems(1).SubItems(4)
    If lwh.ListItems.Count > 0 Then lblHaber.Caption = "Doc: " & lwh.ListItems(1).SubItems(4)
    Exit Sub
ECargarDatos:
        MuestraError Err.Number, "Cargando datos", Err.Description
        Set RS = Nothing
End Sub



'Private Sub CargarDatos3()
'Dim ItmX As ListItem
'On Error GoTo ECargarDatos
'
'    'Deberiamos bloquear la cuenta en punteos, es decir
'    'en alguna tabla poner que se esta punteando la cuenta X
'
'
'    'Limpiamos listview
'    lwT.ListItems.Clear
'
'
'    'Resetamos importes punteados
'    De = 0
'    Ha = 0
'    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
'
'    Set RS = New ADODB.Recordset
'    RC = "SELECT numdiari, linliapu, fechaent, numasien, ampconce,"
'    RC = RC & " timporteD, timporteH,punteada,numdocum FROM hlinapu"
'    If EjerciciosCerrados Then RC = RC & "1"
'
'    RC = RC & " WHERE "
'    RC = RC & " codmacta ='" & Me.Text4.Text & "' AND "
'    RC = RC & Text3(0).Tag
'    'Si solo mostramos los sin puntear
'    If chkSin.Value = 1 Then RC = RC & " AND punteada =0 "
'    'Si esta marcado ordenar por importe o no
'    If Me.chkImp.Value = 1 Then
'        RC = RC & " ORDER BY timported desc,timporteh desc "
'    Else
'        RC = RC & " ORDER BY fechaent"
'    End If
'
'    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    i = 0
'    While Not RS.EOF
'
'        Set ItmX = lwT.ListItems.Add()
'
'        If IsNull(RS!timported) Then
'            'Va al haber
'
'            ItmX.SubItems(3) = Format(RS!timporteH, FormatoImporte)
'        Else
'            'AL DEBE
'            ItmX.SubItems(3) = Format(RS!timported, FormatoImporte)
'        End If
'        ItmX.Text = Format(RS!fechaent, "dd/mm/yyyy")
'
'        ItmX.SubItems(1) = RS!Numasien
'        ItmX.SubItems(2) = RS!ampconce
'        ItmX.SubItems(4) = RS!numdocum
'
'        'En el tag, para actualizaciones i demas pondremos
'        'Separado por pipes los valores de numdiari y linliapu
'        'claves de la tabla hlinapu
'        ItmX.Tag = RS!NumDiari & "|" & RS!Linliapu & "|"
'
'        'El check
'        ItmX.Checked = (RS!punteada = 1)
'
'
'
'        'Siguiente
'        RS.MoveNext
'        i = i + 1
'        'Por si refrescamos
'        If i > 3000 Then
'            i = 0
'            Me.Refresh
'        End If
'    Wend
'    RS.Close
'
'
'    Exit Sub
'ECargarDatos:
'        MuestraError Err.Number, "Cargando datos", Err.Description
'        Set RS = Nothing
'End Sub




Private Sub BusquedaEnHaber()
    ContadorBus = 1
    Checkear = False
    Do
        I = 1
        While I <= lwh.ListItems.Count
            'Comprobamos k no esta chekeado
            If Not lwh.ListItems(I).Checked Then
                'K tiene el mismo importe
                If lwD.SelectedItem.SubItems(3) = lwh.ListItems(I).SubItems(3) Then
                    'Comprobamos k tienen el mismo DOCUM
                    'Si no es el mismo, pero es la segunda busqueda, tb aceptamos
                    If ContadorBus > 1 Then
                        Checkear = True
                    Else
                        Checkear = (lwD.SelectedItem.SubItems(4) = lwh.ListItems(I).SubItems(4))
                    End If
                
                    If Checkear Then
                        'Tiene el mismo importe y no esta chequeado
                        Set lwh.SelectedItem = lwh.ListItems(I)
                        lblHaber.Caption = "Doc: " & lwh.ListItems(I).SubItems(4)
                        lwh.SelectedItem.EnsureVisible
                        lwh.SetFocus
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
        While I <= lwD.ListItems.Count
            If lwh.SelectedItem.SubItems(3) = lwD.ListItems(I).SubItems(3) Then
                'Lo hemos encontrado. Comprobamos que no esta chequeado
                If Not lwD.ListItems(I).Checked Then
                    'Tiene el mismo importe y no esta chequeado
                    'Comprobamos k el docum es el mismo
                    'Si no es el mismo, pero es la segunda busqueda, tb aceptamos
                    If ContadorBus > 1 Then

                        Checkear = True
                    Else
                        Checkear = (lwh.SelectedItem.SubItems(4) = lwD.ListItems(I).SubItems(4))
                    End If
                    If Checkear Then
                        Set lwD.SelectedItem = lwD.ListItems(I)
                        lblDebe.Caption = "Doc: " & lwD.ListItems(I).SubItems(4)
                        lwD.SelectedItem.EnsureVisible
                        lwD.SetFocus
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
Dim Sng As Currency
On Error GoTo EPuntea
    
        
    
    SQL = "UPDATE hlinapu"
    If EjerciciosCerrados Then SQL = SQL & "1"
    SQL = SQL & " SET "
    If IT.Checked Then
        RC = "1"
        Sng = 1
        Else
        RC = "0"
        Sng = -1
    End If
    Sng = Sng * CCur(IT.SubItems(3))
    If EnDEBE Then
        De = De + Sng
    Else
        Ha = Ha + Sng
    End If




    SQL = SQL & " punteada = " & RC
    SQL = SQL & " WHERE fechaent='" & Format(IT.Text, FormatoFecha) & "'"
    SQL = SQL & " AND numasien=" & IT.SubItems(1) & " AND numdiari ="
    RC = RecuperaValor(IT.Tag, 1)
    SQL = SQL & RC & " AND linliapu ="
    RC = RecuperaValor(IT.Tag, 2)
    SQL = SQL & RC & ";"
    If ModoPunteo = 0 Then
        Conn.Execute SQL
        InsertarCtaCadenaPunteados
    End If
    
    'Ponemos los importes
    If De <> 0 Then
        Text2(0).Text = Format(De, FormatoImporte)
        Else
        Text2(0).Text = ""
    End If
    If Ha <> 0 Then
        Text2(1).Text = Format(Ha, FormatoImporte)
        Else
        Text2(1).Text = ""
    End If
    Sng = De - Ha
    If Sng <> 0 Then
        Text2(2).Text = Format(Sng, FormatoImporte)
        Else
        Text2(2).Text = ""
    End If
    
    
    Exit Sub
EPuntea:
    MuestraError Err.Number, "Accediendo BD para puntear", Err.Description
End Sub



Private Function ObtenerCuenta(Siguiente As Boolean) As Boolean

    SQL = "select codmacta from hlinapu"
    If EjerciciosCerrados Then SQL = SQL & "1"
    SQL = SQL & " WHERE codmacta "
    If Siguiente Then
        SQL = SQL & ">"
    Else
        SQL = SQL & "<"
    End If
    SQL = SQL & " '" & Text4.Text & "'"
    SQL = SQL & " AND  fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    
    If chkSin.Value = 1 Then SQL = SQL & " AND punteada =0 "
    
    SQL = SQL & " group by codmacta ORDER BY codmacta"
    If Siguiente Then
        SQL = SQL & " ASC"
    Else
        SQL = SQL & " DESC"
    End If
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        SQL = "No se ha obtenido la cuenta "
        If Siguiente Then
            SQL = SQL & " siguiente."
        Else
            SQL = SQL & " anterior."
        End If
        MsgBox SQL, vbExclamation
        ObtenerCuenta = False
    Else
        Text4.Text = RS!codmacta
        Text5.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RS!codmacta, "T")
        ObtenerCuenta = True
    End If
    RS.Close
    Set RS = Nothing
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Image2_Click
Case 4
    OtraCuenta False
Case 5
    OtraCuenta True
Case 7
    ' en historicoCalculamos saldo. Lleva ya el sql montado
    SaldoHistorico Text4.Text, Text3(0).Tag, Text5.Text, EjerciciosCerrados
Case 8
    Screen.MousePointer = vbHourglass
    HazSumas
    Screen.MousePointer = vbDefault
Case 9

    Screen.MousePointer = vbHourglass
    DesmarcaTodo
    Screen.MousePointer = vbDefault


Case 11
    'Comprobamos si hay algun ITEM seleccionado
    If AlgunNodoPunteado Then
        MsgBox "Existen lineas punteadas.  Debe seleccionar solo 'Sin puntear'", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    'Punteo automatico por importes
    PonerModoPunteo True
    'Refrescamos el form
    Me.Refresh
    PuntearImportesAutomaticamente
    Screen.MousePointer = vbDefault
    DoEvents
    
Case 13
    Unload Me
End Select
End Sub


Private Sub HazSumas()
Dim Im As Currency
Dim PuntD As Currency
Dim PuntH As Currency
Dim d As Currency
Dim H As Currency
    On Error GoTo EHazSumas
    d = 0
    H = 0
    PuntD = 0: PuntH = 0
    'Recorremos el debe
    With lwD
        If .ListItems.Count > 0 Then
            For I = 1 To .ListItems.Count
                Im = CCur(ImporteFormateado(.ListItems(I).SubItems(3)))
                If Not .ListItems(I).Checked Then
                 
                    d = d + Im
                    
                Else
                    PuntD = PuntD + Im
                End If
            Next I
        End If
    End With
    
    
    With lwh
        If .ListItems.Count > 0 Then
            For I = 1 To .ListItems.Count
                Im = CCur(ImporteFormateado(.ListItems(I).SubItems(3)))
                If Not .ListItems(I).Checked Then
                    H = H + Im
                Else
                    PuntH = PuntH + Im
                End If
            Next I
        End If
    End With
    
    
    SQL = Format(PuntD, FormatoImporte) & "|" & Format(d, FormatoImporte) & "|" & Format(PuntD + d, FormatoImporte) & "|"
    SQL = SQL & Format(PuntH, FormatoImporte) & "|" & Format(H, FormatoImporte) & "|" & Format(PuntH + H, FormatoImporte) & "|"
    'Las diferencias
    SQL = SQL & Format(PuntD - PuntH, FormatoImporte) & "|" & Format(d - H, FormatoImporte) & "|" & Format((PuntD - PuntH) + (d - H), FormatoImporte) & "|"
    
    frmMensajes.Parametros = SQL
    frmMensajes.Opcion = 18
    frmMensajes.Show vbModal
    
    Exit Sub
EHazSumas:
    
    MuestraError Err.Number, "Realizando sumas Debe/haber"
End Sub


Private Sub DesmarcaTodo()

    SQL = "Va a desmarcar todos los punteos para: " & vbCrLf & vbCrLf
    SQL = SQL & "Cuenta: " & Text4.Text & " - " & Text5.Text & vbCrLf
    SQL = SQL & "Fecha inicio: " & Text3(0).Text & vbCrLf
    SQL = SQL & "Fecha fin:     " & Text3(1).Text & vbCrLf & vbCrLf & vbCrLf
    SQL = SQL & "          ¿Desea continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    SQL = "UPDATE hlinapu"
    If EjerciciosCerrados Then SQL = SQL & "1"
    SQL = SQL & " SET punteada=0 WHERE codmacta = '" & Text1.Text & "'"
    SQL = SQL & " AND fechaent>= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent<= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    Conn.Execute SQL
    InsertarCtaCadenaPunteados
    CargarDatos
End Sub


Private Sub PonerModoPunteo(ModoImporte As Boolean)
    ModoPunteo = 0
    If ModoImporte Then
        ModoPunteo = 1
        lwD.Height = 6435
        lwh.Height = 6435
    Else
        ModoPunteo = 0
        lwD.Height = 6555
        lwh.Height = 6555
    End If
        
    FramePorImportes.Visible = ModoImporte
    With Toolbar1
    
        .Buttons(1).Enabled = Not ModoImporte
        .Buttons(4).Enabled = Not ModoImporte
        .Buttons(5).Enabled = Not ModoImporte
        .Buttons(7).Enabled = Not ModoImporte
        .Buttons(8).Enabled = Not ModoImporte
        .Buttons(9).Enabled = Not ModoImporte
        .Buttons(11).Enabled = Not ModoImporte
 
    End With
    
    
    For I = 0 To 3
        imgCheck(I).Visible = Not ModoImporte
    Next I
End Sub



Private Function AlgunNodoPunteado() As Boolean


    AlgunNodoPunteado = True
    
    For I = 1 To lwD.ListItems.Count
        If lwD.ListItems(I).Checked Then Exit Function
    Next I
    For I = 1 To lwh.ListItems.Count
        If lwh.ListItems(I).Checked Then Exit Function
    Next I
    'Si llega aqui es que NO hay ninguno punteado
    AlgunNodoPunteado = False
End Function



Private Sub PuntearImportesAutomaticamente()
Dim J As Integer
Dim T1 As Single


    T1 = Timer - 1
    For I = 1 To lwD.ListItems.Count
        'Label
        SQL = lwD.ListItems(I).SubItems(3) 'Cargo el importe
        
        If Timer - T1 > 1 Then
            Me.Label7(1).Visible = Not Me.Label7(1).Visible
            If Me.Label7(1).Visible Then Me.Label7(1).Refresh
            T1 = Timer
        End If
        
        For J = 1 To lwh.ListItems.Count
            If Not lwh.ListItems(J).Checked Then
                RC = lwh.ListItems(J).SubItems(3)
                If SQL = RC Then
                    'EUREKA!!!!!! El mismo importe
                    lwD.ListItems(I).Checked = True
                    PunteaEnBD lwD.ListItems(I), True
                    lwh.ListItems(J).Checked = True
                    PunteaEnBD lwh.ListItems(J), False
                    'Nos salimos del for
                    Exit For
                End If
            End If
        Next J
        
        
    Next I
    Me.Label7(1).Visible = False
End Sub



'-------------------------------------------------------
'-------------------------------------------------------
'Para el LOG de punteo de cuentas
'-------------------------------------------------------
'-------------------------------------------------------
Private Sub InsertarCtaCadenaPunteados()
Dim Aux As String

    Aux = Me.Text4.Text & "|"
    If InStr(1, CtasQueHaPunteado, Aux) = 0 Then CtasQueHaPunteado = CtasQueHaPunteado & Aux
        
End Sub


Private Sub VerLogPunteado()

    On Error GoTo Evl
    If CtasQueHaPunteado <> "" Then
        CtasQueHaPunteado = Replace(CtasQueHaPunteado, "|", " ")
        CtasQueHaPunteado = "Cuentas punteadas: " & CtasQueHaPunteado
        vLog.Insertar 17, vUsu, CtasQueHaPunteado
    End If
    
    Exit Sub
Evl:
    MuestraError Err.Number, Err.Description
End Sub
