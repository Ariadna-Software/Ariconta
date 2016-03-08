VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm340 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación IVA intracomunitarias"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   12495
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   10920
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "Asignar"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NIF"
         Object.Width           =   3387
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5211
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nº Fact"
         Object.Width           =   2884
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Importe"
         Object.Width           =   2532
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "% IVA"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Imp. IVA"
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "% I.V.A."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   615
   End
End
Attribute VB_Name = "frm340"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQL As String
    
Private Sub Command2_Click()

End Sub

Private Sub cmdAceptar_Click()

    Unload Me
End Sub

Private Sub cmdAsignar_Click()
Dim primerSel As Integer
Dim ParaClientes As Boolean
Dim Importe As Currency

    If Combo2.ListIndex < 0 Then Exit Sub

    CadenaDesdeOtroForm = ""   'Veremos si ha seleccionado de dos tipos distintos (entragas/adquisiciones)
    SQL = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Selected Then SQL = SQL & "X"
    Next
    
    If SQL = "" Then Exit Sub
    
    
    CadenaDesdeOtroForm = "Va a asignar el iva " & UCase(Combo2.List(Combo2.ListIndex)) & " a " & Len(SQL) & " factura(s).    ¿Continuar?"
    
    If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) <> vbYes Then
        CadenaDesdeOtroForm = ""
    Else
        'Ha dicho que si
        Screen.MousePointer = vbHourglass
        
        primerSel = -1
        For NumRegElim = 1 To ListView1.ListItems.Count
            With ListView1.ListItems(NumRegElim)
                If .Selected Then
                    If primerSel < 0 Then primerSel = .Index
                    SQL = Combo2.List(Combo2.ListIndex)
                    SQL = Trim(Mid(SQL, 1, Len(SQL) - 1))
                    
                    Importe = ImporteFormateado(.SubItems(4)) * (CCur(SQL) / 100)
                    
                    
                    'totiva tipo
                    SQL = "UPDATE tmp340 SET tipo = " & SQL
                    SQL = SQL & ", totiva = " & TransformaComasPuntos(CStr(Importe))
                    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND codigo =" & .Tag
                    Conn.Execute SQL
                    
                    SQL = Combo2.List(Combo2.ListIndex)
                    .SubItems(5) = Trim(Mid(SQL, 1, Len(SQL) - 1)) & ".00"
                    .SubItems(6) = Format(Importe, FormatoImporte)
                    ListView1.Tag = 1
                End If
            End With
        Next
        Combo2.ListIndex = -1
        'CargarDatos
        'If primerSel <= ListView1.ListItems.Count Then
        '    Set ListView1.SelectedItem = ListView1.ListItems(primerSel)
        '    ListView1.SelectedItem.EnsureVisible
        'End If
        CadenaDesdeOtroForm = ""
        Screen.MousePointer = vbDefault
    End If
    
    
End Sub

Private Sub Command1_Click()


    Conn.Execute "DElete FROM tmp340 where codusu = " & vUsu.Codigo
    espera 0.3
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag = 0 Then
        Me.Tag = 1
        
        'Para saber si ha cambiado algo. Si no toca nada, no recalcularemos
        ListView1.Tag = 0
        
        CargarDatos
        

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon
    Me.Tag = 0
    
    Me.ListView1.SmallIcons = frmPpal.ImageList1
    


    Combo2.AddItem "21 %"
    Combo2.AddItem "10 %"
    Combo2.AddItem "4 %"
    Combo2.AddItem "0 %"
    
    
End Sub



Private Sub CargarDatos()
Dim IT As ListItem
On Error GoTo eCargaDatos


    ListView1.ListItems.Clear

    SQL = "select * from tmp340 where codusu=" & vUsu.Codigo
    SQL = SQL & "  and ucase(codpais)<>'ES' and clavelibro='R' "
    SQL = SQL & " ORDER BY fechaexp,nifdeclarado"
    Set miRsAux = New ADODB.Recordset
    
   
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
        'Codigo , nifdeclarado, codpais, nifresidencia, razosoci, idfactura, totalfac fechaexp
        Set IT = ListView1.ListItems.Add()
        IT.Text = miRsAux!codpais & miRsAux!nifdeclarado   'nif
        IT.SubItems(1) = miRsAux!razosoci 'nomnbre
        IT.SubItems(2) = miRsAux!idfactura  'nº fac
        IT.SubItems(3) = miRsAux!fechaexp   'fecha
        IT.SubItems(4) = Format(miRsAux!totalfac, FormatoImporte) 'importe
        
        If miRsAux!Tipo = 0 Then
            IT.SubItems(5) = " "
            SQL = " "
        Else
            IT.SubItems(5) = Format(miRsAux!Tipo, FormatoImporte)
            SQL = Format(miRsAux!totiva, FormatoImporte)
        End If
        IT.SubItems(6) = SQL
        IT.Tag = miRsAux!Codigo
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    CadenaDesdeOtroForm = ""
    
    
    
eCargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number
    
    Set miRsAux = Nothing
End Sub



