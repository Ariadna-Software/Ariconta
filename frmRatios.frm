VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRatios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ratios y gráficas"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   Icon            =   "frmRatios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   6720
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ratios"
      TabPicture(0)   =   "frmRatios.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6(14)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkRatio(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkRatio(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkRatio(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDescrip(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDescrip(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescrip(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Graficas"
      TabPicture(1)   =   "frmRatios.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4(33)"
      Tab(1).Control(1)=   "Label4(0)"
      Tab(1).Control(2)=   "Label4(1)"
      Tab(1).Control(3)=   "chkGraf1(0)"
      Tab(1).Control(4)=   "List1"
      Tab(1).Control(5)=   "cboMes"
      Tab(1).ControlCount=   6
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   17
         Top             =   5880
         Width           =   1095
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   -69360
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
      Begin VB.ListBox List1 
         Height          =   2985
         Left            =   -74040
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkGraf1 
         Caption         =   "Resumen"
         Height          =   255
         Index           =   0
         Left            =   -69360
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtDescrip 
         Height          =   1695
         Index           =   2
         Left            =   4680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmRatios.frx":0044
         Top             =   4440
         Width           =   4455
      End
      Begin VB.TextBox txtDescrip 
         Height          =   1695
         Index           =   1
         Left            =   4680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmRatios.frx":004A
         Top             =   2520
         Width           =   4455
      End
      Begin VB.TextBox txtDescrip 
         Height          =   1695
         Index           =   0
         Left            =   4680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmRatios.frx":0050
         Top             =   600
         Width           =   4455
      End
      Begin VB.CheckBox chkRatio 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   600
         TabIndex        =   4
         Top             =   4440
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkRatio 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkRatio 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmRatios.frx":0056
         Top             =   5895
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   16
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comparativa mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   -71280
         TabIndex        =   15
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   -71280
         TabIndex        =   13
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Años"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   33
         Left            =   -74760
         TabIndex        =   12
         Top             =   840
         Width           =   420
      End
   End
   Begin VB.Label lblInd 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   6720
      Width           =   2775
   End
End
Attribute VB_Name = "frmRatios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim SQL As String
Dim I As Integer

Private Sub chkRatio_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim B As Boolean
Dim Aux As String

    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    If Me.SSTab1.Tab = 0 Then
        'Ratios
        B = HacerRatios
    Else
        'Graficos
        B = HacerGraficas
    End If
    
    lblInd.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    
    If B Then
            '96 .- Ratios
            '97 .- Graficas
            '98 .- Graficas resumen
    
    
            'OK vamos a lanzar los reports
            With frmImprimir
                If Me.SSTab1.Tab = 0 Then
                    .Opcion = 96
                    
                    SQL = " hasta " & Text3(0).Text
                    
                    .OtrosParametros = "Desde= "" " & SQL & """|"
                    .FormulaSeleccion = "{ztesoreriacomun.codusu}=" & vUsu.Codigo
                Else
                    If Me.chkGraf1(0).Value = 0 Then
                        'Grafica
                        .Opcion = 97
                        SQL = ""
                        Aux = ""
                        For NumRegElim = List1.ListCount - 1 To 0 Step -1
                            If List1.Selected(NumRegElim) Then
                                SQL = SQL & "1"
                                If Aux = "" Then
                                    'Preimer ejercicio
                                    Aux = "TextoEjer1= """ & List1.List(NumRegElim) & """|"
                                    
                                Else
                                    'Segundo
                                    Aux = Aux & "TextoEjer2= """ & List1.List(NumRegElim) & """|"
                                End If
                            End If
                        Next
                        I = 0
                        If Len(SQL) > 1 Then I = 1
                        
                        SQL = "Comparativo= " & I & "|" & Aux
                    
                        
                        .FormulaSeleccion = "{ztmpbalancesumas.codusu}=" & vUsu.Codigo
                        .OtrosParametros = SQL
                        .NumeroParametros = 1
                    Else
                        'Resumen
                        .Opcion = 98
                        .FormulaSeleccion = "{zsaldoscc.codusu}=" & vUsu.Codigo
                        
                        SQL = ""
                        Aux = ""
                        For NumRegElim = 0 To List1.ListCount - 1
                            If List1.Selected(NumRegElim) Then
                                If Aux = "" Then Aux = "UltAno= " & Mid(List1.List(NumRegElim), 1, 4) & "|"
                            End If
                        Next
                        SQL = Aux
                        NumRegElim = 1
                        If cboMes.ListIndex > 0 Then
                            'ha seleccionado mes
                            SQL = SQL & "Desde= ""Hasta " & cboMes.Text & """|"
                            NumRegElim = 2
                        End If
                        
                        .OtrosParametros = SQL
                        .NumeroParametros = NumRegElim
                    End If
                End If

                
                .SoloImprimir = False
                .Show vbModal
            End With
        
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Me.Icon = frmPpal.Icon
   SQL = "01/" & Month(Now) & "/" & Year(Now)
   SQL = DateAdd("d", -1, CDate(SQL))
   Text3(0).Text = SQL
   CargaDatosRatios
   CargaDatosGraficas
End Sub



Private Sub CargaDatosRatios()

    'NO puede dar error

    'En sbalan, del 51 al 53 tiene que existir  CUANDO ESTEN TODOS sera hasta el 55
    SQL = "Select * from sbalan where numbalan>=51 and numbalan<=54 order by numbalan"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        If I < 2 Then
    
            I = miRsAux!NumBalan - 51
            
            Me.chkRatio(I).Caption = miRsAux!nombalan
            Me.txtDescrip(I).Text = miRsAux!Descripcion
            
        End If
        
        miRsAux.MoveNext
        
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    Me.cboMes.Clear
    Me.cboMes.AddItem " " 'todos
    For I = 1 To 12
        Me.cboMes.AddItem Format("23/" & I & "/2000", "mmmm")
    Next
    
End Sub


Private Function HacerRatios() As Boolean
    HacerRatios = False

    If Text3(0).Text = "" Then
        MsgBox "Ponga la fecha", vbExclamation
        Exit Function
    End If
    
    If FechaCorrecta2(CDate(Text3(0).Text)) > 2 Then
        MsgBox "Fuera de ejercicios", vbExclamation
        Exit Function
    End If
    
    
    
    NumRegElim = DiasMes(Month(Text3(0).Text), Year(Text3(0).Text))
    If Day(Text3(0).Text) <> NumRegElim Then
        MsgBox "Saldos mensuales", vbExclamation
        SQL = NumRegElim & "/" & Format(Month(Text3(0).Text), "00") & "/" & Year(Text3(0).Text)
        Text3(0).Text = SQL
    End If


    Conn.Execute "DELETE FROM tmpimpbalance where codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM usuarios.ztmpimpbalan where codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    
    If Me.chkRatio(0).Value = 1 Then CargarDatosRatio 51
    If Me.chkRatio(1).Value = 1 Then CargarDatosRatio 52
    If Me.chkRatio(2).Value = 1 Then CargarDatosRatio 53
    'If Me.chkRatio(0).Value = 1 Then CargarDatosRatio 54
    
    
    
    SQL = "Select count(*) from tmpimpbalance where codusu=" & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then NumRegElim = miRsAux.Fields(0)
    End If
    miRsAux.Close
    If NumRegElim = 0 Then
        MsgBox "No existen datos"
        Exit Function
    End If
    
    'Insertaremos en la usuarios.z
    SQL = "insert into usuarios.ztmpimpbalan (`codusu`,`Pasivo`,`codigo`,`descripcion`,`linea`,`importe1`)"
    SQL = SQL & " select codusu,pasivo,codigo,descripcion,linea,importe1 from tmpimpbalance where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    
    HacerRatios = True
    
    
    
    
End Function


Private Sub CargarDatosRatio(Cual As Integer)
Dim Lin As Collection
Dim Col As Collection
Dim J As Integer
Dim Importe As Currency
Dim ImpLin As Currency
Dim EsPasivo As Boolean
    
    
    Set Lin = New Collection

    SQL = "Select * from sperdid where numbalan=" & Cual
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Lin.Add CStr(miRsAux!Codigo)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Lin.Count = 0 Then
        Set Lin = Nothing
        Exit Sub
    End If
    
    For I = 1 To Lin.Count
    
        Me.lblInd.Caption = "Lineas " & Lin.Item(I)
        Me.lblInd.Refresh
    
        Set Col = New Collection
        SQL = "Select * from sperdi2 where numbalan=" & Cual & " AND codigo=" & Lin.Item(I)
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            SQL = SQL & ", '" & miRsAux!codmacta & "'"
            If Len(SQL) > 100 Then
                SQL = Mid(SQL, 2)
                Col.Add SQL
                SQL = ""
            End If
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If SQL <> "" Then
            SQL = Mid(SQL, 2)
            Col.Add SQL
        End If
        
        
        '-------------------------------------------------------
        '
        '
        
        
        Importe = 0
        
        'Cuentas de pasivo. Van con Haber-Debe
        EsPasivo = False
        Select Case Cual
        Case 51
            'Ratio tesoreria
            If Lin.Item(I) = 3 Then EsPasivo = True
        Case 52
            'Liquidez
            If Lin.Item(I) = 2 Then EsPasivo = True
        Case 53
            If Lin.Item(I) >= 3 Then EsPasivo = True  '3 y 4
        
        End Select
        
       
        
        For J = 1 To Col.Count
            Me.lblInd.Caption = "Saldos " & Lin.Item(I) & ": " & J & " de " & Col.Count
            Me.lblInd.Refresh
                
            SQL = "SELECT sum(impmesde-impmesha) FROM hsaldos WHERE "
            If Year(vParam.fechaini) = Year(vParam.fechafin) Then
                'año natural
                If Year(vParam.fechaini) = Year(Text3(0).Text) Then
                    'Año ejercicio actual
                    SQL = SQL & " anopsald = " & Year(Text3(0).Text)
                    SQL = SQL & " and mespsald <= " & Month(Text3(0).Text)
                Else
                    'Año siguiente
                    SQL = SQL & "(( anopsald = " & Year(vParam.fechaini)
                    SQL = SQL & ") OR (anopsald = " & Year(Text3(0).Text)
                    SQL = SQL & " and mespsald <= " & Month(Text3(0).Text) & "))"
                
                End If
                
            Else
                
                If Year(vParam.fechaini) = Year(Text3(0).Text) Then
                    'Este trozo de año actual
                    SQL = SQL & " (anopsald=" & Year(vParam.fechaini) & " and mespsald between " & Month(vParam.fechaini) & " AND  " & Month(Text3(0).Text) & ")"
                Else
                    If Year(vParam.fechafin) = Year(Text3(0).Text) Then
                        'Lo que queda de este año
                        SQL = SQL & " ((anopsald=" & Year(vParam.fechaini) & " and mespsald >= " & Month(vParam.fechaini) & " ) OR "
                        SQL = SQL & " (anopsald=" & Year(Text3(0).Text) & " and mespsald <= " & Month(Text3(0).Text) & " ))"
                    Else
                        'Hasta siguiente
                            SQL = SQL & " ((anopsald=" & Year(vParam.fechaini) & " and mespsald >= " & Month(vParam.fechaini) & " ) OR "
                            SQL = SQL & " anopsald = " & Year(vParam.fechafin)
                            SQL = SQL & " OR (anopsald=" & Year(Text3(0).Text) & " and mespsald <= " & Month(Text3(0).Text) & " ))"
                    End If
                
                End If
            End If
            SQL = SQL & " AND codmacta IN (" & Col.Item(J) & ")"
            
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux.Fields(0)) Then
                    If EsPasivo Then
                        ImpLin = -miRsAux.Fields(0)
                    Else
                        ImpLin = miRsAux.Fields(0)
                    End If
                    
                    Importe = Importe + ImpLin
                End If
            End If
            miRsAux.Close
            
        Next J
        Set Col = Nothing
        
        NumRegElim = Cual * 100
        NumRegElim = NumRegElim + Val(Lin.Item(I))
        
        SQL = "insert into `tmpimpbalance` (`codusu`,`Pasivo`,`codigo`,`importe1`,`descripcion`,`linea`,"
        SQL = SQL & "`importe2`,`negrita`,`orden`,`QueCuentas`) values ( " & vUsu.Codigo & ",'" & Chr(Cual + 14) & "',"
        SQL = SQL & NumRegElim & "," & TransformaComasPuntos(CStr(Importe))
        SQL = SQL & ",'',NULL,NULL,NULL,'0',NULL)"
        Conn.Execute SQL
        
        'Lo que seran los textos
        
    Next I
        
    SQL = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,observa1,`Texto`)"
    SQL = SQL & " select " & vUsu.Codigo & ",sbalan.numbalan*100+codigo,nombalan,deslinea,descripcion from sbalan,sperdid where sbalan.numbalan=sperdid.numbalan and sperdid.numbalan=" & Cual & " order by orden"
    Conn.Execute SQL
        
        
End Sub



Private Sub CargaDatosGraficas()

    SQL = "select anopsald from hsaldos group by 1 order by 1 desc"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    SQL = ""
    While Not miRsAux.EOF
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            'Año natural
            SQL = miRsAux!anopsald
        
        Else
            'Sera yyyy - yyyy  . Posiciones fijas.  4 prim año 1  desde la 8 año 2
            If SQL = "" Then
                
                    If miRsAux!anopsald > Year(vParam.fechaini) Then
                        List1.AddItem Format(miRsAux!anopsald, "0000") & " - " & Format(miRsAux!anopsald + 1, "0000")
                    End If

            End If
        
            SQL = Format(miRsAux!anopsald - 1, "0000") & " - " & Format(miRsAux!anopsald, "0000")
                    
        End If
        List1.AddItem SQL
        I = I + 1
        miRsAux.MoveNext
    Wend
    If I > 0 Then List1.Selected(0) = True
    miRsAux.Close
        
    
End Sub



Private Function HacerGraficas() As Boolean
Dim Veces As Byte  'para años partidos SON dos
Dim Ingresos As Currency
Dim Gastos As Currency
Dim Aux As Currency
Dim AnyoMes As Long
Dim Comparativo As Boolean

    
    HacerGraficas = False
    
    SQL = ""
    For I = 0 To Me.List1.ListCount - 1
        If List1.Selected(I) Then SQL = SQL & "1"
    Next
    If Len(SQL) < 1 Then
        MsgBox "Seleccione un año", vbExclamation
        Exit Function
    End If
    Comparativo = False
    If Len(SQL) = 2 Then
        Comparativo = True
        If cboMes.ListIndex <= 0 Then
            MsgBox "Seleccione el mes para el comparativo", vbExclamation
            Exit Function
        End If
    
    End If
    If Me.chkGraf1(0).Value = 0 And Len(SQL) > 2 Then
        MsgBox "Seleccione un año(dos para el comparativo)", vbExclamation
        Exit Function
    End If
    
    
'    If Me.chkGraf1(0).Value = 1 Then
'        If Year(vParam.fechafin) <> Year(vParam.fechaini) Then
'            MsgBox "Error grafica resumen. Consulte soporte técnico", vbExclamation
'            Exit Function
'        End If
'    End If
    
    Me.lblInd.Caption = "Prepara datos"
    Me.lblInd.Refresh
    
    
    SQL = "DELETE FROM tmpgraficas where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    Conn.Execute "DELETE FROM usuarios.ztmpbalancesumas where codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM usuarios.zsaldoscc where codusu = " & vUsu.Codigo
    
    'la de los informes tb
    For I = 0 To List1.ListCount - 1
        Veces = 1
        If Year(vParam.fechafin) <> Year(vParam.fechaini) Then Veces = 2
        If List1.Selected(I) Then
            Me.lblInd.Caption = List1.List(I)
            Me.lblInd.Refresh
            'Este esta selecionado
            While Veces <> 0
                SQL = "select anopsald,mespsald,codmacta,impmesde,impmesha"
                SQL = SQL & "  from hsaldos where"
                SQL = SQL & " (codmacta='6' or codmacta ='7') "
            
                If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                    'AÑO NATURAL
                    SQL = SQL & " AND anopsald= " & List1.List(I)
                    
                                    'Quiere hasta un mes
                    If Me.cboMes.ListIndex > 0 Then SQL = SQL & " AND mespsald<= " & cboMes.ListIndex
                    
                    
                Else
                    'Años aprtidos
                    'Si veces=1 entonces el primer trozo de año partido
                    If Veces = 2 Then
                        'Segundo trozo
                        SQL = SQL & " AND anopsald= " & Mid(List1.List(I), 8)
                        SQL = SQL & " AND mespsald <=  " & Month(vParam.fechafin)
                        'Quiere hasta un mes
                        If Me.cboMes.ListIndex > 0 Then
                            If cboMes.ListIndex < Month(vParam.fechaini) Then SQL = SQL & " AND mespsald<= " & cboMes.ListIndex
                        End If
                        
                    Else
                        SQL = SQL & " AND anopsald= " & Mid(List1.List(I), 1, 4)
                        SQL = SQL & " AND mespsald >=  " & Month(vParam.fechaini)
                        If Me.cboMes.ListIndex > 0 Then
                            If cboMes.ListIndex >= Month(vParam.fechaini) Then SQL = SQL & " AND mespsald<= " & cboMes.ListIndex
                        End If
                        
                    End If
                End If
                SQL = SQL & " ORDER BY 1,2,3"
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                AnyoMes = 0
                While Not miRsAux.EOF
                    
                    NumRegElim = miRsAux!anopsald * 100 + miRsAux!mespsald
                    If NumRegElim <> AnyoMes Then
                        'Nuevo ano,mes
                        If AnyoMes > 0 Then
                            'Ya tienen valor
                            InsertaEnTmpGraf AnyoMes, Ingresos, Gastos
                            
                        End If
                       
                        Ingresos = 0: Gastos = 0
                        AnyoMes = NumRegElim
                    End If
                
                    Aux = miRsAux!impmesde - miRsAux!impmesha
                    If miRsAux!codmacta = "6" Then
                        Gastos = Gastos + Aux
                    Else
                        Ingresos = Ingresos - Aux 'va saldo
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                'El ultimo
                If AnyoMes > 0 Then InsertaEnTmpGraf AnyoMes, Ingresos, Gastos
                
                Veces = Veces - 1
                
             
            Wend
            
        End If
    Next
    
    
    'Si no el el de RESUMEN
    If chkGraf1(0).Value = 0 Then
            'Ya tengo en tmpgrafiacs los valores de los meses
            'Insertare los 12 meses a ceros
            Me.lblInd.Caption = "Carga meses"
            Me.lblInd.Refresh
            SQL = ""
            If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                'SQL = "INSERT INTO ztmpbalancesumas (`codusu`,`cta`,`nomcta`,`aperturaD`,`aperturaH`,`acumAntD`,`acumAntH`,`acumPerD`,"
                'SQL = SQL & "`acumPerH`,`TotalD`,`TotalH`) values "
              
                For Veces = 1 To 12
                    SQL = SQL & ", (" & vUsu.Codigo & ",'" & Format(Veces, "00") & "','" & Format("20/" & Veces & "/2000", "mmmm") & "',0,0,0,0,0,0,0,0)"
                Next Veces
                SQL = Mid(SQL, 2) 'quito la primera cma
                SQL = "INSERT INTO usuarios.ztmpbalancesumas (`codusu`,`cta`,`nomcta`,`aperturaD`,`aperturaH`,`acumAntD`,`acumAntH`,`acumPerD`," & _
                    "`acumPerH`,`TotalD`,`TotalH`) values " & SQL
                Conn.Execute SQL
            
            Else
                SQL = ""
                For Veces = Month(vParam.fechaini) To 12
                    SQL = SQL & ", (" & vUsu.Codigo & ",'00" & Format(Veces, "00") & "','" & Format("20/" & Veces & "/2000", "mmmm") & "',0,0,0,0,0,0,0,0)"
                Next Veces
                For Veces = 1 To Month(vParam.fechafin)
                    SQL = SQL & ", (" & vUsu.Codigo & ",'10" & Format(Veces, "00") & "','" & Format("20/" & Veces & "/2000", "mmmm") & "',0,0,0,0,0,0,0,0)"
                Next Veces
                SQL = Mid(SQL, 2) 'quito la primera cma
                SQL = "INSERT INTO usuarios.ztmpbalancesumas (`codusu`,`cta`,`nomcta`,`aperturaD`,`aperturaH`,`acumAntD`,`acumAntH`,`acumPerD`," & _
                    "`acumPerH`,`TotalD`,`TotalH`) values " & SQL
                Conn.Execute SQL
                    
            End If
            
            
            SQL = "select * from tmpgraficas where codusu = " & vUsu.Codigo & " order by anyo,mes"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 0 'tendre el primer año
            While Not miRsAux.EOF
                Me.lblInd.Caption = miRsAux!Anyo & " " & miRsAux!Mes
                Me.lblInd.Refresh
                If I = 0 Then I = miRsAux!Anyo
                
                SQL = "UPDATE usuarios.ztmpbalancesumas SET "
                
                If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                    'años normales
                    If miRsAux!Anyo = I Then
                        'Año 1
                        'aperturaD aperturaH TotalD
                        SQL = SQL & "aperturaD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        SQL = SQL & ",aperturaH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        SQL = SQL & ",TotalD = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    Else
                        '`acumAntD`,`acumAntH` TotalH
                        SQL = SQL & "acumAntD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        SQL = SQL & ",acumAntH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        SQL = SQL & ",TotalH = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    End If
                    SQL = SQL & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & Format(miRsAux!Mes, "00") & "'"
                    
                Else
                    'años partidos
                    Veces = 0
                    If miRsAux!Anyo <> I Then
                        'Es año siguiente. Pero si el mes es anterior a mesini entonces todavia es ejercicio anterior
                        If miRsAux!Mes < Month(vParam.fechaini) Then
                            Veces = 0
                        Else
                            Veces = 1
                        End If
                    End If
                    
                    If Veces = 0 Then
                        'Año 1
                        'aperturaD aperturaH TotalD
                        SQL = SQL & "aperturaD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        SQL = SQL & ",aperturaH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        SQL = SQL & ",TotalD = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    Else
                        '`acumAntD`,`acumAntH` TotalH
                        SQL = SQL & "acumAntD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        SQL = SQL & ",acumAntH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        SQL = SQL & ",TotalH = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    End If
                    SQL = SQL & " WHERE codusu = " & vUsu.Codigo & " AND cta like '%" & Format(miRsAux!Mes, "00") & "'"
                
                End If
                Conn.Execute SQL
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            
            'Debemos borrar los datos de los meses
            If cboMes.ListIndex > 0 Then
                If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                    SQL = "DELETE FROM usuarios.ztmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta > '" & Format(cboMes.ListIndex, "00") & "'"
                    Conn.Execute SQL
                Else
                    If Month(vParam.fechaini) <= cboMes.ListIndex Then
        
                        
                        SQL = "DELETE FROM usuarios.ztmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta > '00" & Format(cboMes.ListIndex, "00") & "'"
                        Conn.Execute SQL
                    Else
                        'Quiere  hasta parte del años siguiente
                        SQL = "DELETE FROM usuarios.ztmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta > '10" & Format(cboMes.ListIndex, "00") & "'"
                        Conn.Execute SQL
                    
                    End If
                End If
            End If
            
            'Si NO es comparativo ponogo los importes a NULL
            If Not Comparativo Then
                SQL = "update usuarios.ztmpbalancesumas set `acumAntD`=NULL,`acumAntH`=NULL,`acumPerD`=NULL,`acumPerH`=NULL,`TotalH`=NULL"
                SQL = SQL & " where `codusu`=" & vUsu.Codigo
                Conn.Execute SQL
            End If
            
            'Renumeramos mes
            
            SQL = "Select * from usuarios.ztmpbalancesumas WHERE codusu = " & vUsu.Codigo & " ORDER BY cta"
            NumRegElim = 1
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                SQL = "UPDATE usuarios.ztmpbalancesumas SET cta = '" & Format(NumRegElim, "00") & "' WHERE codusu = " & vUsu.Codigo & " AND cta ='" & miRsAux!Cta & "'"
                NumRegElim = NumRegElim + 1
                miRsAux.MoveNext
                Conn.Execute SQL
            Wend
            miRsAux.Close
            
    
    Else
        'RESUMEN
        
        
        SQL = "INSERT INTO  usuarios.zsaldoscc( codusu,codccost,nomccost,ano,mes,impmesde,impmesha)"
        SQL = SQL & " SELECT codusu,'','',anyo,mes,ingresos,gastos from tmpgraficas where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
        
        'Debemos borrar los datos de los meses
        If cboMes.ListIndex > 0 Then
            If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                SQL = "DELETE FROM usuarios.zsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes > " & Format(cboMes.ListIndex, "00")
                Conn.Execute SQL
            Else
                If Month(vParam.fechaini) <= cboMes.ListIndex Then
                    SQL = "DELETE FROM usuarios.zsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes < " & Month(vParam.fechaini)
                    Conn.Execute SQL
                    
                    SQL = "DELETE FROM usuarios.zsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes > " & cboMes.ListIndex
                    Conn.Execute SQL
                Else
                    'Quiere  hasta parte del años siguiente
                    SQL = "DELETE FROM usuarios.zsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes < " & Month(vParam.fechaini) & " AND mes > " & cboMes.ListIndex
                    Conn.Execute SQL
                    
                    
                
                End If
            End If
        End If
        
            
        'El ejercicio va en NOMCOST
       If Year(vParam.fechafin) = Year(vParam.fechaini) Then
            SQL = "UPDATE usuarios.zsaldoscc SET nomccost=ano WHERE codusu=" & vUsu.Codigo
            
        Else
            'SQL = "UPDATE usuarios.zsaldoscc set nomccost=if(mes<9,ano-1,ano)*100+if(mes<9,mes+12,mes)  WHERE codusu=" & vUsu.Codigo
            SQL = "UPDATE usuarios.zsaldoscc set nomccost=if(mes<" & Month(vParam.fechaini) & ",ano-1,ano)  WHERE codusu=" & vUsu.Codigo
        
        End If
        Conn.Execute SQL
    End If
        
    
    HacerGraficas = True
End Function



Private Sub InsertaEnTmpGraf(Id As Long, Ingr As Currency, Gast As Currency)
Dim Aux As Currency
    If Month(vParam.fechafin) = Val(Mid(CStr(Id), 5, 2)) Then
        'MEs del cierre. Hay que quitar PyG
        If CDate("01/" & Mid(CStr(Id), 5, 2) & "/" & Mid(CStr(Id), 1, 4)) < vParam.fechaini Then
            'Hay que quitar Cierre y Pyg
            SQL = "fechaent='" & Mid(CStr(Id), 1, 4) & "-" & Mid(CStr(Id), 5, 2) & "-" & Day(vParam.fechafin) & "'  AND codmacta like '7%' AND codconce"
            SQL = DevuelveDesdeBD("sum(if(isnull(timported),0,timported))-sum(if(isnull(timporteh),0,timporteh))", "hlinapu", SQL, "960")
            If SQL = "" Then SQL = "0"
            Aux = CCur(SQL)
            Ingr = Ingr + Aux
            
            SQL = "fechaent='" & Mid(CStr(Id), 1, 4) & "-" & Mid(CStr(Id), 5, 2) & "-" & Day(vParam.fechafin) & "'  AND codmacta like '6%' AND codconce"
            SQL = DevuelveDesdeBD("sum(if(isnull(timporteh),0,timporteh))-sum(if(isnull(timported),0,timported))", "hlinapu", SQL, "960")
            If SQL = "" Then SQL = "0"
            Aux = CCur(SQL)
            Gast = Gast + Aux
        End If
            
    End If
    SQL = "insert into `tmpgraficas` (`codusu`,`anyo`,`mes`,`ingresos`,`gastos`,`beneficio`) "
    SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Mid(CStr(Id), 1, 4) & "," & Mid(CStr(Id), 5, 2) & ","
    SQL = SQL & TransformaComasPuntos(CStr(Ingr)) & "," & TransformaComasPuntos(CStr(Gast)) & ","
    Ingr = Ingr - Gast
    SQL = SQL & TransformaComasPuntos(CStr(Ingr)) & ")"
    Conn.Execute SQL
End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    SQL = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If SQL <> "" Then
        Text3(Index).Text = SQL
        Text3(Index).SetFocus
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub


Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    Else
        If FechaCorrecta2(CDate(Text3(Index).Text)) > 2 Then
            MsgBox "Fuera de ejercicios", vbExclamation
            Text3(Index).Text = ""
            Text3(Index).SetFocus
        End If
    End If
End Sub
