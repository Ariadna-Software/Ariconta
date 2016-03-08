VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmColBalan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balances"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmColBalan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Copiar"
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   12
      Top             =   3960
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Index           =   4
      Left            =   6180
      TabIndex        =   11
      Top             =   2760
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Index           =   6
      Left            =   6180
      TabIndex        =   10
      Top             =   4680
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Comprobar"
      Height          =   375
      Index           =   3
      Left            =   6180
      TabIndex        =   9
      Top             =   2145
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Eliminar"
      Height          =   375
      Index           =   2
      Left            =   6180
      TabIndex        =   8
      Top             =   1590
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modificar"
      Height          =   375
      Index           =   1
      Left            =   6180
      TabIndex        =   7
      Top             =   1035
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nuevo"
      Height          =   375
      Index           =   0
      Left            =   6180
      TabIndex        =   6
      Top             =   480
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo balance"
      Height          =   675
      Left            =   180
      TabIndex        =   1
      Top             =   4440
      Width           =   5775
      Begin VB.OptionButton Option1 
         Caption         =   "Situación"
         Height          =   255
         Index           =   2
         Left            =   4500
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pérdidas y ganancias"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Top             =   300
         Width           =   1875
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6060
      Top             =   4080
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColBalan.frx":030A
      Height          =   3915
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6906
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   60
      Width           =   7335
   End
End
Attribute VB_Name = "frmColBalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean



Private Sub Command1_Click(Index As Integer)
Dim CodigoBalanceBuscar As Integer
    If Index >= 1 And Index <= 5 Then
        If Adodc1.Recordset.EOF Then Exit Sub
    End If
    If Index > 2 And Index < 6 Then
        CodigoBalanceBuscar = Adodc1.Recordset!NumBalan
    End If
    
    Select Case Index
    Case 0, 1
        Screen.MousePointer = vbHourglass
        'Nuevo balance   y modificar
        If Index = 0 Then
            NumRegElim = ObtenerSiguiente
            frmBalances.numBalance = 0
        Else
            NumRegElim = 0
            frmBalances.numBalance = Adodc1.Recordset!NumBalan
        End If
        frmBalances.Show vbModal
        'Para luego hace la busqueda
        If Index = 1 Then NumRegElim = Adodc1.Recordset!NumBalan
        CodigoBalanceBuscar = NumRegElim
    Case 2
    
        EliminarBalance
        
    Case 3
    
    
        Screen.MousePointer = vbHourglass
        Label1.Tag = Label1.Caption
        Label1.Caption = "Comprobaciones ....."
        Label1.Refresh
        ComprobarBalance Adodc1.Recordset!NumBalan, Adodc1.Recordset!perd = "SI"
        Label1.Caption = Label1.Tag
        Label1.Tag = ""
        Screen.MousePointer = vbDefault
        
    
    Case 4
        
        'Vamos a utilizar la temporal de balances donde dejara los valores
        GeneraDatosBalanConfigImpresion Adodc1.Recordset!NumBalan
        
        
        CadenaDesdeOtroForm = "Titulo= """ & DevNombreSQL(Adodc1.Recordset!nombalan) & """|"
        

        With frmImprimir
            .OtrosParametros = CadenaDesdeOtroForm
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .Opcion = 73
            .Show vbModal
        End With
        
        
    Case 5
        CadenaDesdeOtroForm = Adodc1.Recordset!NumBalan & "|" & Adodc1.Recordset!nombalan & "|"
        frmListado.Opcion = 57
        frmListado.Show vbModal
        CargaGrid
        
    Case Else
        Unload Me
        
        
    
    End Select
    If Index <> 6 Then
        Screen.MousePointer = vbHourglass
        CargaGrid
        If Index <> 2 Then
            Adodc1.Recordset.Find "Numbalan = " & CodigoBalanceBuscar
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaGrid
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Label1.Caption = "Relación de balances configurados"
    If vParam.NuevoPlanContable Then Label1.Caption = Label1.Caption & "  P.G.C.  2008"
End Sub



Private Sub CargaGrid()
Dim SQL As String
    
    SQL = ""
    If Option1(1).Value Then
        SQL = " WHERE Perdidas = 1"
    Else
        If Option1(2).Value Then
            SQL = " WHERE Perdidas = 0"
        End If
    End If
    SQL = SQL & " ORDER BY Numbalan"
    SQL = "select numbalan,nombalan, if(perdidas=1,'SI','NO') as Perd ,if(predeterminado=1,'*','') as Pre  from sbalan " & SQL
    Adodc1.RecordSource = SQL
    Adodc1.ConnectionString = Conn
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    DataGrid1.Columns(0).Caption = "Número"
    DataGrid1.Columns(0).Width = 700
    
    DataGrid1.Columns(1).Caption = "Nombre"
    DataGrid1.Columns(1).Width = 3000

    DataGrid1.Columns(2).Caption = "P y G"
    DataGrid1.Columns(2).Width = 700
    
    DataGrid1.Columns(3).Caption = "Pred."
    DataGrid1.Columns(3).Width = 700
    
End Sub
    

Private Function ObtenerSiguiente() As Long

    ObtenerSiguiente = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select max(numbalan) from sbalan", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        ObtenerSiguiente = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    ObtenerSiguiente = ObtenerSiguiente + 1
End Function

Private Sub Option1_Click(Index As Integer)
    If PrimeraVez Then Exit Sub
    Screen.MousePointer = vbHourglass
    CargaGrid
    Screen.MousePointer = vbDefault
End Sub


Private Sub EliminarBalance()
Dim SQL As String
    SQL = "Seguro que desea eliminar el balance: " & Adodc1.Recordset!nombalan & "?"
    If MsgBox(SQL, vbExclamation + vbYesNo) <> vbYes Then Exit Sub
    
    'Eliminamos las cuentas
    SQL = "DELETE FROM sperdi2 WHere numbalan=" & Adodc1.Recordset!NumBalan
    Conn.Execute SQL
    
    'Eliminamos las lineas del balance
    SQL = "DELETE FROM sperdid WHere numbalan=" & Adodc1.Recordset!NumBalan
    Conn.Execute SQL
    
    'Eliminamos el balance
    SQL = "DELETE FROM sbalan WHere numbalan=" & Adodc1.Recordset!NumBalan
    Conn.Execute SQL
    
End Sub



Private Sub ComprobarBalance(NumBal, EsPerdidas As Boolean)
Dim Cad As String
    
    
    'UPDATEAMOS TIENE CUENTAS A 0
    Conn.Execute "UPDATE sperdid SET tienenctas=0 where numbalan=" & NumBal
    
    Set miRsAux = New ADODB.Recordset
    Cad = "select numbalan,pasivo,codigo from sperdi2 group by"
    Cad = Cad & " numbalan,pasivo,codigo having numbalan=" & NumBal
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = "UPDATE sperdid set tienenctas=1 where numbalan=" & miRsAux!NumBalan
        Cad = Cad & " and pasivo='" & miRsAux!Pasivo & "' AND codigo = " & miRsAux!Codigo
        Conn.Execute Cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    'Haremos una segunda comprobacion
    Label1.Caption = "Obteniendo cuentas ejercicio actual"
    Me.Refresh
    DoEvents
            
            Conn.Execute "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
            Cad = "INSERT INTO tmpcierre1(codusu,cta) "
            Cad = Cad & "Select " & vUsu.Codigo & ",codmacta from hlinapu where "
            Cad = Cad & " fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            Cad = Cad & "' AND fechaent<='" & Format(vParam.fechafin, FormatoFecha) & "' AND "

            If Not EsPerdidas Then Cad = Cad & " NOT "
            Cad = Cad & "(codmacta like '" & vParam.grupogto & "%' OR "
            Cad = Cad & "codmacta like '" & vParam.grupovta & "%'"
            If vParam.Subgrupo1 <> "" Then Cad = Cad & " OR " & "codmacta like '" & vParam.Subgrupo1 & "%'"
            Cad = Cad & ") GROUP BY codmacta"
            Conn.Execute Cad
            
            
            'Ya tengo todas las cuentas que entran en hlinapu
            'Cogere la configuracion y para cada cuenta ire quitando
            'Las que esten configuradas
            Label1.Caption = "Comprobando configuracion"
            Me.Refresh
            DoEvents
    
            
            Cad = "Select * from sperdi2 where numbalan=" & NumBal
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo & " AND cta like '"
            While Not miRsAux.EOF
                Conn.Execute Cad & miRsAux!codmacta & "%'"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
                
            'Veremos las que queden
            Cad = "SELECT * FROM tmpcierre1 where codusu =" & vUsu.Codigo
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            Cad = ""
            While Not miRsAux.EOF
                Cad = Cad & "      " & miRsAux!Cta
                NumRegElim = NumRegElim + 1
                If (NumRegElim Mod 11) = 0 Then Cad = Cad & vbCrLf
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If NumRegElim > 0 Then
                Cad = "Hay cuentas(" & NumRegElim & ") que parecen no haber sido configuradas en el balance" & vbCrLf & vbCrLf & Cad
                MsgBox Cad, vbInformation
                NumRegElim = 0
            Else
                
                'Haremos una tercera comprobacion
                Label1.Caption = "Nodos superiores con cuentas"
                Me.Refresh
                DoEvents
            
                Cad = "select * from sperdid where numbalan=1 and tienenctas=1 and (pasivo,codigo) in ("
                Cad = Cad & " select pasivo,padre from sperdid where numbalan=1 and padre>=0) order by pasivo,codigo"
                miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                NumRegElim = 0
                Cad = ""
                While Not miRsAux.EOF
                    Cad = Cad & "      -" & miRsAux!Pasivo & miRsAux!Codigo & "     " & miRsAux!deslinea & vbCrLf
                    NumRegElim = NumRegElim + 1
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                
                If NumRegElim > 0 Then
                    Cad = "No es nodo de ultimo nivel y tienen cuentas configuradas" & vbCrLf & Cad
                    MsgBox Cad, vbExclamation
                Else
                    MsgBox "Comprobacion  finalizada", vbInformation
                End If
            End If
    Set miRsAux = Nothing
    

End Sub
