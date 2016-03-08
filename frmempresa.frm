VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmempresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresa"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8250
   Icon            =   "frmempresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   8250
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   43
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8070
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmempresa.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(10)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(11)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(13)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(14)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(15)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(17)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(8)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(10)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(12)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(13)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(14)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(15)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(16)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(17)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Otros datos"
      TabPicture(1)   =   "frmempresa.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(39)"
      Tab(1).Control(1)=   "Text1(38)"
      Tab(1).Control(2)=   "Text1(36)"
      Tab(1).Control(3)=   "Text1(20)"
      Tab(1).Control(4)=   "Text1(26)"
      Tab(1).Control(5)=   "Text1(25)"
      Tab(1).Control(6)=   "Text1(24)"
      Tab(1).Control(7)=   "Text1(23)"
      Tab(1).Control(8)=   "Text1(22)"
      Tab(1).Control(9)=   "Text1(21)"
      Tab(1).Control(10)=   "Text1(19)"
      Tab(1).Control(11)=   "Text1(3)"
      Tab(1).Control(12)=   "Text1(4)"
      Tab(1).Control(13)=   "Text1(5)"
      Tab(1).Control(14)=   "Text1(6)"
      Tab(1).Control(15)=   "Text1(18)"
      Tab(1).Control(16)=   "Label1(20)"
      Tab(1).Control(17)=   "Label1(30)"
      Tab(1).Control(18)=   "Label1(28)"
      Tab(1).Control(19)=   "Label1(26)"
      Tab(1).Control(20)=   "Label1(25)"
      Tab(1).Control(21)=   "Label1(24)"
      Tab(1).Control(22)=   "Label1(23)"
      Tab(1).Control(23)=   "Label1(22)"
      Tab(1).Control(24)=   "Label1(21)"
      Tab(1).Control(25)=   "Label1(19)"
      Tab(1).Control(26)=   "Label1(2)"
      Tab(1).Control(27)=   "Label1(3)"
      Tab(1).Control(28)=   "Label1(5)"
      Tab(1).Control(29)=   "Label1(6)"
      Tab(1).Control(30)=   "Label1(18)"
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "Presentación IVA"
      TabPicture(2)   =   "frmempresa.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(35)"
      Tab(2).Control(1)=   "Text1(34)"
      Tab(2).Control(2)=   "Text1(33)"
      Tab(2).Control(3)=   "Text1(32)"
      Tab(2).Control(4)=   "Text1(37)"
      Tab(2).Control(5)=   "Text1(31)"
      Tab(2).Control(6)=   "Text1(30)"
      Tab(2).Control(7)=   "Text1(29)"
      Tab(2).Control(8)=   "Text1(28)"
      Tab(2).Control(9)=   "Text1(27)"
      Tab(2).Control(10)=   "Label1(29)"
      Tab(2).Control(11)=   "Line1"
      Tab(2).Control(12)=   "Label4(6)"
      Tab(2).Control(13)=   "Label4(5)"
      Tab(2).Control(14)=   "Label4(4)"
      Tab(2).Control(15)=   "Label4(3)"
      Tab(2).Control(16)=   "Label4(2)"
      Tab(2).Control(17)=   "Label4(1)"
      Tab(2).Control(18)=   "Label4(0)"
      Tab(2).Control(19)=   "Label1(27)"
      Tab(2).ControlCount=   20
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   39
         Left            =   -74760
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "NIF|T|S|||empresa2|siglaempre|||"
         Text            =   "Text1"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   17
         Left            =   6900
         MaxLength       =   8
         TabIndex        =   13
         Tag             =   "Digitos 10º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3930
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   16
         Left            =   6900
         MaxLength       =   8
         TabIndex        =   12
         Tag             =   "Digitos 9º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   15
         Left            =   6900
         MaxLength       =   8
         TabIndex        =   11
         Tag             =   "Digitos 8º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3030
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   14
         Left            =   6900
         MaxLength       =   8
         TabIndex        =   10
         Tag             =   "Digitos 7º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   2595
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   6900
         MaxLength       =   8
         TabIndex        =   9
         Tag             =   "Digitos 6º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   2145
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   12
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   8
         Tag             =   "Digitos 5º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3930
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   11
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   7
         Tag             =   "Digitos 4º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   10
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "Digitos 3er nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3030
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   9
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "Digitos 2º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   2595
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   8
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "Digitos 1er nivel|N|N|||||||"
         Text            =   "Text1"
         Top             =   2145
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1485
         MaxLength       =   1
         TabIndex        =   3
         Tag             =   "Numero niveles|N|N|||||||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   480
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   5685
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1695
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   38
         Left            =   -68640
         TabIndex        =   67
         Tag             =   "NIF|T|S|||empresa2|codigo||S|"
         Text            =   "CODIGO"
         Top             =   600
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   35
         Left            =   -72180
         MaxLength       =   4
         TabIndex        =   35
         Tag             =   "banco2|N|S|||empresa2|banco2|0000||"
         Text            =   "Text1"
         Top             =   3180
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   34
         Left            =   -71460
         MaxLength       =   4
         TabIndex        =   36
         Tag             =   "oficina ing|N|S|||empresa2|oficina2|0000||"
         Text            =   "Text1"
         Top             =   3180
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   33
         Left            =   -70740
         MaxLength       =   2
         TabIndex        =   37
         Tag             =   "Digi. cotnrol ing.|T|S|||empresa2|dc2|||"
         Text            =   "Text1"
         Top             =   3180
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   32
         Left            =   -70260
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "Cta ingreso|N|S|||empresa2|cuenta2|0000000000||"
         Text            =   "Text1"
         Top             =   3180
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   37
         Left            =   -69900
         MaxLength       =   4
         TabIndex        =   30
         Tag             =   "Letras|T|S|||empresa2|letraseti|||"
         Text            =   "Text1"
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   36
         Left            =   -70440
         MaxLength       =   9
         TabIndex        =   17
         Tag             =   "Código postal|T|S|||empresa2|telefono|||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1545
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   -74040
         TabIndex        =   15
         Tag             =   "NIF|T|S|||empresa2|nifempre|||"
         Text            =   "Text1"
         Top             =   840
         Width           =   1110
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   31
         Left            =   -70260
         MaxLength       =   10
         TabIndex        =   34
         Tag             =   "Cta devolucion|N|S|||empresa2|cuenta1|0000000000||"
         Text            =   "Text1"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   30
         Left            =   -70740
         MaxLength       =   2
         TabIndex        =   33
         Tag             =   "Digi. cotnrol dev.|T|S|||empresa2|dc1|00||"
         Text            =   "Text1"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   29
         Left            =   -71460
         MaxLength       =   4
         TabIndex        =   32
         Tag             =   "oficina dev|N|S|||empresa2|oficina1|0000||"
         Text            =   "Text1"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   28
         Left            =   -72180
         MaxLength       =   4
         TabIndex        =   31
         Tag             =   "banco1|N|S|||empresa2|banco1|0000||"
         Text            =   "Text1"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   27
         Left            =   -72300
         MaxLength       =   5
         TabIndex        =   29
         Tag             =   "Admon|T|S|||empresa2|administracion|||"
         Text            =   "Text1"
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   26
         Left            =   -70440
         MaxLength       =   9
         TabIndex        =   19
         Tag             =   "Tfno|T|S|||empresa2|tfnocontacto|||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   1545
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   25
         Left            =   -68040
         MaxLength       =   2
         TabIndex        =   25
         Tag             =   "pta|T|S|||empresa2|puerta|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   24
         Left            =   -68640
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "Piso|T|S|||empresa2|piso|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   23
         Left            =   -69240
         MaxLength       =   2
         TabIndex        =   23
         Tag             =   "Esca|T|S|||empresa2|escalera|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   22
         Left            =   -69960
         MaxLength       =   5
         TabIndex        =   22
         Tag             =   "Numero|T|S|||empresa2|numero|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   21
         Left            =   -74760
         MaxLength       =   2
         TabIndex        =   20
         Tag             =   "Via|T|S|||empresa2|siglasvia|||"
         Text            =   "Text1"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   19
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   18
         Tag             =   "Contacto|T|S|||empresa2|contacto|||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   3960
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   -74160
         MaxLength       =   30
         TabIndex        =   21
         Tag             =   "Dirección|T|S|||empresa2|direccion|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   3990
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   -70905
         TabIndex        =   27
         Tag             =   "Código postal|T|S|||empresa2|codpos|||"
         Text            =   "Text1"
         Top             =   3975
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   26
         Tag             =   "Población|T|S|||empresa2|poblacion|||"
         Text            =   "Text1"
         Top             =   3990
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   -69585
         MaxLength       =   30
         TabIndex        =   28
         Tag             =   "Provincia|T|S|||empresa2|provincia|||"
         Text            =   "Text1"
         Top             =   3990
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   18
         Left            =   -74745
         MaxLength       =   30
         TabIndex        =   16
         Tag             =   "Apoderado|T|S|||empresa2|apoderado|||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   3960
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   255
         Index           =   20
         Left            =   -74040
         TabIndex        =   64
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Siglas"
         Height          =   255
         Index           =   30
         Left            =   -74760
         TabIndex        =   83
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "DIGITOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2535
         TabIndex        =   82
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "10º Nivel"
         Height          =   195
         Index           =   17
         Left            =   6150
         TabIndex        =   81
         Top             =   3975
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "9º Nivel"
         Height          =   195
         Index           =   16
         Left            =   6150
         TabIndex        =   80
         Top             =   3540
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "8º Nivel"
         Height          =   195
         Index           =   15
         Left            =   6150
         TabIndex        =   79
         Top             =   3090
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "7º Nivel"
         Height          =   195
         Index           =   14
         Left            =   6150
         TabIndex        =   78
         Top             =   2655
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "6º Nivel"
         Height          =   195
         Index           =   13
         Left            =   6150
         TabIndex        =   77
         Top             =   2220
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "5º Nivel"
         Height          =   195
         Index           =   12
         Left            =   3750
         TabIndex        =   76
         Top             =   3990
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "4º Nivel"
         Height          =   195
         Index           =   11
         Left            =   3750
         TabIndex        =   75
         Top             =   3540
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "3er Nivel"
         Height          =   195
         Index           =   10
         Left            =   3750
         TabIndex        =   74
         Top             =   3105
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "2º Nivel"
         Height          =   195
         Index           =   9
         Left            =   3750
         TabIndex        =   73
         Top             =   2655
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "1er Nivel"
         Height          =   195
         Index           =   8
         Left            =   3750
         TabIndex        =   72
         Top             =   2220
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Nº de niveles"
         Height          =   255
         Index           =   7
         Left            =   405
         TabIndex        =   71
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Abreviado"
         Height          =   255
         Index           =   4
         Left            =   5685
         TabIndex        =   70
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre empresa"
         Height          =   255
         Index           =   1
         Left            =   1695
         TabIndex        =   69
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. empresa"
         Height          =   255
         Index           =   0
         Left            =   375
         TabIndex        =   68
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Letras etiqueta"
         Height          =   195
         Index           =   29
         Left            =   -71100
         TabIndex        =   66
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   195
         Index           =   28
         Left            =   -70440
         TabIndex        =   65
         Top             =   1320
         Width           =   630
      End
      Begin VB.Line Line1 
         X1              =   -73740
         X2              =   -68820
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta"
         Height          =   195
         Index           =   6
         Left            =   -70020
         TabIndex        =   63
         Top             =   2280
         Width           =   510
      End
      Begin VB.Label Label4 
         Caption         =   "D.C."
         Height          =   195
         Index           =   5
         Left            =   -70740
         TabIndex        =   62
         Top             =   2280
         Width           =   315
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal"
         Height          =   195
         Index           =   4
         Left            =   -71460
         TabIndex        =   61
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Entidad"
         Height          =   195
         Index           =   3
         Left            =   -72180
         TabIndex        =   60
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta bancaria:"
         Height          =   255
         Index           =   2
         Left            =   -73740
         TabIndex        =   59
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Banco ingreso"
         Height          =   255
         Index           =   1
         Left            =   -73680
         TabIndex        =   58
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Banco devolución"
         Height          =   255
         Index           =   0
         Left            =   -73680
         TabIndex        =   57
         Top             =   2700
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Administración"
         Height          =   195
         Index           =   27
         Left            =   -73980
         TabIndex        =   56
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   195
         Index           =   26
         Left            =   -70440
         TabIndex        =   55
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Pta"
         Height          =   255
         Index           =   25
         Left            =   -68040
         TabIndex        =   54
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Piso"
         Height          =   255
         Index           =   24
         Left            =   -68640
         TabIndex        =   53
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Esca."
         Height          =   255
         Index           =   23
         Left            =   -69240
         TabIndex        =   52
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Num."
         Height          =   255
         Index           =   22
         Left            =   -69960
         TabIndex        =   51
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Via"
         Height          =   255
         Index           =   21
         Left            =   -74760
         TabIndex        =   50
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Persona de contacto"
         Height          =   240
         Index           =   19
         Left            =   -74760
         TabIndex        =   49
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección"
         Height          =   255
         Index           =   2
         Left            =   -74160
         TabIndex        =   48
         Top             =   2955
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. pos"
         Height          =   255
         Index           =   3
         Left            =   -70905
         TabIndex        =   47
         Top             =   3735
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   46
         Top             =   3750
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   240
         Index           =   6
         Left            =   -69585
         TabIndex        =   45
         Top             =   3750
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre apoderado"
         Height          =   240
         Index           =   18
         Left            =   -74745
         TabIndex        =   44
         Top             =   1320
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5640
      Top             =   840
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5955
      TabIndex        =   39
      Top             =   5400
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
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
      Left            =   120
      TabIndex        =   41
      Top             =   5400
      Width           =   2310
   End
End
Attribute VB_Name = "frmempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public PrimeraConfiguracion As Boolean

Dim RS As ADODB.Recordset
Dim Modo As Byte

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1


    Select Case Modo
    Case 0
    
        
    Case 1
        
        If DatosOk Then
            If vEmpresa Is Nothing Then InsertarModificarEmpresa True
            If InsertarDesdeForm(Me) Then PonerModo 0
        End If
    
    Case 2
            'Modificar
            If DatosOk Then
                InsertarModificarEmpresa False
                '-----------------------------------------
                'Hacemos insertar
                If Adodc1.Recordset.EOF Then
                    I = InsertarDesdeForm(Me)
                Else
                    I = ModificaDesdeFormulario(Me)
                End If
                If I = -1 Then PonerModo 0
            End If

    End Select

        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub





Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
End Sub


Private Sub cmdCancelar_Click()
Select Case Modo
Case 0
   
Case 1
    PonerModo 1
Case 2
    PonerCampos
    PonerModo 0
End Select
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Top = 200
    Me.Left = 400
    Limpiar Me
    'Lista imagen
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4
        .Buttons(2).Image = 15
    End With
    Text1(0).Enabled = False
    'If PrimeraConfiguracion Then
    '    Text1(0).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
    'End If
    Me.SSTab1.Tab = 0
    'ASignamos las imagenes
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = "Select * from empresa2"
    Adodc1.Refresh
    If vEmpresa Is Nothing Then
        'No hay datos
        PonerModo 1
                
        'SQl
        Me.Tag = "select * from usuarios.empresas where conta='" & vUsu.CadenaConexion & "'"
        Set RS = New ADODB.Recordset
        RS.Open Me.Tag, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If RS.EOF Then
            MsgBox "Error fatal.  ---  NO HAY EMPRESA ---", vbCritical
            End
            Exit Sub
        End If
        Text1(0).Text = RS!codempre
        Text1(1).Text = RS!nomempre
        Text1(2).Text = RS!nomresum
        RS.Close
    Else
        PonerCampos
        PonerModo 0
    End If
    If Adodc1.Recordset.EOF Then Text1(38).Text = "1"  'Codigo para la tabla 2 de empresa
    If Toolbar1.Buttons(1).Enabled Then _
        Toolbar1.Buttons(1).Enabled = (vUsu.Nivel <= 1)
    cmdAceptar.Enabled = (vUsu.Nivel <= 1)
End Sub


'

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
    
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim mTag As CTag
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
    FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Modo = Kmodo
    Select Case Kmodo
    Case 0
        'Preparamos para ver los datos
        For I = 1 To Text1.Count - 1
            Text1(I).Locked = True
        Next I
        Label3.Caption = ""
    Case 1
            'Preparamos para que pueda insertar
        For I = 1 To Text1.Count - 1
            Text1(I).Text = ""
            Text1(I).Locked = False
        Next I
        Label3.Caption = "INSERTAR"
        Label3.ForeColor = vbBlue
    Case 2
        For I = 1 To Text1.Count - 1
            Text1(I).Locked = False
        Next I
        Label3.Caption = "MODIFICAR"
        Label3.ForeColor = vbRed
    End Select
    Me.Toolbar1.Buttons(1).Enabled = Modo <> 1
    cmdAceptar.Visible = Modo > 0
    cmdCancelar.Visible = Modo > 0
End Sub

Private Sub PonerCampos()
    If Not vEmpresa Is Nothing Then
        With vEmpresa
            Text1(0).Text = .codempre
            Text1(1).Text = .nomempre
            Text1(2).Text = .nomresum
            Text1(7).Text = .numnivel
            Text1(8).Text = .numdigi1
            Text1(9).Text = PonTextoNivel(.numdigi2)
            Text1(10).Text = PonTextoNivel(.numdigi3)
            Text1(11).Text = PonTextoNivel(.numdigi4)
            Text1(12).Text = PonTextoNivel(.numdigi5)
            Text1(13).Text = PonTextoNivel(.numdigi6)
            Text1(14).Text = PonTextoNivel(.numdigi7)
            Text1(15).Text = PonTextoNivel(.numdigi8)
            Text1(16).Text = PonTextoNivel(.numdigi9)
            Text1(17).Text = PonTextoNivel(.numdigi10)
        End With
            
    Else
        Limpiar Me
    End If
    If Adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Adodc1
End Sub

Private Function PonTextoNivel(Nivel As Integer) As String
If Nivel <> 0 Then
    PonTextoNivel = Nivel
Else
    PonTextoNivel = ""
End If
End Function


Private Function DatosOk() As Boolean
    Dim RS As ADODB.Recordset
    Dim B As Boolean
    Dim I As Integer
    Dim J As Integer
    Dim k As Integer
    
    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Otras cosas importantes
    'Comprobamos que tienen n niveles, y solo n
    J = CInt(Text1(7).Text)
    k = 0
    For I = 8 To 17
        If Text1(I).Text <> "" Then k = k + 1
    Next I
    If k <> J Then
        MsgBox "Niveles contables: " & J & vbCrLf & "Niveles parametrizados: " & k, vbExclamation
        Exit Function
    End If
    
    'K los niveles sean consecitivos sin saltar ninguno y sin ser menor
    J = 1
    k = CInt(Text1(8).Text)
    For I = 9 To 17
        If Text1(I).Text = "" Then
            J = 1000 + I
        Else
            J = CInt(Text1(I).Text)
        End If
        If J <= k Then
            MsgBox "Error en la asignacion de niveles. ", vbExclamation
            Exit Function
        End If
        k = J
    Next I
    
    
    
    DatosOk = True
End Function




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        'A modificar
        PonerModo 2
    Case 2
        Unload Me
    End Select
End Sub


Private Sub InsertarModificarEmpresa(Insertar As Boolean)

If Insertar Then Set vEmpresa = New Cempresa
With vEmpresa
    If Insertar Then .codempre = Val(Text1(0).Text)
    .nomempre = Text1(1).Text
    .nomresum = Text1(2).Text
    .numnivel = Val(Text1(7).Text)
    .numdigi1 = Val(Text1(8).Text)
    .numdigi2 = Val(Text1(9).Text)
    .numdigi3 = Val(Text1(10).Text)
    .numdigi4 = Val(Text1(11).Text)
    .numdigi5 = Val(Text1(12).Text)
    .numdigi6 = Val(Text1(13).Text)
    .numdigi7 = Val(Text1(14).Text)
    .numdigi8 = Val(Text1(15).Text)
    .numdigi9 = Val(Text1(16).Text)
    .numdigi10 = Val(Text1(17).Text)
    If Insertar Then
        If .Agregar = 1 Then
            MsgBox "Error fatal insertando datos empresa.", vbCritical
            End
        End If
    Else
        .Modificar
    End If
End With
End Sub


