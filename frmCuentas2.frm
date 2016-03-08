VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos cuentas"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9015
   Icon            =   "frmCuentas2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.CommandButton cmdCopiarDatos 
      Height          =   300
      Index           =   2
      Left            =   1080
      Picture         =   "frmCuentas2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "copiar cuentas OTRA SECCION/EMPRESA"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdCopiarDatos 
      Height          =   300
      Index           =   0
      Left            =   720
      Picture         =   "frmCuentas2.frx":6B5C
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Copiar cuenta"
      Top             =   120
      Width           =   300
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   31
      Top             =   240
      Width           =   1530
      Begin VB.CheckBox chkUltimo 
         Caption         =   "Ultimo nivel"
         Height          =   300
         Left            =   0
         TabIndex        =   32
         Top             =   210
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   11
         Left            =   360
         MaxLength       =   30
         TabIndex        =   34
         Tag             =   "Ultimo nbivel|T|N|||cuentas|apudirec|||"
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   3900
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7800
      TabIndex        =   25
      Top             =   7200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6600
      TabIndex        =   23
      Top             =   7200
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   7200
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   120
      TabIndex        =   29
      Top             =   7080
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   255
         Width           =   2955
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   8865
      Begin TabDlg.SSTab SSTab1 
         Height          =   6135
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos cuentas con apuntes directos"
         TabPicture(0)   =   "frmCuentas2.frx":D3AE
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Text1(29)"
         Tab(0).Control(1)=   "cboPais"
         Tab(0).Control(2)=   "Text1(23)"
         Tab(0).Control(3)=   "Text1(10)"
         Tab(0).Control(4)=   "Text1(9)"
         Tab(0).Control(5)=   "Text1(8)"
         Tab(0).Control(6)=   "Text1(7)"
         Tab(0).Control(7)=   "Text1(6)"
         Tab(0).Control(8)=   "Text1(5)"
         Tab(0).Control(9)=   "Text1(4)"
         Tab(0).Control(10)=   "Text1(3)"
         Tab(0).Control(11)=   "Text1(2)"
         Tab(0).Control(12)=   "Check1"
         Tab(0).Control(13)=   "Text1(12)"
         Tab(0).Control(14)=   "Text1(13)"
         Tab(0).Control(15)=   "Text1(14)"
         Tab(0).Control(16)=   "Text1(15)"
         Tab(0).Control(17)=   "Text1(16)"
         Tab(0).Control(18)=   "Label1(26)"
         Tab(0).Control(19)=   "imgppal(2)"
         Tab(0).Control(20)=   "Line2"
         Tab(0).Control(21)=   "Label1(22)"
         Tab(0).Control(22)=   "Label1(10)"
         Tab(0).Control(23)=   "Label1(9)"
         Tab(0).Control(24)=   "Label1(8)"
         Tab(0).Control(25)=   "Label1(6)"
         Tab(0).Control(26)=   "Label1(5)"
         Tab(0).Control(27)=   "Label1(4)"
         Tab(0).Control(28)=   "Label1(3)"
         Tab(0).Control(29)=   "Label1(7)"
         Tab(0).Control(30)=   "Label1(2)"
         Tab(0).Control(31)=   "Label1(11)"
         Tab(0).Control(32)=   "Label1(12)"
         Tab(0).Control(33)=   "Label1(13)"
         Tab(0).Control(34)=   "Label1(14)"
         Tab(0).Control(35)=   "Label1(15)"
         Tab(0).ControlCount=   36
         TabCaption(1)   =   "Arimoney"
         TabPicture(1)   =   "frmCuentas2.frx":D3CA
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label1(25)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label3(0)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label1(16)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Line1"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label1(17)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label1(18)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label1(19)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label1(20)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "imgppal(0)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "imgppal(1)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label3(1)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label3(2)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label3(3)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label1(21)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Label1(24)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Image1(1)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Image1(0)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Label3(4)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "imgppal(3)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Label3(5)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "Label1(27)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "imgppal(4)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "Label1(28)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "Text1(19)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "Text1(20)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "Text1(21)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "Text1(22)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "Text1(25)"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "Text1(26)"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "Text2(0)"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "Text2(1)"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "Text1(27)"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "Text1(17)"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "Text1(18)"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "Text1(28)"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "Text1(30)"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "Text1(31)"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).ControlCount=   37
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   31
            Left            =   3120
            MaxLength       =   35
            TabIndex        =   53
            Tag             =   "GT|T|S|||cuentas|SEPA_Refere|||"
            Top             =   1440
            Width           =   1905
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   30
            Left            =   6720
            TabIndex        =   54
            Tag             =   "Fecha|F|S|||cuentas|SEPA_FecFirma|dd/mm/yyyy||"
            Top             =   1440
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   29
            Left            =   -74040
            MaxLength       =   4
            TabIndex        =   10
            Tag             =   "Iban|T|S|||cuentas|iban|||"
            Text            =   "Text1"
            Top             =   2730
            Width           =   690
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   28
            Left            =   1680
            TabIndex        =   59
            Tag             =   "F. baja credito|F|S|||cuentas|fecbajcre|dd/mm/yyyy||"
            Top             =   3720
            Width           =   1185
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            Left            =   -70200
            TabIndex        =   9
            Text            =   "Combo1"
            Top             =   2130
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   5280
            TabIndex        =   57
            Tag             =   "Fl|F|S|||cuentas|fecsolic|dd/mm/yyyy||"
            Top             =   3000
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   17
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   56
            Tag             =   "Razón social|T|S|||cuentas|numpoliz|||"
            Top             =   3000
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   27
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   55
            Tag             =   "GT|T|S|||cuentas|grupotesoreria|||"
            Top             =   1920
            Width           =   2625
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   4560
            TabIndex        =   81
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   4560
            TabIndex        =   80
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   26
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   52
            Tag             =   "Cta banco|T|S|||cuentas|ctabanco|||"
            Top             =   960
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   25
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   51
            Tag             =   "For. pago|N|S|||cuentas|forpa|||"
            Text            =   "123456789012345678901234567890"
            Top             =   480
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   23
            Left            =   -67800
            MaxLength       =   15
            TabIndex        =   18
            Tag             =   "NIF|F|S|||cuentas|fecbloq|||"
            Text            =   "Text1"
            Top             =   4920
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            Height          =   1635
            Index           =   22
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Tag             =   "Razón social|T|S|||cuentas|observa|||"
            Text            =   "frmCuentas2.frx":D3E6
            Top             =   4200
            Width           =   6825
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   21
            Left            =   7200
            TabIndex        =   61
            Tag             =   "lmpor1|N|S|||cuentas|credicon|#0.00||"
            Top             =   3720
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   20
            Left            =   5280
            TabIndex        =   60
            Tag             =   "Fecha|F|S|||cuentas|fecconce|dd/mm/yyyy||"
            Top             =   3720
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   19
            Left            =   7200
            TabIndex        =   58
            Tag             =   "Imp1|N|S|||cuentas|credisol|#0.00||"
            Top             =   3000
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Height          =   1635
            Index           =   10
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
            Text            =   "frmCuentas2.frx":D3F1
            Top             =   4260
            Width           =   6825
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   -70785
            MaxLength       =   50
            TabIndex        =   16
            Tag             =   "Direccion web|T|S|||cuentas|webdatos|||"
            Text            =   "Text1"
            Top             =   3570
            Width           =   4320
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   -74865
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "E-Mail|T|S|||cuentas|maidatos|||"
            Text            =   "Text1"
            Top             =   3555
            Width           =   3765
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   -68880
            MaxLength       =   15
            TabIndex        =   3
            Tag             =   "NIF|T|S|||cuentas|nifdatos|||"
            Text            =   "Text1"
            Top             =   675
            Width           =   1845
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   -73320
            MaxLength       =   30
            TabIndex        =   8
            Tag             =   "Provincia|T|S|||cuentas|desprovi|||"
            Text            =   "Text1"
            Top             =   2130
            Width           =   2850
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   5
            Left            =   -70800
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "Población|T|S|||cuentas|despobla|||"
            Text            =   "Text1"
            Top             =   1320
            Width           =   4320
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   4
            Left            =   -74865
            MaxLength       =   6
            TabIndex        =   7
            Tag             =   "Cod. Postal|T|S|||cuentas|codposta|||"
            Text            =   "Text1"
            Top             =   2130
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   3
            Left            =   -74880
            MaxLength       =   30
            TabIndex        =   5
            Tag             =   "Domicilio|T|S|||cuentas|dirdatos|||"
            Text            =   "Text1"
            Top             =   1320
            Width           =   3960
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   2
            Left            =   -74880
            MaxLength       =   60
            TabIndex        =   2
            Tag             =   "Razón social|T|S|||cuentas|razosoci|||"
            Top             =   675
            Width           =   5865
         End
         Begin VB.CheckBox Check1 
            Caption         =   "347"
            Height          =   225
            Left            =   -66960
            TabIndex        =   4
            Tag             =   "Modelo|N|S|||cuentas|model347|||"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   12
            Left            =   -70200
            MaxLength       =   15
            TabIndex        =   19
            Tag             =   "Pais|T|S|||cuentas|pais|||"
            Text            =   "Text1"
            Top             =   2130
            Width           =   3330
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   13
            Left            =   -73305
            MaxLength       =   4
            TabIndex        =   11
            Tag             =   "entidad|T|S|||cuentas|entidad|||"
            Text            =   "Text1"
            Top             =   2730
            Width           =   570
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   14
            Left            =   -72720
            MaxLength       =   4
            TabIndex        =   12
            Tag             =   "oficina|T|S|||cuentas|oficina|||"
            Text            =   "Text1"
            Top             =   2730
            Width           =   570
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   15
            Left            =   -72120
            MaxLength       =   2
            TabIndex        =   13
            Tag             =   "cc|T|S|||cuentas|cc|||"
            Text            =   "Text1"
            Top             =   2730
            Width           =   450
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   -71640
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Cta. banco|T|S|||cuentas|cuentaba|||"
            Text            =   "9999999999"
            Top             =   2730
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha firma"
            Height          =   195
            Index           =   28
            Left            =   5280
            TabIndex        =   88
            Top             =   1440
            Width           =   930
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   4
            Left            =   6360
            Picture         =   "frmCuentas2.frx":D3F7
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Referencia"
            Height          =   195
            Index           =   27
            Left            =   1680
            TabIndex        =   87
            Top             =   1440
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Mandato SEPA"
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
            Index           =   5
            Left            =   120
            TabIndex        =   86
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   26
            Left            =   -74880
            TabIndex        =   85
            Top             =   2760
            Width           =   705
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   3
            Left            =   2640
            Picture         =   "frmCuentas2.frx":D482
            Top             =   3480
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Grupo tesoreria"
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
            Index           =   4
            Left            =   120
            TabIndex        =   82
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   2760
            Picture         =   "frmCuentas2.frx":D50D
            Top             =   480
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   2760
            Picture         =   "frmCuentas2.frx":DF0F
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta banco"
            Height          =   195
            Index           =   24
            Left            =   1680
            TabIndex        =   79
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Forma pago"
            Height          =   195
            Index           =   21
            Left            =   1680
            TabIndex        =   78
            Top             =   480
            Width           =   915
         End
         Begin VB.Label Label3 
            Caption         =   "Vencimientos"
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
            Index           =   3
            Left            =   120
            TabIndex        =   77
            Top             =   480
            Width           =   2295
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   2
            Left            =   -66840
            Picture         =   "frmCuentas2.frx":E911
            Top             =   4560
            Width           =   240
         End
         Begin VB.Line Line2 
            X1              =   -67920
            X2              =   -67920
            Y1              =   4200
            Y2              =   5880
         End
         Begin VB.Label Label1 
            Caption         =   "Bloqueo cta"
            Height          =   255
            Index           =   22
            Left            =   -67800
            TabIndex        =   71
            Top             =   4560
            Width           =   960
         End
         Begin VB.Label Label3 
            Caption         =   "CONCEDIDO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   4080
            TabIndex        =   70
            Top             =   3720
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "SOLICITADO"
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
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   69
            Top             =   2520
            Width           =   1125
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   1
            Left            =   5880
            Picture         =   "frmCuentas2.frx":E99C
            Top             =   3480
            Width           =   240
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   0
            Left            =   5880
            Picture         =   "frmCuentas2.frx":EA27
            Top             =   2760
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Importe"
            Height          =   195
            Index           =   20
            Left            =   7200
            TabIndex        =   68
            Top             =   3480
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   19
            Left            =   5280
            TabIndex        =   67
            Top             =   3480
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Importe"
            Height          =   195
            Index           =   18
            Left            =   7200
            TabIndex        =   66
            Top             =   2760
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   17
            Left            =   5280
            TabIndex        =   65
            Top             =   2760
            Width           =   915
         End
         Begin VB.Line Line1 
            X1              =   2640
            X2              =   8400
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Poliza"
            Height          =   195
            Index           =   16
            Left            =   1680
            TabIndex        =   63
            Top             =   2760
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Observaciones"
            Height          =   255
            Index           =   10
            Left            =   -74865
            TabIndex        =   50
            Top             =   4050
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Dirección web"
            Height          =   255
            Index           =   9
            Left            =   -70785
            TabIndex        =   49
            Top             =   3330
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "e-MAIL"
            Height          =   255
            Index           =   8
            Left            =   -74865
            TabIndex        =   48
            Top             =   3330
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   6
            Left            =   -73305
            TabIndex        =   47
            Top             =   1890
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   195
            Index           =   5
            Left            =   -70800
            TabIndex        =   46
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "C.Postal"
            Height          =   255
            Index           =   4
            Left            =   -74880
            TabIndex        =   45
            Top             =   1890
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   195
            Index           =   3
            Left            =   -74865
            TabIndex        =   44
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "N.I.F."
            Height          =   255
            Index           =   7
            Left            =   -68880
            TabIndex        =   43
            Top             =   480
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Razón social"
            Height          =   195
            Index           =   2
            Left            =   -74880
            TabIndex        =   42
            Top             =   390
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "País"
            Height          =   255
            Index           =   11
            Left            =   -70200
            TabIndex        =   41
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Entidad"
            Height          =   255
            Index           =   12
            Left            =   -74145
            TabIndex        =   40
            Top             =   2580
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Oficina"
            Height          =   255
            Index           =   13
            Left            =   -73320
            TabIndex        =   39
            Top             =   2580
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "C.C"
            Height          =   255
            Index           =   14
            Left            =   -72480
            TabIndex        =   38
            Top             =   2580
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta "
            Height          =   255
            Index           =   15
            Left            =   -71880
            TabIndex        =   37
            Top             =   2580
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Operaciones aseguradas"
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
            TabIndex        =   64
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha BAJA"
            Height          =   195
            Index           =   25
            Left            =   1680
            TabIndex        =   84
            Top             =   3480
            Width           =   915
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Codigo cuenta|T|N|||cuentas|codmacta||S|"
      Top             =   390
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1440
      MaxLength       =   60
      TabIndex        =   1
      Tag             =   "Denominación cuenta|T|N|||cuentas|nommacta|||"
      Top             =   405
      Width           =   5940
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   1635
      Index           =   24
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Text            =   "frmCuentas2.frx":EAB2
      Top             =   3000
      Width           =   6825
   End
   Begin VB.Frame FrGranEmpresa 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   73
      Top             =   4920
      Width           =   8055
      Begin VB.CommandButton cmdCopiarDatos 
         Height          =   375
         Index           =   1
         Left            =   3840
         Picture         =   "frmCuentas2.frx":EABA
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtRegularizacion 
         Height          =   285
         Left            =   4320
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Grandes empresas.   Regularización grupos 7 y 8"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Label lbl347 
      Caption         =   "Ofertar la marca de 347 para las cuentas del subgrupo"
      Height          =   195
      Left            =   600
      TabIndex        =   76
      Top             =   2280
      Width           =   5520
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   23
      Left            =   255
      TabIndex        =   72
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "NO es cuenta último nivel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   33
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   27
      Top             =   165
      Width           =   3465
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   165
      Width           =   735
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Public CodCta As String
Public vModo As Byte
' 0.- Ver solo
' 1.- Añadir
' 2.- Modificar
' 3.- Buscar
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private kCampo As Integer
Dim SQL As String


'Para saber si han bloquedao una cuenta, si tienen que avisar de
Private varBloqCta As String



Private Sub cboPais_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case vModo
    Case 1
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            
            'estoy aqui, da problemas, creo que es el  chcek para indicar si es ultimomnivel o no
            If InsertarDesdeForm(Me) Then
                
                If Len(Text1(0).Text) = vEmpresa.DigitosUltimoNivel Then
                           
                    If vParam.EnlazaCtasMultibase <> "" Then
                        Screen.MousePointer = vbHourglass
                        lblIndicador.Caption = "ENLACE GESTION"
                        Me.Refresh
                        DoEvents
                               'Cta                     nomcta              NIF
                        SQL = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(7).Text & "|"
                        HacerEnlaceMultibase 0, SQL
                    
                    End If
                    
                    
                    If Text1(23).Text <> varBloqCta Then
                        'Siginifica que el bloqueo de cuenta ha sido modificado
                        SQL = "Hay conectados los siguientes PCs. Deberian reiniciar." & vbCrLf
                        If UsuariosConectados(SQL) Then
                        
                        End If
                        'Volvemos a leer las cuentas bloqueadas
                        vParam.ObtenerCuentasBloqueadas
                    End If
                    
''''                    'Si es cuenta de ultimo nivel. Compruebo si la insercion tiene que ver
''''                    'con la variable GRAN EMPRESA
''''                    If Val(Mid(Text1(0).Text, 1, 1)) >= 8 Then
''''                        If Not vEmpresa.GranEmpresa Then vEmpresa.GranEmpresa = True
''''                    End If
                    
                End If
                'Salimos
                CadenaDesdeOtroForm = Text1(0).Text
                Unload Me
               
               
               
               
               
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos modificar
                
                If ModificaDesdeFormulario(Me) Then
                    'SOLO ACTAULZIAMOS CUENTAS DE ULTIMO NIVEL
                    If Len(Text1(0).Text) = vEmpresa.DigitosUltimoNivel Then
                        If vParam.EnlazaCtasMultibase <> "" Then
                            Screen.MousePointer = vbHourglass
                            lblIndicador.Caption = "ENLACE GESTION"
                            Me.Refresh
                            DoEvents
                                   'Cta                     nomcta              NIF
                            SQL = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(7).Text & "|"
                            HacerEnlaceMultibase 1, SQL
                        
                        End If
                    End If
                    
                    If Text1(23).Text <> varBloqCta Then
                        'Siginifica que el bloqueo de cuenta ha sido modificado
                        SQL = "Hay conectados los siguientes PCs. Deberian reiniciar." & vbCrLf
                        If UsuariosConectados(SQL) Then
                        
                        End If
                        'Volvemos a leer las cuentas bloqueadas
                        vParam.ObtenerCuentasBloqueadas
                    End If
                    CadenaDesdeOtroForm = Text1(0).Text
                    Unload Me
                End If
            End If
    Case 3
            'Si hay busqueda
            CadenaDesdeOtroForm = ""
            SQL = ObtenerBusqueda(Me)
            If SQL <> "" Then
                CadenaDesdeOtroForm = SQL
                Unload Me
            Else
                MsgBox "Especifique algun campo de búsqueda", vbExclamation
            End If
    End Select
    
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub



'0.- Cuenta normal
'1.- Forpa
'2.- Cuenta bancaria
Private Sub AbrirSelCuentas2(vOpcion As Byte, OtraSeccion As String)

    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    Select Case vOpcion
    Case 0
        

        frmB.vTabla = OtraSeccion & "cuentas"
        frmB.vSQL = "apudirec = ""S"""
        If OtraSeccion = "" Then
            frmB.vTitulo = "Cuentas"
        Else
            SQL = DevuelveDesdeBD("nomresum", OtraSeccion & "empresa", "1", "1")
            frmB.vTitulo = "Cuentas" & "  - " & SQL
        End If
        
        SQL = ParaGrid(Text1(0), 20, "Codigo Cta")
        SQL = SQL & ParaGrid(Text1(1), 55, "Descripcion")
        SQL = SQL & ParaGrid(Text1(7), 25)
        
    Case 1
        SQL = "Forma pago|codforpa|N|10·Descripcion|nomforpa|T|55·Tipo|descformapago|T|33·"
        frmB.vTabla = "sforpa,stipoformapago"
        frmB.vSQL = " sforpa.tipforpa=tipoformapago"
        frmB.vTitulo = "Formas de pago"
        
    Case 2
        'Cta bancaria
        SQL = "Cuenta|codmacta|N|20·Desc. banco|descripcion|T|61·"
        frmB.vTabla = "ctabancaria"
        frmB.vSQL = ""
        frmB.vTitulo = "Cuentas de BANCOS"
        
    End Select
    frmB.vDevuelve = "0|1|"
    frmB.vSelElem = 1
    frmB.VCampos = SQL
    SQL = ""
    frmB.Show vbModal
    Set frmB = Nothing

End Sub


Private Sub cmdCopiarDatos_Click(Index As Integer)
Dim EmpresaSt As String
    If Index = 0 Or Index = 2 Then
       If Not Frame1.visible Then
            MsgBox "Solo se pueden copiar datos para las cuentas a ultimo nivel", vbExclamation
            Exit Sub
        End If
    
    
    Else
        'Para poner contra que cuenta regularizan las 8 y 9
        
    End If
    
    
    EmpresaSt = ""
    
    If Index = 2 Then
        'Abrimos para que seleccione las empresas
            SQL = ""
            CadenaDesdeOtroForm = "NO"  'Para que no seleccione ninguna empresa por defecto
            frmMensajes.Opcion = 4
            frmMensajes.Show vbModal
            If CadenaDesdeOtroForm = "" Then Exit Sub
            NumRegElim = RecuperaValor(CadenaDesdeOtroForm, 1)
            If NumRegElim <> 1 Then
                SQL = "Seleccione una única empresa"
                
            Else
                EmpresaSt = RecuperaValor(CadenaDesdeOtroForm, 3)
                EmpresaSt = "conta" & EmpresaSt & "."
                
                CadenaDesdeOtroForm = DevuelveDesdeBD("numnivel", EmpresaSt & "empresa", "1", "1")
                If CadenaDesdeOtroForm = "" Then
                   SQL = "Error obteniendo datos empresa : " & EmpresaSt
                Else
                    CadenaDesdeOtroForm = "numdigi" & CadenaDesdeOtroForm
                    CadenaDesdeOtroForm = DevuelveDesdeBD(CadenaDesdeOtroForm, EmpresaSt & "empresa", "1", "1")
                    If CadenaDesdeOtroForm = "" Then
                        SQL = "Error obteniendo datos ultimo nivel: " & EmpresaSt
                    Else
                        If vEmpresa.DigitosUltimoNivel <> Val(CadenaDesdeOtroForm) Then
                            SQL = "Disitintos digitos ultimo nivel"
                        End If
                    End If
                End If
            End If
            
            If SQL <> "" Then
                MsgBox SQL, vbExclamation
                SQL = ""
                Exit Sub
            End If
                
    
    End If
    AbrirSelCuentas2 0, EmpresaSt  '0. Cuentas normal
    
    If SQL <> "" Then
        SQL = RecuperaValor(SQL, 1)
        'Ha devuelto datos
        Me.Refresh
        DoEvents
        Screen.MousePointer = vbHourglass
        
            
        If Index = 0 Or Index = 2 Then
            PonerDatosDeOtraCuenta EmpresaSt
            
        Else
            Me.txtRegularizacion.Text = SQL
        End If
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdRegresar_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub




Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    SSTab1.Tab = 0
    Me.SSTab1.TabVisible(1) = vEmpresa.TieneTesoreria
    Text1(0).Enabled = True
    Text1(0).MaxLength = vEmpresa.DigitosUltimoNivel
    EnablarText (vModo <> 0)
    cmdCopiarDatos(0).visible = vModo = 1
    cmdCopiarDatos(1).visible = vModo = 1 Or vModo = 2
    Me.imgppal(0).visible = vModo > 0
    Me.imgppal(1).visible = vModo > 0
    Me.imgppal(2).visible = vModo > 0
    Me.imgppal(3).visible = vModo > 0
    Me.imgppal(4).visible = vModo > 0
    FrGranEmpresa.visible = False
    
    If vModo = 1 Or vModo = 2 Then CargarComboPais
    
    
    Select Case vModo
    Case 0
            'Vamos a ver los datos
            PonerCampos ""
            
            lblIndicador.Caption = "Ver cuenta"
    Case 1
            LimpiarCampos
            If CodCta <> "" Then Text1(0).Text = CodCta
            '347
            Check1.Value = 1
            Frame1.visible = True
            Frame1.Enabled = False
            lblIndicador.Caption = "Insertar"
            
            Me.cmdCopiarDatos(2).visible = HayMasDeUnaEmpresa
            
    Case 2
            Text1(0).Enabled = False
            Text1(1).Enabled = True
            PonerCampos ""
            lblIndicador.Caption = "Modificar"
    Case 3
            LimpiarCampos
            Frame1.visible = True
            lblIndicador.Caption = "Búsqueda"
    End Select
    
    
    If vModo = 0 Or vModo = 2 Then
        If Text1(11).Text = "S" Then
            kCampo = vModo
            vModo = 2
            Text1_LostFocus 25
            Text1_LostFocus 26
            vModo = kCampo
            kCampo = 0
        End If
    End If

    
    If vModo = 1 Or (vModo = 2 And Text1(11).Text = "S") Then
        Me.cboPais.visible = True
        Me.Text1(12).Enabled = False
    Else
        Me.cboPais.visible = False
    End If
    
    If vModo = 2 Then
        Text1(0).BackColor = &H80000018
        Else
        Text1(0).BackColor = &H80000005
    End If
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    'Aqui va el especifico de cada form es
    '### a mano
    chkUltimo.Value = 0
End Sub

Private Sub PonerCampos(QueEmpresa As String)
Dim RS As ADODB.Recordset
Dim mTag As CTag
Dim I  As Integer
Dim T As Object
Dim Valor

    Set RS = New ADODB.Recordset
    SQL = "Select * from " & QueEmpresa & "cuentas where codmacta='" & CodCta & "'"
    RS.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If RS.EOF Then
        LimpiarCampos
        lblIndicador.Caption = "MODIFICAR"
    Else
        Set mTag = New CTag
        
      
        
        For I = 0 To Text1.Count - 1
            Set T = Text1(I)
            mTag.Cargar T
            If mTag.Cargado Then
                'Columna en la BD
                SQL = mTag.Columna
                If mTag.Vacio = "S" Then
                    Valor = DBLet(RS.Fields(SQL))
                Else
                    Valor = RS.Fields(SQL)
                End If
                If mTag.Formato <> "" Then Valor = Format(Valor, mTag.Formato)
                
                Text1(I).Text = Valor
            Else
                Text1(I).Text = ""
            End If
        Next I
        varBloqCta = ""
        If RS.Fields!apudirec = "S" Then
            chkUltimo.Value = 1
            Text1(11).Text = "S"
            Me.Frame1.visible = True
            varBloqCta = Text1(23).Text

            Else
            chkUltimo.Value = 0
            Frame1.visible = False
            Text1(24).Text = Text1(10).Text
            Text1(11).Text = "N"
        End If
        Check1.Value = RS!model347
        Check2.Value = Check1.Value
        Check2.Enabled = (vModo = 2)
        
        Check2.visible = (Len(Text1(0).Text) = 3)
        lbl347.visible = (Len(Text1(0).Text) = 3)
        
        PonerFrameGranEmpresa
        
        If vModo = 2 And chkUltimo.Value = 1 Then cboPais.Text = Text1(12).Text
        Set mTag = Nothing

    End If
End Sub

Private Sub PonerFrameGranEmpresa()
Dim B As Boolean
    
    B = False
    If vParam.GranEmpresa Then
        'y Si len 3 y cta 8 y 9
        If Len(Text1(0).Text) = 3 Then
            '8 y 9
            If Val(Mid(Text1(0), 1, 1)) >= 8 Then
                B = True
                'cuentaba en cuentas 7 y 8 a 3 digitos quiere decir DONDE regularizara
                txtRegularizacion.Text = Text1(16).Text
            End If
        End If
    End If
    Me.FrGranEmpresa.visible = B
End Sub
Private Sub frmB_Selecionado(CadenaDevuelta As String)
    SQL = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    imgppal(0).Tag = vFecha
End Sub



Private Sub Image1_Click(Index As Integer)

    AbrirSelCuentas2 1 + Index, "" '1.- Forpa    2.- Cta bancaria
    If SQL <> "" Then
        If Index = 1 Then
            'CUENTAS
            Text1(26).Text = RecuperaValor(SQL, 1)
            Text2(1).Text = RecuperaValor(SQL, 2)
        Else
            'FORPA
            Text1(25).Text = RecuperaValor(SQL, 1)
            Text2(0).Text = RecuperaValor(SQL, 2)
        End If
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
Dim Ix As Integer
    imgppal(0).Tag = ""
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    Select Case Index
    Case 0
        Ix = 18
    Case 1
        Ix = 20
    Case 3
        Ix = 28
    Case 4
        Ix = 30

    Case Else
        Ix = 23
    End Select
    
    If Text1(Ix).Text <> "" Then frmC.Fecha = CDate(Text1(Ix).Text)
    frmC.Show vbModal
    
    If imgppal(0).Tag <> "" Then Text1(Ix).Text = Format(imgppal(0).Tag, "dd/mm/yyyy")
        
    
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    
    If vModo = 3 Then
        Text1(kCampo).BackColor = vbWhite
        Text1(Index).BackColor = vbYellow
        Else
            If Index <> 10 And Index <> 22 Then PonFoco Text1(Index)
    End If
    kCampo = Index
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 10 And Index <> 22 And Index <> 24 Then KEYpress KeyAscii
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
    Dim I As Integer
    Dim SQL2 As String
    Dim mTag As CTag
    Dim Im As Currency
    
    If vModo = 3 Or vModo = 0 Then Exit Sub 'Busqueda avanzada o ver solo
    
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
    If Text1(Index).Text = "" Then
        If Index = 0 Then
            Frame1.visible = True
            chkUltimo.Value = 0
        ElseIf Index = 25 Then
            Text2(0).Text = ""
        ElseIf Index = 26 Then
            Text2(1).Text = ""
        End If
        Exit Sub
    End If
    If Index <> 10 And Index <> 24 Then Text1(Index).Text = UCase(Text1(Index).Text)
    'Si queremos hacer algo ..
    Select Case Index
        Case 0
            PierdeFocoCodigoCuenta
        Case 1
            If vModo = 1 Then
                If Text1(2).Text = "" Then Text1(2).Text = Text1(1).Text
                If Text1(12).Text = "" Then Text1(12).Text = "ESPAÑA"
            End If
        '....
        Case 13 To 16
            If vModo = 2 Then
                If Not IsNumeric(Text1(Index).Text) Then
                    Text1(Index).Text = ""
                    PonFoco Text1(Index)
                    Exit Sub
                End If
                If Index = 15 Then
                    I = 2
                Else
                    If Index = 16 Then
                        I = 10
                    Else
                        I = 4
                    End If
                End If
                SQL = Mid("0000000000", 1, I)
                Text1(Index).Text = Format(Text1(Index).Text, SQL)
                
                
                'IBAN
        
                SQL = ""
                For I = 13 To 16
                    SQL = SQL & Text1(I).Text
                Next
                
                If Len(SQL) = 20 Then
                    'OK. Calculamos el IBAN
                    
                    
                    If Text1(29).Text = "" Then
                        'NO ha puesto IBAN
                        If DevuelveIBAN2("ES", SQL, SQL) Then Text1(29).Text = "ES" & SQL
                    Else
                        SQL2 = CStr(Mid(Text1(29).Text, 1, 2))
                        If DevuelveIBAN2(CStr(SQL2), SQL, SQL) Then
                            If Mid(Text1(29).Text, 3) <> SQL Then
                                
                                MsgBox "Codigo IBAN distinto del calculado [" & SQL2 & SQL & "]", vbExclamation
                                'Text1(29).Text = "ES" & SQL
                            End If
                        End If
                    End If
                End If
                        
                
                
                
                
                
                
                
                
                
             End If
        Case 18, 20, 23, 28, 30
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
            End If
        
        Case 19, 21
            If Not CadenaCurrency(Text1(Index).Text, Im) Then
                MsgBox "Importe incorrecto: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
            Else
                Text1(Index).Text = Format(Im, FormatoImporte)
            End If
        Case 25
                SQL = ""
                
                If IsNumeric(Text1(25).Text) Then
                    SQL = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", Text1(25).Text, "N")
                Else
                    MsgBox "Campo debe ser numerico: " & Text1(25).Text, vbExclamation
                    Text1(25).Text = ""
                End If
    
            Text2(0).Text = SQL
            If SQL = "" Then PonleFoco Text1(25)
        Case 26
            SQL = Text1(26).Text
            If CuentaCorrectaUltimoNivel(SQL, SQL2) Then
                SQL = DevuelveDesdeBD("codmacta", "ctabancaria", "codmacta", SQL, "T")
                If SQL = "" Then
                    MsgBox "La cuenta NO pertenece a ningúna cta. bancaria", vbExclamation
                    SQL2 = ""
                    
                Else
                    'CORRECTO
                End If
            Else
                SQL = ""
                MsgBox SQL2, vbExclamation
                SQL2 = ""
            End If
            Text1(26).Text = SQL
            Text2(1).Text = SQL2
            If SQL = "" Then PonleFoco Text1(26)
    End Select
    '---
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim Nivel As Integer
Dim RC As Byte
Dim RC2 As String
    
    
    DatosOk = False
    
    Text1(1).Text = UCase(Text1(1).Text)
    Text1(2).Text = UCase(Text1(2).Text)
    
    
    
       
    'Asignamos el campo apunte directo
    If chkUltimo.Value = 0 Then
        Text1(11).Text = "N"
    Else
        Text1(11).Text = "S"
    End If
    
    B = True
    If Len(Text1(0).Text) = vEmpresa.DigitosUltimoNivel Then
        'Digitos ultimo nivel
        If chkUltimo.Value = 0 Then
            MsgBox "La longitud de la cuenta es de ultimo nivel y no esta marcado", vbExclamation
            B = False
        End If
    Else
        'No tiene longitud de ultimo nivel
        If chkUltimo.Value = 1 Then
            MsgBox "No  es cuenta de ultimo nivel pero esta marcado", vbExclamation
            B = False
        End If
        
    End If
    If Not B Then Exit Function
    
    
    
    If Len(Text1(0).Text) < vEmpresa.DigitosUltimoNivel Then
        Check1.Value = 0
        '--------------------------------
        'Si es nivel 3 entonces guardamos la oferta
        If Len(Text1(0).Text) = 3 Then
            Check1.Value = Check2.Value
            'Es gran empresa y digitos 8 9
            If Me.FrGranEmpresa.visible Then
            
                If Mid(txtRegularizacion.Text, 1, 1) <> "1" Then
                    MsgBox "La regularizacion será contra las cuentas del grupo 1", vbExclamation
                    Exit Function
                End If
            
                'Compruebo que la cuenta existe
                SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", txtRegularizacion.Text, "T")
                If SQL = "" Then
                    MsgBox "La cuenta " & txtRegularizacion.Text & " NO existe", vbExclamation
                    PonFoco txtRegularizacion
                    Exit Function
                End If
                Text1(16).Text = txtRegularizacion.Text
            End If
        End If
        
        'Si ha puesto observaciones las guardo
        Text1(10).Text = Text1(24).Text
    Else
        'Si estamos modificando o añadiendo, el pais(text1(12)  cogera el valor que tenga el combo
        Text1(12).Text = cboPais.Text
    End If
    
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
    
    If Not IsNumeric(Text1(0).Text) Then
        MsgBox "Campo cuenta debe ser numérico", vbExclamation
        Exit Function
    End If
    
    
    'Comprobamos de que nivel es la cuenta
    Nivel = NivelCuenta(Text1(0).Text)
    If Nivel < 1 Then
        MsgBox "El número de dígitos no pertenece a ningún nivel contable", vbExclamation
        Exit Function
    End If
    
    'NIF
    If Text1(7).Text <> "" Then
        'Ha escrito el NIF
        If cboPais.Text = "ESPAÑA" Then
            If Not Comprobar_NIF(Text1(7).Text) Then
                If MsgBox("NIF incorrecto. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
        End If
        'Comprobacion NIFs
        'Comprobaremos si el NIF existe en cualquier otra contabilidad
        'comprobando que tenga permisos para ello
        ComprobarNifTodasContas
    End If
    
    
    
    
    
    
    If Nivel > 1 Then
    
    
        B = ExistenSubcuentas(Text1(0).Text, Nivel - 1)
        If Not B Then
            RC = MsgBox("No existen subcuentas inferiores para la cuenta : " & Text1(0).Text & vbCrLf & "Desea crealas ?", vbQuestion + vbYesNoCancel)
            If RC = vbYes Then
                'Hay que crear subcuentas
                B = CreaSubcuentas(Text1(0).Text, Nivel - 1, Text1(1).Text)
                If Not B Then Exit Function
            Else
                Exit Function
            End If
        End If
        
        
        
        
        
        
    End If
    
    
    'Compruebo cuenta bancaria
    
    If Text1(11).Text = "S" Then
        SQL = Text1(13).Text & Text1(14).Text & Text1(16).Text
        If SQL = "" Then
            Text1(15).Text = ""
        Else
            If Len(SQL) <> 18 Then
                MsgBox "Longitud cuenta bancaria incorrecta", vbExclamation
                Exit Function
            Else
                RC2 = SQL
            
                SQL = CodigoDeControl(SQL)
                If SQL <> Text1(15).Text Then
                    
                    SQL = "Código de control para la cuenta bancaria: " & SQL & vbCrLf
                    SQL = SQL & "Código de control introducido: " & Text1(15).Text & vbCrLf & vbCrLf
                    SQL = SQL & "Continuar?"
                    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
                    
                'Noviembre 2013
                'Compruebo EL IBAN
                'Meto el CC
                RC2 = Mid(RC2, 1, 8) & Me.Text1(15).Text & Mid(RC2, 9)
                SQL = ""
                If Me.Text1(29).Text <> "" Then SQL = Mid(Text1(29).Text, 1, 2)
                    
                If DevuelveIBAN2(SQL, RC2, RC2) Then
                    If Me.Text1(29).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(29).Text = RC2
                    Else
                        If Mid(Text1(29).Text, 3) <> RC2 Then
                            RC2 = "Calculado : " & SQL & RC2
                            RC2 = "Introducido: " & Me.Text1(29).Text & vbCrLf & RC2 & vbCrLf
                            RC2 = "Error en codigo IBAN" & vbCrLf & RC2 & "Continuar?"
                            If MsgBox(RC2, vbQuestion + vbYesNo) = vbNo Then Exit Function
                        End If
                    End If
                End If
                    
                    
                    
                    
            End If
        End If
    End If
    DatosOk = True
End Function




Private Sub PierdeFocoCodigoCuenta()
Dim B As Boolean
If vModo = 3 Then Exit Sub  'Búsqueda


If vModo = 1 Then Text1(0).Text = Trim(Text1(0).Text)

'Si no compruebo que es un campo numerico
If Not IsNumeric(Text1(0).Text) Then
    MsgBox "El código de cuenta es un campo numérico", vbExclamation
    Exit Sub
End If

'Vemos si a puesto el punto para rellenar
Text1(0).Text = RellenaCodigoCuenta(Text1(0).Text)

If Len(Text1(0).Text) > vEmpresa.DigitosUltimoNivel Then
    MsgBox "El número máximo de dígitos para las cuentas es de " & vEmpresa.DigitosUltimoNivel & _
        vbCrLf & "La cuenta que ha puesto tiene " & Len(Text1(0).Text), vbExclamation
    Exit Sub
End If

'Comprobamos que ya existe la cuenta, solo en nueva
If vModo = 1 Then
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(0).Text, "T")
    If SQL <> "" Then
        MsgBox "La cuenta: " & Text1(0).Text & " ya esta asignada." & vbCrLf & "      .-" & SQL, vbExclamation
        Text1(0).SetFocus
        Exit Sub
    End If
End If
'Ponemos , si es de ultimo nivel habilitados los campos

B = EsCuentaUltimoNivel(Text1(0).Text)
Frame1.visible = B
Frame1.Enabled = True
chkUltimo.Value = Abs(CInt(B))
'Check2.Value = 0
If Not B Then
    'Si no es ultimo nivel
    Check2.Enabled = Len(Text1(0).Text) = 3
    PonerFrameGranEmpresa
Else
    'Ultimo nivel
    If vModo = 1 Then
        'Añadir cuenta
        SQL = DevuelveDesdeBD("model347", "cuentas", "codmacta", Mid(Text1(0).Text, 1, 3), "T")
        If SQL = "1" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    End If
End If

End Sub



Private Sub EnablarText(Si As Boolean)
Dim T As TextBox
    For Each T In Text1
        T.Locked = Not Si
    Next
    Image1(0).Enabled = Si
    Image1(1).Enabled = Si
    Check1.Enabled = Si
    Me.Check2.Enabled = Si
    Me.txtRegularizacion.Enabled = Si
    Me.chkUltimo.Enabled = Si
    'Solo administradores puden bloquear cuenta
    Text1(23).Enabled = vUsu.Nivel <= 1
    imgppal(2).Enabled = vUsu.Nivel <= 1
    
End Sub

Private Sub PonerDatosDeOtraCuenta(QueEmpresa_ As String)
Dim C As String
    C = Text1(0).Text
    Text1(0).visible = False
    CodCta = SQL
    PonerCampos QueEmpresa_
    lblIndicador.Caption = "Insertar"
    If QueEmpresa_ = "" Then
        Text1(0).Text = C
    Else
        If C <> "" Then Text1(0).Text = C
    End If
    Text1(0).visible = True
    CodCta = ""
End Sub

Private Sub txtRegularizacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtRegularizacion_LostFocus()
    If vModo = 3 Or vModo = 0 Then Exit Sub 'Busqueda avanzada o ver solo
    
    
    If txtRegularizacion.Text = "" Then Exit Sub
    
    'Si no compruebo que es un campo numerico
    If Not IsNumeric(txtRegularizacion.Text) Then
        MsgBox "El código de cuenta es un campo numérico", vbExclamation
        txtRegularizacion.Text = ""
        PonFoco txtRegularizacion
        Exit Sub
    End If
    
    'Vemos si a puesto el punto para rellenar
    txtRegularizacion.Text = RellenaCodigoCuenta(txtRegularizacion.Text)
    
    
    
    'Solo son validad cuentas del grupo 1
    If Mid(txtRegularizacion.Text, 1, 1) <> "1" Then
        MsgBox "La regularizacion será contra las cuentas del grupo 1", vbExclamation
        txtRegularizacion.Text = ""
        PonFoco txtRegularizacion
        Exit Sub
    End If
    
    
    
    If Len(Text1(0).Text) > vEmpresa.DigitosUltimoNivel Then
        MsgBox "El número máximo de dígitos para las cuentas es de " & vEmpresa.DigitosUltimoNivel & _
            vbCrLf & "La cuenta que ha puesto tiene " & Len(Text1(0).Text), vbExclamation
        txtRegularizacion.Text = ""
        PonFoco txtRegularizacion
        Exit Sub
    End If
    
    
    
    
    
End Sub





Private Sub ComprobarNifTodasContas()
    Set miRsAux = New ADODB.Recordset
    DoEvents
    cargaempresas
    lblIndicador.Caption = "Modificar"
    Set miRsAux = Nothing
End Sub


Private Sub cargaempresas()
Dim Mensa As String
Dim Prohibidas As Boolean
Dim C As String
On Error GoTo Ecargaempresas

    
    
    SQL = "Select count(*) from Usuarios.usuarioempresa WHERE codusu = " & (vUsu.Codigo Mod 1000)
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Prohibidas = False
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then Prohibidas = True
    End If
    miRsAux.Close

    
    SQL = "Select * from Usuarios.Empresas order by codempre"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & miRsAux!codempre & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Mensa = ""
    Do
        kCampo = InStr(1, SQL, "|")
        If kCampo > 0 Then
                C = Mid(SQL, 1, kCampo - 1)
                SQL = Mid(SQL, kCampo + 1)
                
                NumRegElim = Val(C)
                C = "conta" & C
                lblIndicador.Caption = "Comprobando NIF: " & C
                lblIndicador.Refresh
                C = "Select codmacta,nommacta FROM " & C & ".cuentas where apudirec='S'"
                If NumRegElim = vEmpresa.codempre Then
                    'Es esta empresa.
                    'Si es modificar añadire el codmacta <> de esta cuenta
                    If vModo = 2 Then C = C & " AND codmacta <> '" & Text1(0).Text & "'"
                End If
                C = C & " AND nifdatos ='" & Text1(7).Text & "'"
                miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                C = "Conta: " & NumRegElim & vbCrLf
                kCampo = 0
                While Not miRsAux.EOF
                    kCampo = 1
                    C = C & "    " & miRsAux!codmacta & " - " & miRsAux!nommacta & vbCrLf
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                If kCampo > 0 Then
                    Mensa = Mensa & C & vbCrLf
                Else
                    kCampo = 1
                End If
         End If
    Loop Until kCampo = 0
    
    
    If Mensa <> "" Then
        If Prohibidas Then
            Mensa = "YA existe el NIF en la contabilidad"
        Else
            Mensa = "El NIF aparece en la contabilidad." & vbCrLf & vbCrLf & Mensa
        End If
        MsgBox Mensa, vbExclamation
    End If
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
   
End Sub



Private Sub CargarComboPais()
Dim Aux As String
    cboPais.Clear
    SQL = "Alemania DE|Austria AT|Bélgica BE|Bulgaria BG|Chipre CY|"
    SQL = SQL & "Chequia CZ|Dinamarca DK|Estonia EE|Finlandia FI|Francia FR|"
    SQL = SQL & "Grecia EL|Gran Bretaña GB|Holanda NL|Hungría HU|Italia IT|"
    SQL = SQL & "Irlanda IE|Lituania LT|Luxemburgo LU|Letonia LV|Malta MT|Polonia PL|"
    SQL = SQL & "Portugal PT|Rumania RO|Suecia SE|Eslovenia SI|Eslovaquia SK|"
    
    Do
        kCampo = InStr(1, SQL, "|")
        If kCampo = 0 Then
            SQL = ""
        Else
            Aux = Mid(SQL, 1, kCampo - 1)
            SQL = Mid(SQL, kCampo + 1)
            cboPais.AddItem Right(Aux, 2) & " " & Mid(Aux, 1, Len(Aux) - 3)
            If UCase(Mid(Aux, 1, 4)) = "DINA" Then cboPais.AddItem "ESPAÑA"
        End If
    Loop Until SQL = ""
End Sub



Private Function HayMasDeUnaEmpresa() As Boolean

    HayMasDeUnaEmpresa = False
    SQL = " not codempre in (select codempre from usuarios.usuarioempresa where codusu=" & vUsu.Codigo Mod 1000 & ") and 1"
    SQL = DevuelveDesdeBD("count(*)", "usuarios.empresas", SQL, "1", "N")
    If SQL <> "" Then
        If Val(SQL) > 1 Then HayMasDeUnaEmpresa = True
    End If

End Function
