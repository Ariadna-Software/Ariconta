VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmparametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de la contabilidad"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10380
   Icon            =   "frmparametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   10380
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5160
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4515
      Left            =   120
      TabIndex        =   46
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7964
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos varios"
      TabPicture(0)   =   "frmparametros.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgFec(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgFec(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgFec(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(27)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(16)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(17)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(31)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Clientes - Proveedores "
      TabPicture(1)   =   "frmparametros.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "I.V.A./Norma 43/Enlaces"
      TabPicture(2)   =   "frmparametros.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame11"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame12"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame13"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Internet"
      TabPicture(3)   =   "frmparametros.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame13 
         Caption         =   "Autofactura"
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
         Height          =   900
         Left            =   6960
         TabIndex        =   117
         Top             =   2040
         Width           =   3135
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   33
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   111
            Tag             =   "C|T|S|||parametros|LetraSerieAutofactura|||"
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label11 
            Caption         =   "Letra de serie"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Automocion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   6960
         TabIndex        =   110
         Top             =   3120
         Width           =   3135
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   32
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   112
            Tag             =   "C|T|S|||parametros|automocion|||"
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label10 
            Caption         =   "Grupo exlcusion P y G"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   31
         Left            =   -71880
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha fin|F|S|||parametros|fechaActiva|dd/mm/yyyy||"
         Text            =   "1/2/3"
         Top             =   840
         Width           =   1275
      End
      Begin VB.Frame Frame11 
         Caption         =   "Enlaces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   6960
         TabIndex        =   107
         Top             =   480
         Width           =   3135
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   30
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   28
            Tag             =   "ODBC gestion|T|S|||parametros|enlaza_cta|||"
            Top             =   480
            Width           =   1485
         End
         Begin VB.Label Label9 
            Caption         =   "ODBC gestion"
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Parámetros importación datos bancarios"
         Height          =   1395
         Left            =   240
         TabIndex        =   98
         Top             =   480
         Width           =   6615
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   24
            Left            =   2280
            TabIndex        =   100
            Text            =   "Text2"
            Top             =   900
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   24
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   27
            Tag             =   "Diario norma 43|N|N|0|100|parametros|diario43|000||"
            Text            =   "1"
            Top             =   900
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   2280
            TabIndex        =   99
            Text            =   "Text2"
            Top             =   420
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   23
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   26
            Tag             =   "Concepto norma 43|N|N|0||parametros|conce43|||"
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   102
            Top             =   960
            Width           =   405
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   2
            Left            =   1020
            Picture         =   "frmparametros.frx":037A
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   4
            Left            =   1020
            Picture         =   "frmparametros.frx":6BCC
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   101
            Top             =   480
            Width           =   765
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1755
         Left            =   -74700
         TabIndex        =   93
         Top             =   2400
         Width           =   9555
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   27
            Left            =   2340
            MaxLength       =   100
            TabIndex        =   42
            Tag             =   "Web|T|S|||parametros|webversion|||"
            Text            =   "3"
            Top             =   1320
            Width           =   6060
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   26
            Left            =   2340
            MaxLength       =   100
            TabIndex        =   41
            Tag             =   "M|T|S|||parametros|mailsoporte|||"
            Text            =   "3"
            Top             =   780
            Width           =   6060
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   25
            Left            =   2340
            MaxLength       =   100
            TabIndex        =   40
            Tag             =   "W|T|S|||parametros|websoporte|||"
            Text            =   "3"
            Top             =   300
            Width           =   6060
         End
         Begin VB.Label Label8 
            Caption         =   "Soporte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   97
            Top             =   0
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Web check version"
            Height          =   195
            Index           =   16
            Left            =   300
            TabIndex        =   96
            Top             =   1380
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Mail soporte"
            Height          =   195
            Index           =   12
            Left            =   300
            TabIndex        =   95
            Top             =   840
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Web de soporte"
            Height          =   195
            Index           =   8
            Left            =   300
            TabIndex        =   94
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1815
         Left            =   -74700
         TabIndex        =   87
         Top             =   480
         Width           =   9555
         Begin VB.CheckBox Check1 
            Caption         =   "Enviar desde el Outlook"
            Height          =   225
            Index           =   14
            Left            =   6720
            TabIndex        =   116
            Tag             =   "Outlook|N|N|||parametros|EnvioDesdeOutlook|||"
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   19
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   36
            Tag             =   "Direccion e-mail|T|S|||parametros|diremail|||"
            Text            =   "3"
            Top             =   420
            Width           =   4860
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   20
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   37
            Tag             =   "Servidor SMTP|T|S|||parametros|smtpHost|||"
            Text            =   "3"
            Top             =   900
            Width           =   4860
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   21
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   38
            Tag             =   "Usuario SMTP|T|S|||parametros|smtpUser|||"
            Text            =   "3"
            Top             =   1440
            Width           =   4260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   7080
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   39
            Tag             =   "Password SMTP|T|S|||parametros|smtpPass|||"
            Text            =   "3"
            Top             =   1440
            Width           =   2220
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   20
            Left            =   300
            TabIndex        =   92
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            Height          =   195
            Index           =   21
            Left            =   300
            TabIndex        =   91
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   22
            Left            =   300
            TabIndex        =   90
            Top             =   1500
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   23
            Left            =   6300
            TabIndex        =   89
            Top             =   1500
            Width           =   840
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   88
            Top             =   0
            Width           =   1320
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -70440
         MaxLength       =   8
         TabIndex        =   86
         Text            =   "1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -71160
         MaxLength       =   8
         TabIndex        =   85
         Text            =   "1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   -71880
         MaxLength       =   8
         TabIndex        =   84
         Text            =   "1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   -72600
         MaxLength       =   8
         TabIndex        =   83
         Text            =   "1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74820
         TabIndex        =   47
         Top             =   2400
         Width           =   9615
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmparametros.frx":D41E
            Left            =   6780
            List            =   "frmparametros.frx":D428
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Tag             =   "Documento proveedores|T|S|||parametros|codinume|||"
            Top             =   1200
            Width           =   2235
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   5520
            TabIndex        =   50
            Text            =   "Text2"
            Top             =   540
            Width           =   3675
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   840
            TabIndex        =   49
            Text            =   "Text2"
            Top             =   1260
            Width           =   3675
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   840
            TabIndex        =   48
            Text            =   "Text2"
            Top             =   540
            Width           =   3675
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   240
            MaxLength       =   8
            TabIndex        =   22
            Tag             =   "Diario proveedores|N|N|0|100|parametros|numdiapr|000||"
            Text            =   "3"
            Top             =   540
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   4920
            MaxLength       =   8
            TabIndex        =   23
            Tag             =   "Conceptos facturas proveedores|N|S|0|1000|parametros|concefpr|000||"
            Text            =   "2"
            Top             =   540
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   240
            MaxLength       =   8
            TabIndex        =   24
            Tag             =   "Conceptos abonos proveedores|N|S|0|1000|parametros|conceapr|000||"
            Text            =   "1"
            Top             =   1260
            Width           =   525
         End
         Begin VB.Label Label7 
            Caption         =   "Proveedores"
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
            Height          =   255
            Left            =   180
            TabIndex        =   74
            Top             =   0
            Width           =   1170
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   1
            Left            =   1560
            Picture         =   "frmparametros.frx":D44D
            Top             =   1020
            Width           =   240
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   0
            Left            =   6240
            Picture         =   "frmparametros.frx":13C9F
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   1
            Left            =   780
            Picture         =   "frmparametros.frx":1A4F1
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Documento proveedores"
            Height          =   255
            Index           =   19
            Left            =   4920
            TabIndex        =   64
            Top             =   1260
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   53
            Top             =   300
            Width           =   405
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto facturas"
            Height          =   195
            Index           =   6
            Left            =   4920
            TabIndex        =   52
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto abonos"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   51
            Top             =   1020
            Width           =   1260
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2475
         Left            =   240
         TabIndex        =   76
         Top             =   1920
         Width           =   6615
         Begin VB.CheckBox Check1 
            Caption         =   "340. Tickets en serie de factura"
            Height          =   225
            Index           =   13
            Left            =   3240
            TabIndex        =   115
            Tag             =   "TicketsEn340LetraSerie|N|N|||parametros|TicketsEn340LetraSerie|||"
            Top             =   840
            Width           =   2775
         End
         Begin VB.CheckBox Check1 
            Caption         =   "349. Presentacion mensual"
            Height          =   225
            Index           =   12
            Left            =   240
            TabIndex        =   114
            Tag             =   "Periodo mensual|N|N|||parametros|Presentacion349Mensual|||"
            Top             =   840
            Width           =   2775
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Periodo MENSUAL"
            Height          =   225
            Index           =   5
            Left            =   240
            TabIndex        =   29
            Tag             =   "Periodo mensual|N|N|||parametros|periodos|||"
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   5160
            MaxLength       =   2
            TabIndex        =   32
            Tag             =   "Ultimo periodo liquidación I.V.A.|N|S|0|100|parametros|perfactu|||"
            Text            =   "2"
            Top             =   1320
            Width           =   360
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   31
            Tag             =   "Ultimo año liquidación I.V.A.|N|S|0|9999|parametros|anofactu|||"
            Text            =   "1999"
            Top             =   1320
            Width           =   660
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   3480
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "Limite 347|N|S|0||parametros|limimpcl|0.00||"
            Text            =   "3"
            Top             =   2040
            Width           =   1020
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Modificar apuntes factura"
            Height          =   225
            Index           =   10
            Left            =   3240
            TabIndex        =   30
            Tag             =   "Mod apuntes factura|N|N|||parametros|modhcofa|||"
            Top             =   360
            Width           =   2190
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Año fiscal"
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   33
            Top             =   1920
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Año natural"
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   34
            Top             =   1920
            Width           =   1035
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Año natural"
            Height          =   225
            Index           =   4
            Left            =   2640
            TabIndex        =   77
            Tag             =   "Año natural|N|N|||parametros|tinumfac|||"
            Top             =   1560
            Visible         =   0   'False
            Width           =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Limite 347"
            Height          =   195
            Index           =   13
            Left            =   3480
            TabIndex        =   79
            Top             =   1800
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Último período liquidación"
            Height          =   195
            Index           =   15
            Left            =   3120
            TabIndex        =   82
            Top             =   1320
            Width           =   1860
         End
         Begin VB.Label Label1 
            Caption         =   "Último año liquidación"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   81
            Top             =   1320
            Width           =   1740
         End
         Begin VB.Label Label2 
            Caption         =   "I.V.A."
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
            Left            =   360
            TabIndex        =   80
            Top             =   0
            Width           =   630
         End
         Begin VB.Label Label5 
            Caption         =   "Contadores:"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   1920
            Width           =   855
         End
         Begin VB.Shape Shape1 
            Height          =   555
            Left            =   1200
            Top             =   1800
            Width           =   2175
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1035
         Left            =   -74760
         TabIndex        =   70
         Top             =   1200
         Width           =   6375
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   18
            Left            =   240
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1740
            TabIndex        =   72
            Text            =   "Text4"
            Top             =   360
            Width           =   4455
         End
         Begin VB.Image imgCta 
            Height          =   240
            Left            =   1440
            Picture         =   "frmparametros.frx":20D43
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta pérdidas y ganancias"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3795
         Left            =   -68280
         TabIndex        =   69
         Top             =   480
         Width           =   3315
         Begin VB.CheckBox Check1 
            Caption         =   "Gran empresa (8 y 9)"
            Height          =   225
            Index           =   15
            Left            =   240
            TabIndex        =   10
            Tag             =   "c|N|N|||parametros|granempresa|||"
            Top             =   3360
            Width           =   2235
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Agencia de viajes"
            Height          =   225
            Index           =   8
            Left            =   240
            TabIndex        =   9
            Tag             =   "c|N|N|||parametros|agenciaviajes|||"
            Top             =   2850
            Width           =   2235
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Constructoras-Promotoras"
            Height          =   225
            Index           =   9
            Left            =   240
            TabIndex        =   4
            Tag             =   "Constructoras|N|N|||parametros|constructoras|||"
            Top             =   300
            Width           =   2235
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Contabilizar factura automáticamente"
            Height          =   225
            Index           =   6
            Left            =   240
            TabIndex        =   8
            Tag             =   "Asiento automatico|N|N|||parametros|ContabilizaFact|||"
            Top             =   2340
            Width           =   3030
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Emite diario al modificar"
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Tag             =   "Emite diarioal modificat|N|N|||parametros|listahco|||"
            Top             =   1320
            Width           =   2070
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Emite diario al actualizar"
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Tag             =   "Emite diario|N|N|||parametros|emitedia|||"
            Top             =   810
            Width           =   2235
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Actualizar asiento automaticamente"
            Height          =   225
            Index           =   11
            Left            =   240
            TabIndex        =   7
            Tag             =   "Asiento automatico|N|N|||parametros|asienactauto|||"
            Top             =   1830
            Width           =   2850
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Analítica"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   65
         Top             =   2280
         Width           =   6375
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   735
            Left            =   3480
            TabIndex        =   106
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   29
            Left            =   5400
            MaxLength       =   1
            TabIndex        =   104
            Tag             =   "Grupo 2|T|S|||parametros|Subgrupo2|||"
            Text            =   "Text1"
            Top             =   1440
            Width           =   600
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   28
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   16
            Tag             =   "Grupo 1|T|S|||parametros|Subgrupo1|||"
            Text            =   "Text1"
            Top             =   1440
            Width           =   600
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Grabar C.C. en la contabilización de facturas"
            Height          =   225
            Index           =   2
            Left            =   2400
            TabIndex        =   12
            Tag             =   "Autocoste|N|S|||paramertros|CCenFacturas|||"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   5820
            MaxLength       =   1
            TabIndex        =   15
            Tag             =   "Grupo|T|S|||parametros|grupoord|||"
            Text            =   "2"
            Top             =   810
            Width           =   360
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   3600
            MaxLength       =   1
            TabIndex        =   14
            Tag             =   "Grupo|T|S|||parametros|grupovta|||"
            Text            =   "1"
            Top             =   810
            Width           =   360
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   13
            Tag             =   "Grupo|T|S|||parametros|grupogto|||"
            Text            =   "Text1"
            Top             =   840
            Width           =   360
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Contabilidad analítica"
            Height          =   225
            Index           =   0
            Left            =   420
            TabIndex        =   11
            Tag             =   "Autocoste|N|N|||paramertros|autocoste|||"
            Top             =   360
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "Subgrupo a 3 digitos"
            Height          =   255
            Index           =   26
            Left            =   3600
            TabIndex        =   105
            Top             =   1440
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Subgrupo a 3 digitos"
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   103
            Top             =   1440
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Otro Grupo Analitica"
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   68
            Top             =   825
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Grupo de ventas"
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   67
            Top             =   825
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Grupo de gastos"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   66
            Top             =   825
            Width           =   1365
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -74760
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Fecha inicio|F|N|||parametros|fechaini|dd/mm/yyyy|S|"
         Text            =   "1/2/3"
         Top             =   825
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -73245
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha fin|F|N|||parametros|fechafin|dd/mm/yyyy||"
         Text            =   "1/2/3"
         Top             =   825
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74820
         TabIndex        =   54
         Top             =   420
         Width           =   9615
         Begin VB.CheckBox Check1 
            Caption         =   "Abonos negativos"
            Height          =   225
            Index           =   7
            Left            =   7560
            TabIndex        =   21
            Tag             =   "Abonos negativos|N|N|||parametros|abononeg|||"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmparametros.frx":27595
            Left            =   4440
            List            =   "frmparametros.frx":275A2
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "Ampliación clientes|T|S|||parametros|nctafact|||"
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   240
            MaxLength       =   8
            TabIndex        =   19
            Tag             =   "Conceptos abonos clientes|N|S|0|1000|parametros|conceacl|000||"
            Text            =   "3"
            Top             =   1320
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   4440
            MaxLength       =   8
            TabIndex        =   18
            Tag             =   "Conceptos facturas clientes|N|S|0|1000|parametros|concefcl|000||"
            Text            =   "2"
            Top             =   540
            Width           =   540
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   240
            MaxLength       =   8
            TabIndex        =   17
            Tag             =   "Diario clientes|N|N|0|100|parametros|numdiacl|000||"
            Text            =   "1"
            Top             =   540
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   840
            TabIndex        =   57
            Text            =   "Text2"
            Top             =   540
            Width           =   3435
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   5040
            TabIndex        =   56
            Text            =   "Text2"
            Top             =   540
            Width           =   3735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   900
            TabIndex        =   55
            Text            =   "Text2"
            Top             =   1320
            Width           =   3315
         End
         Begin VB.Label Label6 
            Caption         =   "Clientes"
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
            Left            =   240
            TabIndex        =   73
            Top             =   0
            Width           =   810
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   3
            Left            =   1560
            Picture         =   "frmparametros.frx":275CB
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   2
            Left            =   5760
            Picture         =   "frmparametros.frx":2DE1D
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   0
            Left            =   720
            Picture         =   "frmparametros.frx":3466F
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Ampliación clientes / Proveedores"
            Height          =   255
            Index           =   18
            Left            =   4440
            TabIndex        =   63
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto abonos"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   60
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto facturas"
            Height          =   195
            Index           =   10
            Left            =   4440
            TabIndex        =   59
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   58
            Top             =   300
            Width           =   405
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha activa"
         Height          =   255
         Index           =   27
         Left            =   -71880
         TabIndex        =   109
         Top             =   540
         Width           =   1020
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   -70800
         Picture         =   "frmparametros.frx":3AEC1
         Top             =   547
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   -72540
         Picture         =   "frmparametros.frx":3AF4C
         Top             =   540
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   -73860
         Picture         =   "frmparametros.frx":3AFD7
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inicio"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   62
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha fin"
         Height          =   255
         Index           =   1
         Left            =   -73245
         TabIndex        =   61
         Top             =   540
         Width           =   1020
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1020
      Top             =   4620
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8040
      TabIndex        =   43
      Top             =   5160
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
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
      TabIndex        =   45
      Top             =   5280
      Width           =   2310
   End
End
Attribute VB_Name = "frmparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCo As frmConceptos
Attribute frmCo.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Dim RS As ADODB.Recordset
Dim Modo As Byte
Dim I As Integer



Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim ModificaClaves As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1


    Select Case Modo
    Case 0
        'Preparao para modificar
        PonerModo 2
        
    Case 1
        
        If DatosOk Then
            'Cambiamos el path
            'CambiaPath True
            If InsertarDesdeForm(Me) Then PonerModo 0
        End If
    
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                'CambiaPath True
                
                
                ModificaClaves = False
                If vUsu.Nivel = 0 Then
                    If vParam.fechaini <> Text1(0).Text Then
                        ModificaClaves = True
                        Cad = " fechaini = '" & Format(vParam.fechaini, FormatoFecha) & "'"
                    End If
                End If
                If ModificaClaves Then
                    If ModificaDesdeFormularioClaves(Me, Cad) Then
                        ReestableceVPARAM
                        PonerModo 0
                    End If
                Else
                    If ModificaDesdeFormulario(Me) Then PonerModo 0
                End If
'                CambiaPath False
            End If

    End Select
    
    'Si el modo es 0 significa k han insertado o modificado cosas
    If Modo = 0 Then _
        MsgBox "Para que los cambios tengan efecto debe reiniciar la aplicación.", vbExclamation
        
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
If Modo = 2 Then PonerCampos
PonerModo 0
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    Me.Top = 200
    Me.Left = 100
        ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4
        .Buttons(2).Image = 15
    End With
    
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = "Select * from parametros"
    adodc1.Refresh
    If adodc1.Recordset.EOF Then
        'No hay datos
        Limpiar Me
        PonerModo 1
        Else
        PonerCampos
        PonerModo 0
        'Campos que nos se tocaran los ponemos con colorcitos bonitos
        If vUsu.Nivel <> 0 Then
            Text1(0).BackColor = &H80000018
            Text1(1).BackColor = &H80000018
        End If
    End If
    Toolbar1.Buttons(1).Enabled = (vUsu.Nivel <= 1)
    cmdAceptar.Enabled = (vUsu.Nivel <= 1)
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    imgFec(1).Tag = vFecha
End Sub

Private Sub frmCo_DatoSeleccionado(CadenaSeleccion As String)
    imgConcep(1).Tag = CadenaSeleccion
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text1(18).Text = RecuperaValor(CadenaSeleccion, 1)
    Text4.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    imgDiario(1).Tag = CadenaSeleccion
End Sub



Private Sub imgConcep_Click(Index As Integer)
    
    imgConcep(1).Tag = ""
    Select Case Index
    Case 0, 1
        imgConcep(0).Tag = 6
    Case 2, 3
        imgConcep(0).Tag = 8
    Case 4
        imgConcep(0).Tag = 19
    End Select
    Set frmCo = New frmConceptos
    frmCo.DatosADevolverBusqueda = "0|1|"
    frmCo.Show vbModal
    Set frmCo = Nothing
    Index = CInt(imgConcep(0).Tag) + Index
    If imgConcep(1).Tag <> "" Then
        Text1(Index).Text = Format(RecuperaValor(imgConcep(1).Tag, 1), "000")
        Text2(Index).Text = RecuperaValor(imgConcep(1).Tag, 2)
    End If
End Sub

Private Sub imgcta_Click()
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 1
    frmCta.Show vbModal
End Sub

Private Sub imgDiario_Click(Index As Integer)
    imgDiario(1).Tag = ""
    If Index = 0 Then
        Index = 9
    Else
        If Index = 1 Then
            Index = 5
        Else
            Index = 24
        End If
    End If
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
    If imgDiario(1).Tag <> "" Then
        Text1(Index).Text = Format(RecuperaValor(imgDiario(1).Tag, 1), "000")
        Text2(Index).Text = RecuperaValor(imgDiario(1).Tag, 2)
    End If
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim F As Date
    'En los tag
    'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
    imgFec(0).Tag = Index
    F = Now
    imgFec(1).Tag = ""
    If Text1(Index).Text <> "" Then
        If IsDate(Text1(Index).Text) Then F = Text1(Index).Text
    End If
    Set frmC = New frmCal
    frmC.Fecha = F
    frmC.Show vbModal
    Set frmC = Nothing
    If imgFec(1).Tag <> "" Then
        If IsDate(imgFec(1).Tag) Then Text1(Index).Text = Format(CDate(imgFec(1).Tag), "dd/mm/yyyy")
    End If
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

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
    Dim Cad As String
    Dim SQL As String
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)

    If Index > 1 And Index <> 31 Then FormateaCampo Text1(Index)     'Formateamos el campo si tiene valor excepto las fechas
    
    'Si queremos hacer algo ..
    Select Case Index
    Case 0, 1, 31
        If Text1(Index).Text = "" Then Exit Sub
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta : " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Text1(Index).SetFocus
            Exit Sub
        End If
                        
    Case 2, 3, 4
        If Text1(Index).Text = "" Then Exit Sub
        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(Index).Text, "T")
        If SQL = "" Then
            MsgBox "La cuenta no existe: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Text1(Index).SetFocus
        End If
    Case 5, 9, 24
     ' Diarios
       If Not IsNumeric(Text1(Index).Text) Then Exit Sub
       SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(Index).Text)
       If SQL = "" Then
            SQL = "Codigo incorrecto"
            Text1(Index).Text = "-1"
        End If
       Text2(Index).Text = SQL
    Case 6, 7, 10, 11, 23
       'Conceptos
       If Not IsNumeric(Text1(Index).Text) Then Exit Sub
       SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(Index).Text)
       If SQL = "" Then
            SQL = "Codigo incorrecto"
            Text1(Index).Text = "-1"
        End If
       Text2(Index).Text = SQL
        '....
    Case 18
        Cad = Text1(18).Text
        If CuentaCorrectaUltimoNivel(Cad, SQL) Then
            Text1(18).Text = Cad
            Text4.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text1(18).Text = Cad
            Text4.Text = SQL
            If Modo > 2 Then Text1(18).SetFocus
        End If
    Case 28, 29
        Cad = Trim(Text1(Index).Text)
        If Cad <> "" Then
            If Not IsNumeric(Cad) Then
                MsgBox Cad & ": No es un campo numerico", vbExclamation
                Text1(Index).Text = ""
                Text1(Index).SetFocus
            End If
        End If
    Case 33
        Cad = DevuelveDesdeBD("nomregis", "contadores", "tiporegi", Text1(33).Text, "T")
        If Cad = "" Then
            MsgBox "No existe la letra de serie", vbExclamation
            Text1(33).Text = ""
        End If
    End Select
    '---
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim Valor As Boolean
    Modo = Kmodo
    Select Case Kmodo
    Case 0
        'Preparamos para ver los datos
        Valor = True
        Label3.Caption = ""

    Case 1
        'Preparamos para que pueda insertar
        Valor = False
        Label3.Caption = "INSERTAR"
        Label3.ForeColor = vbBlue

    Case 2
        Valor = False
        Label3.Caption = "MODIFICAR"
        Label3.ForeColor = vbRed

    End Select
    cmdAceptar.visible = Modo > 0
    cmdCancelar.visible = Modo > 0
    'Ponemos los valores
        For I = 0 To Text1.Count - 1
            Text1(I).Locked = Valor
        Next I
        Frame1.Enabled = Not Valor
        Frame4.Enabled = Not Valor
        Frame6.Enabled = Not Valor
        Frame3.Enabled = Not Valor
        Frame5.Enabled = Not Valor
        Frame2.Enabled = Not Valor
        Frame7.Enabled = Not Valor

        For I = 0 To imgDiario.Count - 1
            imgDiario(I).Enabled = Not Valor
        Next I
        For I = 0 To imgConcep.Count - 1
            imgConcep(I).Enabled = Not Valor
        Next I
        
        'Campos que solo estan habilitados para insercion
        If Not Valor Then
            Text1(0).Locked = (vUsu.Nivel >= 1)
            Text1(1).Locked = (vUsu.Nivel >= 1)
        End If
        For I = 0 To imgFec.Count - 1
            imgFec(I).Enabled = Not Text1(0).Locked
        Next I
End Sub

Private Sub PonerCampos()
    Dim Cam As String
    Dim Tabla As String
    Dim Cod As String
    
        If adodc1.Recordset.EOF Then Exit Sub
        If PonerCamposForma(Me, adodc1) Then
           'Correcto, ponemos los datos auxiliares
           '----------------------------------------
           ' Diarios
           Cam = "desdiari"
           Tabla = "tiposdiario"
           Cod = "numdiari"
           Text2(9).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(9).Text)
           Text2(5).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(5).Text)
           Text2(24).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(24).Text)
           
           'Conceptos
           Cam = "nomconce"
           Tabla = "conceptos"
           Cod = "codconce"
           Text2(10).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(10).Text)
           Text2(11).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(11).Text)
           Text2(6).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(6).Text)
           Text2(7).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(7).Text)
           Text2(23).Text = DevuelveDesdeBD(Cam, Tabla, Cod, Text1(23).Text)
           
           'Cuenta de pérdidas y ganancias
           Text4.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(18).Text, "T")
           
           'Año natural
           Option1(0).Value = (Check1(4).Value = 1)
           Option1(1).Value = Not Option1(0).Value
           
           'Cambiamos los path
'           CambiaPath False
           
        End If
End Sub
'
Private Function DatosOk() As Boolean
    Dim B As Boolean
    Dim J As Integer
    
    
    DatosOk = False
    
    'Si esta marcado la analitica, entonces debe tener valor grupo ventas y grupo gastos
    If Me.Check1(0).Value = 1 Then
        Text1(2).Text = Trim(Text1(2).Text)
        Text1(3).Text = Trim(Text1(3).Text)
        If Text1(2).Text = "" Or Text1(3).Text = "" Then
            MsgBox "Si selecciona la contabilidad analítica debe poner valor al grupo gastos y ventas.", vbExclamation
            Exit Function
        End If
    End If
    
        
    
    'NO puede marcar constructora y Agencia viajes a la vez
    If Check1(8).Value = 1 And Check1(9).Value Then
        MsgBox "No puede marcar Contructora y Agencia de viajes a la vez", vbExclamation
        Exit Function
    End If
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Si tiene puesta fecha Activa, esta no puede ser ni menor que fechaini ni mayor que fecha fin
    If Text1(31).Text <> "" Then
        If CDate(Text1(31).Text) < CDate(Text1(0).Text) Or CDate(Text1(31).Text) > DateAdd("yyyy", 1, CDate(Text1(1).Text)) Then
            MsgBox "Fecha activa debe estar entre incio ejercicio - fin ejercicio siguiente", vbExclamation
            B = False
            Exit Function
        End If
    End If
    
    
    'Año naural
    Check1(4).Value = Abs(Option1(0).Value)
    
    If Modo = 1 Then
        If CDate(Text1(0).Text) >= CDate(Text1(1).Text) Then
            MsgBox "Fecha inicio mayor o igual que la fecha final de ejercicio.", vbExclamation
            B = False
        End If
    End If
    
    
    'Comprobamos que si el periodo de liquidacion de IVA es mensual
    'Comprobaremos que el valor del periodo no excede de
    If Check1(5).Value = 1 Then
        I = 12  'Mensual
    Else
        I = 4   'Trimestral
    End If
    If Text1(15).Text <> "" Then
        J = CInt(Text1(15).Text)
        If J = 0 Then
            MsgBox "El periodo de liquidación no puede ser 0", vbExclamation
            Exit Function
        End If
        If J > I Then
            MsgBox "Periodo de liquidacion incorrecto", vbExclamation
            Exit Function
        End If
    End If
    
    
    
    If B Then
        J = Len(Text1(32).Text)
        If J > 0 Then
            If Not IsNumeric(Text1(32).Text) Then
                MsgBox "Campo debe ser numerico: subgrupo de excepcion de automocion", vbExclamation
                Exit Function
            End If
        End If
        
        If J > 0 And J < 3 Then
        
            MsgBox "El subgrupo de excepcion(exclusion) de automocion debe ser de 3 DIGITOS", vbExclamation
            Exit Function
        End If
    End If
        
    
    J = 33
    Text1(J).Text = Trim(Text1(J).Text)
    If Text1(J).Text <> "" Then
       If DevuelveDesdeBD("nomregis", "contadores", "tiporegi", Text1(J).Text, "T") = "" Then
            MsgBox "No existe en contadores la serie para las autofacturas", vbExclamation
            B = False
        End If
    End If
    
    
    
    DatosOk = B
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        'Modificar
         PonerModo 2
    Case 2
        'Salir
        Unload Me
    End Select
End Sub



Private Sub ReestableceVPARAM()
    Set vParam = Nothing
    Set vParam = New Cparametros
    vParam.Leer
End Sub
