VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame349 
      Height          =   5535
      Left            =   60
      TabIndex        =   572
      Top             =   0
      Visible         =   0   'False
      Width           =   4995
      Begin VB.CheckBox chk349 
         Caption         =   "Fichero AEAT"
         Height          =   255
         Left            =   120
         TabIndex        =   579
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   35
         Left            =   3360
         TabIndex        =   580
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   578
         Text            =   "Text4"
         Top             =   3600
         Width           =   495
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmListado.frx":030A
         Left            =   1680
         List            =   "frmListado.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   577
         Top             =   2880
         Width           =   1635
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   28
         Left            =   3720
         TabIndex        =   583
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   27
         Left            =   3720
         TabIndex        =   576
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   26
         Left            =   1140
         TabIndex        =   575
         Top             =   780
         Width           =   1095
      End
      Begin VB.CommandButton cmd349 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2400
         TabIndex        =   582
         Top             =   4920
         Width           =   1095
      End
      Begin VB.ListBox List5 
         Enabled         =   0   'False
         Height          =   1425
         ItemData        =   "frmListado.frx":0362
         Left            =   1500
         List            =   "frmListado.frx":0364
         TabIndex        =   584
         Top             =   1200
         Width           =   3315
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   21
         Left            =   2520
         TabIndex        =   757
         Top             =   4440
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   35
         Left            =   3000
         Picture         =   "frmListado.frx":0366
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie autofacturas"
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
         Index           =   108
         Left            =   120
         TabIndex        =   747
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
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
         Index           =   80
         Left            =   180
         TabIndex        =   587
         Top             =   2880
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   17
         Left            =   180
         TabIndex        =   586
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   79
         Left            =   180
         TabIndex        =   585
         Top             =   480
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   27
         Left            =   3360
         Picture         =   "frmListado.frx":03F1
         Top             =   780
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   26
         Left            =   780
         Picture         =   "frmListado.frx":047C
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   16
         Left            =   2760
         TabIndex        =   581
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Modelo 349  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   18
         Left            =   1560
         TabIndex        =   574
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Index           =   78
         Left            =   180
         TabIndex        =   573
         Top             =   1200
         Width           =   825
      End
      Begin VB.Image Image8 
         Height          =   240
         Left            =   1140
         Picture         =   "frmListado.frx":0507
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame frameBalance 
      Height          =   7395
      Left            =   120
      TabIndex        =   108
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CheckBox chkBalIncioEjercicio 
         Caption         =   "Balance a inicio ejercicio"
         Height          =   195
         Left            =   3960
         TabIndex        =   841
         Top             =   6000
         Width           =   2295
      End
      Begin VB.CheckBox chkResetea6y7 
         Caption         =   "Contemplar cuentas gastos y ventas desde ejercicio siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   752
         Top             =   6360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   5355
      End
      Begin VB.CheckBox chkQuitaCierre 
         Caption         =   "Antes del cierre"
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   135
         Top             =   5640
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkAgrupacionCtasBalance 
         Caption         =   "Agrupar cuentas en balance"
         Height          =   255
         Left            =   3960
         TabIndex        =   133
         Top             =   5280
         Width           =   2355
      End
      Begin VB.TextBox txtNpag 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4860
         TabIndex        =   118
         Text            =   "Text2"
         Top             =   3120
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   3120
         Width           =   2910
      End
      Begin VB.CheckBox chkQuitaCierre 
         Caption         =   "Antes de pérdidas y ganancias"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   134
         Top             =   5640
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   4740
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   4140
         TabIndex        =   114
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chkMovimientos 
         Caption         =   "Imprimir acumulados y movimientos del periodo"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   6000
         Value           =   1  'Checked
         Width           =   3675
      End
      Begin VB.CheckBox chkApertura 
         Caption         =   "Desglosar el saldo de apertura"
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   5280
         Width           =   3255
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   375
         Left            =   120
         TabIndex        =   146
         Top             =   6900
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5160
         TabIndex        =   109
         Top             =   6840
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   1035
         Left            =   120
         TabIndex        =   120
         Top             =   3480
         Width           =   5865
         Begin VB.CheckBox Check2 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   121
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox Check2 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   122
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox Check2 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   123
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox Check2 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   124
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox Check2 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   125
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox Check2 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   126
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox Check2 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   127
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox Check2 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   128
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox Check2 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   129
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox Check2 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   130
            Top             =   720
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmdBalance 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   137
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   110
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   111
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2820
         TabIndex        =   138
         Text            =   "Text5"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2820
         TabIndex        =   119
         Text            =   "Text5"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado.frx":0F09
         Left            =   1380
         List            =   "frmListado.frx":0F0B
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   0
         Left            =   2820
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   1
         ItemData        =   "frmListado.frx":0F0D
         Left            =   1380
         List            =   "frmListado.frx":0F0F
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   1
         Left            =   2820
         TabIndex        =   116
         Text            =   "Text1"
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pág:"
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
         Index           =   47
         Left            =   4140
         TabIndex        =   399
         Top             =   3180
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Título"
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
         Index           =   37
         Left            =   240
         TabIndex        =   323
         Top             =   3180
         Width           =   480
      End
      Begin VB.Label label7 
         Caption         =   "Remarcar nivel"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   7
         Left            =   5340
         Picture         =   "frmListado.frx":0F11
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   11
         Left            =   4140
         TabIndex        =   147
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   6
         Left            =   1320
         Picture         =   "frmListado.frx":0F9C
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   7
         Left            =   1320
         Picture         =   "frmListado.frx":199E
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   10
         Left            =   720
         TabIndex        =   145
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   11
         Left            =   720
         TabIndex        =   144
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance de sumas y saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   960
         TabIndex        =   143
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   10
         Left            =   240
         TabIndex        =   142
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   12
         Left            =   660
         TabIndex        =   141
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   13
         Left            =   660
         TabIndex        =   140
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   9
         Left            =   180
         TabIndex        =   139
         Top             =   1740
         Width           =   855
      End
   End
   Begin VB.Frame FrameBalPresupues 
      Height          =   5775
      Left            =   60
      TabIndex        =   231
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkQuitarApertura 
         Caption         =   "Quitar apertura"
         Height          =   255
         Left            =   4560
         TabIndex        =   247
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtMes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   248
         Text            =   "Text4"
         Top             =   2880
         Width           =   915
      End
      Begin MSComctlLib.ProgressBar pb4 
         Height          =   375
         Left            =   240
         TabIndex        =   259
         Top             =   5160
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdBalPre 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   249
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5040
         TabIndex        =   250
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2640
         TabIndex        =   252
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2640
         TabIndex        =   251
         Text            =   "Text5"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   244
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   243
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkPreMensual 
         Caption         =   "Mensual"
         Height          =   255
         Left            =   1440
         TabIndex        =   245
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkPreAct 
         Caption         =   "Ejercicio siguiente"
         Height          =   255
         Left            =   2640
         TabIndex        =   246
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Frame FrameNivelbalPresu 
         Height          =   1035
         Left            =   120
         TabIndex        =   232
         Top             =   3720
         Width           =   5865
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   242
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   241
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   240
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   239
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   238
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   237
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   236
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   235
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   234
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox ChkCtaPre 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   233
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nivel     "
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
            Index           =   25
            Left            =   120
            TabIndex        =   258
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Index           =   124
         Left            =   240
         TabIndex        =   846
         Top             =   2880
         Width           =   345
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
         Index           =   24
         Left            =   240
         TabIndex        =   257
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Balance presupuestario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   7
         Left            =   1440
         TabIndex        =   256
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   23
         Left            =   240
         TabIndex        =   255
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   600
         TabIndex        =   254
         Top             =   1680
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   13
         Left            =   1200
         Picture         =   "frmListado.frx":23A0
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   12
         Left            =   1200
         Picture         =   "frmListado.frx":2DA2
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   600
         TabIndex        =   253
         Top             =   1245
         Width           =   465
      End
   End
   Begin VB.Frame FrameBalancesper 
      Height          =   4395
      Left            =   60
      TabIndex        =   553
      Top             =   20
      Visible         =   0   'False
      Width           =   6075
      Begin VB.CheckBox chkApaisado 
         Caption         =   "Apaisado"
         Height          =   255
         Left            =   2880
         TabIndex        =   828
         Top             =   2520
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame FrameTapa2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   600
         TabIndex        =   570
         Top             =   2760
         Width           =   3795
      End
      Begin VB.CheckBox chkBalPerCompa 
         Caption         =   "Comparativo"
         Height          =   255
         Left            =   720
         TabIndex        =   568
         Top             =   2520
         Width           =   1515
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   25
         Left            =   180
         TabIndex        =   562
         Top             =   3900
         Width           =   1455
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   25
         Left            =   4800
         TabIndex        =   561
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdBalances 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   560
         Top             =   3780
         Width           =   975
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   17
         ItemData        =   "frmListado.frx":37A4
         Left            =   1440
         List            =   "frmListado.frx":37A6
         Style           =   2  'Dropdown List
         TabIndex        =   559
         Top             =   2940
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   17
         Left            =   2880
         TabIndex        =   558
         Text            =   "Text1"
         Top             =   2940
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   16
         ItemData        =   "frmListado.frx":37A8
         Left            =   1440
         List            =   "frmListado.frx":37AA
         Style           =   2  'Dropdown List
         TabIndex        =   557
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   16
         Left            =   2880
         TabIndex        =   556
         Text            =   "Text1"
         Top             =   1980
         Width           =   855
      End
      Begin VB.TextBox txtNumBal 
         Height          =   315
         Index           =   0
         Left            =   780
         TabIndex        =   555
         Text            =   "Text1"
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox TextDescBalance 
         BackColor       =   &H80000018&
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
         Index           =   0
         Left            =   1800
         TabIndex        =   554
         Text            =   "Text1"
         Top             =   1140
         Width           =   4035
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   54
         Left            =   720
         TabIndex        =   569
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   25
         Left            =   1500
         Picture         =   "frmListado.frx":37AC
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   77
         Left            =   180
         TabIndex        =   567
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label Label17 
         Caption         =   "Balances configurables"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   566
         Top             =   300
         Width           =   4875
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   55
         Left            =   720
         TabIndex        =   565
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   76
         Left            =   240
         TabIndex        =   564
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
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
         Index           =   75
         Left            =   180
         TabIndex        =   563
         Top             =   780
         Width           =   660
      End
      Begin VB.Image ImgNumBal 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "frmListado.frx":3837
         Top             =   1140
         Width           =   240
      End
   End
   Begin VB.Frame FrameListFactP 
      Height          =   7095
      Left            =   120
      TabIndex        =   659
      Top             =   0
      Visible         =   0   'False
      Width           =   5235
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   847
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtPag2 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   844
         Text            =   "Text1"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Frame FrameFactCons 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         TabIndex        =   741
         Top             =   600
         Width           =   5055
         Begin VB.ListBox List9 
            Enabled         =   0   'False
            Height          =   1620
            ItemData        =   "frmListado.frx":4239
            Left            =   1440
            List            =   "frmListado.frx":423B
            TabIndex        =   742
            Top             =   360
            Width           =   3375
         End
         Begin VB.Image Image12 
            Height          =   240
            Left            =   1080
            Picture         =   "frmListado.frx":423D
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Empresas"
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
            Index           =   106
            Left            =   120
            TabIndex        =   743
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.CheckBox ChkListFac 
         Caption         =   "Mostrar tipo Iva"
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   812
         Top             =   6000
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkMostrarRetencion 
         Caption         =   "Mostrar retención"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   753
         Top             =   6480
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   32
         Left            =   3840
         TabIndex        =   671
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtNpag2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1020
         TabIndex        =   670
         Text            =   "Text2"
         Top             =   4200
         Width           =   915
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   692
         Top             =   5280
         Width           =   4875
         Begin VB.OptionButton optMostrarFecha 
            Caption         =   "F. liquidacion"
            Height          =   255
            Index           =   2
            Left            =   3420
            TabIndex        =   696
            Top             =   240
            Width           =   1395
         End
         Begin VB.OptionButton optMostrarFecha 
            Caption         =   "Mostrar fecha"
            Height          =   255
            Index           =   1
            Left            =   1980
            TabIndex        =   694
            Top             =   240
            Width           =   1515
         End
         Begin VB.OptionButton optMostrarFecha 
            Caption         =   "Mostrar Nº factura"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   693
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar"
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
            Index           =   96
            Left            =   240
            TabIndex        =   704
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   691
         Top             =   3240
         Width           =   2955
         Begin VB.OptionButton optSelFech 
            Caption         =   "RECEPCION"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   666
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optSelFech 
            Caption         =   "LIQUIDACION"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   667
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.OptionButton optListFacP 
         Caption         =   "Fecha recepcion"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   674
         Top             =   4920
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optListFacP 
         Caption         =   "Nº Registro"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   672
         Top             =   4920
         Width           =   1335
      End
      Begin VB.OptionButton optListFacP 
         Caption         =   "Fecha emisión"
         Height          =   255
         Index           =   1
         Left            =   1740
         TabIndex        =   673
         Top             =   4920
         Width           =   1395
      End
      Begin VB.TextBox txtNumFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3960
         TabIndex        =   661
         Text            =   "Text1"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtNumFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   660
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   3960
         TabIndex        =   680
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdFacProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   679
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   21
         Left            =   1140
         TabIndex        =   665
         Top             =   2220
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   29
         Left            =   1200
         TabIndex        =   668
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   20
         Left            =   1140
         TabIndex        =   664
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   30
         Left            =   3840
         TabIndex        =   669
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2280
         TabIndex        =   678
         Text            =   "Text5"
         Top             =   2220
         Width           =   2715
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2280
         TabIndex        =   676
         Text            =   "Text5"
         Top             =   1800
         Width           =   2715
      End
      Begin VB.CheckBox ChkListFac 
         Caption         =   "Agrupar por cuenta"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   677
         Top             =   6000
         Width           =   2115
      End
      Begin VB.CheckBox ChkListFac 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   675
         Top             =   6000
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   663
         Text            =   "Text4"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   662
         Text            =   "Text4"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmListado.frx":4C3F
         ToolTipText     =   "Filtrar NIF"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N.I.F.                                   Tipo IVA"
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
         Index           =   123
         Left            =   360
         TabIndex        =   845
         Top             =   2760
         Width           =   2685
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   64
         Left            =   3360
         TabIndex        =   756
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   110
         Left            =   300
         TabIndex        =   755
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   63
         Left            =   1440
         TabIndex        =   754
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha recepción"
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
         Index           =   90
         Left            =   360
         TabIndex        =   681
         Top             =   3250
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   100
         Left            =   2220
         TabIndex        =   708
         Top             =   4260
         Width           =   1200
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   32
         Left            =   3540
         Picture         =   "frmListado.frx":5641
         Top             =   4260
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pág:"
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
         Index           =   95
         Left            =   360
         TabIndex        =   703
         Top             =   4200
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ordenación"
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
         Index           =   94
         Left            =   360
         TabIndex        =   697
         Top             =   4680
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   61
         Left            =   1440
         TabIndex        =   690
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   60
         Left            =   3360
         TabIndex        =   689
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registro"
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
         Index           =   93
         Left            =   300
         TabIndex        =   688
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   57
         Left            =   300
         TabIndex        =   687
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   360
         TabIndex        =   686
         Top             =   3645
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   56
         Left            =   300
         TabIndex        =   685
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   19
         Left            =   3000
         TabIndex        =   684
         Top             =   3645
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   30
         Left            =   3540
         Picture         =   "frmListado.frx":56CC
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   29
         Left            =   960
         Picture         =   "frmListado.frx":5757
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Listado facturas proveedores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   22
         Left            =   120
         TabIndex        =   683
         Top             =   120
         Width           =   4995
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   91
         Left            =   300
         TabIndex        =   682
         Top             =   1560
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   21
         Left            =   900
         Picture         =   "frmListado.frx":57E2
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   20
         Left            =   900
         Picture         =   "frmListado.frx":C034
         Top             =   1800
         Width           =   240
      End
   End
   Begin VB.Frame frameListadoCuentas 
      Height          =   5895
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   5475
      Begin VB.TextBox txtPag2 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "frmListado.frx":12886
         Left            =   360
         List            =   "frmListado.frx":12893
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4020
         Width           =   1935
      End
      Begin VB.TextBox txtTitulo 
         Height          =   315
         Left            =   2460
         MaxLength       =   25
         TabIndex        =   10
         Text            =   "EXTRACTOS DE CUENTAS"
         Top             =   540
         Visible         =   0   'False
         Width           =   3495
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   360
         TabIndex        =   63
         Top             =   4920
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtPag2 
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4020
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Salto página por cuenta"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   6
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmListado.frx":128C7
         Left            =   360
         List            =   "frmListado.frx":128DA
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdListExtCta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   12
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N.I.F."
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
         Index           =   121
         Left            =   360
         TabIndex        =   842
         Top             =   2040
         Width           =   405
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado.frx":12921
         ToolTipText     =   "Filtrar NIF"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Saldo en la cuenta"
         Height          =   255
         Index           =   18
         Left            =   360
         TabIndex        =   657
         Top             =   3780
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Título informe"
         Height          =   255
         Index           =   11
         Left            =   1380
         TabIndex        =   295
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   495
         Left            =   360
         TabIndex        =   278
         Top             =   5280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "1ª Página"
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   25
         Top             =   3780
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   4560
         Picture         =   "frmListado.frx":13323
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Formato de extracto"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmListado.frx":133AE
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmListado.frx":13DB0
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Listado extracto de cuentas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmListado.frx":147B2
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   3120
         Picture         =   "frmListado.frx":1483D
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   20
         Top             =   2685
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   2685
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   1000
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha informe"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   24
         Top             =   3120
         Width           =   1005
      End
   End
   Begin VB.Frame frameCtaExpCC 
      Height          =   5640
      Left            =   120
      TabIndex        =   342
      Top             =   0
      Visible         =   0   'False
      Width           =   5715
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Solo mostrar subcentros de reparto"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   353
         Top             =   4320
         Width           =   2775
      End
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Comparativo"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   354
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   30
         Left            =   1320
         TabIndex        =   346
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   30
         Left            =   2520
         TabIndex        =   832
         Text            =   "Text5"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   29
         Left            =   1320
         TabIndex        =   345
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   2520
         TabIndex        =   830
         Text            =   "Text5"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Frame FrameCCComparativo 
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   495
         Left            =   1920
         TabIndex        =   829
         Top             =   4560
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton optCCComparativo 
            Caption         =   "Mes"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   356
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optCCComparativo 
            Caption         =   "Saldo"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   355
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Ver movimientos posteriores"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   352
         Top             =   4320
         Width           =   2415
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   7
         ItemData        =   "frmListado.frx":148C8
         Left            =   4200
         List            =   "frmListado.frx":148CA
         Style           =   2  'Dropdown List
         TabIndex        =   351
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   3
         Left            =   1380
         TabIndex        =   344
         Text            =   "Text2"
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   360
         Text            =   "Text2"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   2
         Left            =   1380
         TabIndex        =   343
         Text            =   "Text2"
         Top             =   1020
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   359
         Text            =   "Text2"
         Top             =   1020
         Width           =   2655
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   4440
         TabIndex        =   358
         Top             =   5160
         Width           =   975
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   6
         ItemData        =   "frmListado.frx":148CC
         Left            =   1620
         List            =   "frmListado.frx":148CE
         Style           =   2  'Dropdown List
         TabIndex        =   349
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   8
         Left            =   2940
         TabIndex        =   350
         Text            =   "Text1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   5
         ItemData        =   "frmListado.frx":148D0
         Left            =   1620
         List            =   "frmListado.frx":148D2
         Style           =   2  'Dropdown List
         TabIndex        =   347
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   7
         Left            =   2940
         TabIndex        =   348
         Text            =   "Text1"
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdCtaExpCC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   357
         Top             =   5160
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   119
         Left            =   240
         TabIndex        =   834
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   480
         TabIndex        =   833
         Top             =   2685
         Width           =   465
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   30
         Left            =   1080
         Picture         =   "frmListado.frx":148D4
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   480
         TabIndex        =   831
         Top             =   2325
         Width           =   465
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   29
         Left            =   1080
         Picture         =   "frmListado.frx":152D6
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   315
         Left            =   180
         TabIndex        =   369
         Top             =   5160
         Width           =   2835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes cálculo"
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
         Index           =   42
         Left            =   4200
         TabIndex        =   368
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta explotación por centro de coste"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Index           =   12
         Left            =   300
         TabIndex        =   367
         Top             =   300
         Width           =   5025
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         Index           =   41
         Left            =   240
         TabIndex        =   366
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   35
         Left            =   540
         TabIndex        =   365
         Top             =   1140
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmListado.frx":15CD8
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   540
         TabIndex        =   364
         Top             =   1500
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmListado.frx":166DA
         Top             =   1470
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   33
         Left            =   960
         TabIndex        =   363
         Top             =   3420
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   32
         Left            =   960
         TabIndex        =   362
         Top             =   3900
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fechas"
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
         Index           =   40
         Left            =   240
         TabIndex        =   361
         Top             =   3240
         Width           =   585
      End
   End
   Begin VB.Frame Frame347 
      Height          =   6360
      Left            =   120
      TabIndex        =   458
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chk347 
         Caption         =   "Datos importados"
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   825
         Top             =   3720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chk347 
         Caption         =   "Datos facturas"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   824
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chk347 
         Caption         =   "Fech Liqui. Clientes"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   738
         Top             =   2760
         Width           =   1815
      End
      Begin VB.OptionButton OptProv 
         Caption         =   "Fecha liquidación"
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   737
         Top             =   2280
         Width           =   1755
      End
      Begin VB.TextBox Text347 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   470
         Text            =   "Text4"
         Top             =   5100
         Width           =   1275
      End
      Begin VB.CommandButton cmdDatosCarta 
         Caption         =   "Datos carta"
         Height          =   315
         Left            =   3600
         TabIndex        =   571
         Top             =   5160
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmListado.frx":170DC
         Left            =   240
         List            =   "frmListado.frx":170EC
         Style           =   2  'Dropdown List
         TabIndex        =   469
         Top             =   5100
         Width           =   1635
      End
      Begin VB.CheckBox chk347 
         Caption         =   "Papel preimpreso"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   467
         Top             =   3120
         Width           =   1815
      End
      Begin VB.OptionButton OptProv 
         Caption         =   "Fecha factura"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   466
         Top             =   1980
         Width           =   1755
      End
      Begin VB.OptionButton OptProv 
         Caption         =   "Fecha recepción"
         Height          =   195
         Index           =   0
         Left            =   3300
         TabIndex        =   465
         Top             =   1680
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.ListBox List3 
         Enabled         =   0   'False
         Height          =   1620
         ItemData        =   "frmListado.frx":1711D
         Left            =   180
         List            =   "frmListado.frx":1711F
         TabIndex        =   474
         Top             =   1740
         Width           =   2775
      End
      Begin VB.TextBox Text347 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   468
         Text            =   "Text2"
         Top             =   4380
         Width           =   4635
      End
      Begin VB.CommandButton cmd347 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   471
         Top             =   5820
         Width           =   855
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   4080
         TabIndex        =   472
         Top             =   5820
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   22
         Left            =   3540
         TabIndex        =   461
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   21
         Left            =   1140
         TabIndex        =   460
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Label24"
         Height          =   495
         Index           =   31
         Left            =   240
         TabIndex        =   699
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   0
         Left            =   2760
         Picture         =   "frmListado.frx":17121
         ToolTipText     =   "Acerca del 347"
         Top             =   5880
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Label24"
         Height          =   255
         Index           =   30
         Left            =   240
         TabIndex        =   698
         Top             =   5520
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe limite"
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
         Index           =   92
         Left            =   1980
         TabIndex        =   695
         Top             =   4860
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo informe"
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
         Index           =   62
         Left            =   240
         TabIndex        =   477
         Top             =   4860
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedores"
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
         Index           =   61
         Left            =   3300
         TabIndex        =   476
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   1260
         Picture         =   "frmListado.frx":17B23
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Index           =   60
         Left            =   180
         TabIndex        =   475
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
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
         Index           =   59
         Left            =   180
         TabIndex        =   473
         Top             =   4140
         Width           =   1080
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   14
         Left            =   2640
         TabIndex        =   463
         Top             =   1020
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   22
         Left            =   3240
         Picture         =   "frmListado.frx":18525
         Top             =   1035
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   21
         Left            =   720
         Picture         =   "frmListado.frx":185B0
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   58
         Left            =   180
         TabIndex        =   462
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Modelo 347"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   459
         Top             =   300
         Width           =   2205
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   15
         Left            =   180
         TabIndex        =   464
         Top             =   1035
         Width           =   615
      End
   End
   Begin VB.Frame FrameRela_x_Cta 
      Height          =   5415
      Left            =   120
      TabIndex        =   783
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkCliproxCtalineas 
         Caption         =   "Comparativo periodo anterior"
         Height          =   255
         Left            =   240
         TabIndex        =   807
         Top             =   4920
         Width           =   2415
      End
      Begin VB.CheckBox chkDesgloseBasexCta 
         Caption         =   "Desglosar cuenta"
         Height          =   255
         Left            =   240
         TabIndex        =   806
         Top             =   4560
         Width           =   2775
      End
      Begin VB.CommandButton cmdRelacion_x_Ctabases 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   811
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Frame FrameProvxGast 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   720
         TabIndex        =   809
         Top             =   3840
         Width           =   4815
         Begin VB.OptionButton optCta_x_gastos 
            Caption         =   "F. recepcion"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   805
            Top             =   30
            Width           =   1575
         End
         Begin VB.OptionButton optCta_x_gastos 
            Caption         =   "F. factura"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   804
            Top             =   30
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   37
         Left            =   3600
         TabIndex        =   803
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   36
         Left            =   1440
         TabIndex        =   800
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   2520
         TabIndex        =   797
         Text            =   "Text5"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   28
         Left            =   1320
         TabIndex        =   796
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   2520
         TabIndex        =   793
         Text            =   "Text5"
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   27
         Left            =   1320
         TabIndex        =   792
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   2520
         TabIndex        =   790
         Text            =   "Text5"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   26
         Left            =   1320
         TabIndex        =   789
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   2520
         TabIndex        =   786
         Text            =   "Text5"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   25
         Left            =   1320
         TabIndex        =   785
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   55
         Left            =   4440
         TabIndex        =   784
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   26
         Left            =   1080
         Picture         =   "frmListado.frx":1863B
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Label27"
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   810
         Top             =   4200
         Width           =   5415
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   23
         Left            =   2760
         TabIndex        =   808
         Top             =   3525
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   37
         Left            =   3360
         Picture         =   "frmListado.frx":1903D
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   802
         Top             =   3525
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   36
         Left            =   1200
         Picture         =   "frmListado.frx":190C8
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   115
         Left            =   120
         TabIndex        =   801
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Clientes x cta gastos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   23
         Left            =   120
         TabIndex        =   799
         Top             =   360
         Width           =   5565
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   70
         Left            =   480
         TabIndex        =   798
         Top             =   2760
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   28
         Left            =   1080
         Picture         =   "frmListado.frx":19153
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   114
         Left            =   120
         TabIndex        =   795
         Top             =   2160
         Width           =   1680
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   27
         Left            =   1080
         Picture         =   "frmListado.frx":19B55
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   69
         Left            =   480
         TabIndex        =   794
         Top             =   2445
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   68
         Left            =   480
         TabIndex        =   791
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bases"
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
         Index           =   113
         Left            =   120
         TabIndex        =   788
         Top             =   960
         Width           =   1140
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   25
         Left            =   1080
         Picture         =   "frmListado.frx":1A557
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   67
         Left            =   480
         TabIndex        =   787
         Top             =   1365
         Width           =   465
      End
   End
   Begin VB.Frame FrameLiq 
      Height          =   5805
      Left            =   60
      TabIndex        =   279
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   33
         Left            =   3720
         TabIndex        =   288
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   1320
         TabIndex        =   712
         Text            =   "Text5"
         Top             =   4800
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   22
         Left            =   240
         TabIndex        =   711
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame11 
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   3120
         TabIndex        =   700
         Top             =   2520
         Width           =   1995
         Begin VB.OptionButton optModeloLiq 
            Caption         =   "303"
            Height          =   195
            Index           =   4
            Left            =   840
            TabIndex        =   840
            Top             =   1680
            Width           =   675
         End
         Begin VB.OptionButton optModeloLiq 
            Caption         =   "332"
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   710
            Top             =   1320
            Width           =   675
         End
         Begin VB.OptionButton optModeloLiq 
            Caption         =   "330"
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   709
            Top             =   960
            Width           =   675
         End
         Begin VB.OptionButton optModeloLiq 
            Caption         =   "320"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   702
            Top             =   600
            Width           =   675
         End
         Begin VB.OptionButton optModeloLiq 
            Caption         =   "300"
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   701
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.CheckBox chkIVAdetallado 
         Caption         =   "Tipos IVA detallados"
         Height          =   195
         Left            =   3180
         TabIndex        =   658
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkLiqDefinitiva 
         Caption         =   "Liquidación definitiva"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   3180
         TabIndex        =   289
         Top             =   1800
         Width           =   1875
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         Height          =   2595
         ItemData        =   "frmListado.frx":1AF59
         Left            =   240
         List            =   "frmListado.frx":1AF5B
         TabIndex        =   290
         Top             =   1860
         Width           =   2775
      End
      Begin VB.TextBox txtperiodo 
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   287
         Text            =   "Text1"
         Top             =   1185
         Width           =   615
      End
      Begin VB.TextBox txtperiodo 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   285
         Text            =   "Text1"
         Top             =   1185
         Width           =   615
      End
      Begin VB.TextBox txtperiodo 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   283
         Text            =   "Text1"
         Top             =   1185
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   291
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdCanListExtr 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4320
         TabIndex        =   293
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Modelo 303"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   120
         Left            =   3240
         TabIndex        =   839
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   101
         Left            =   3720
         TabIndex        =   714
         Top             =   840
         Width           =   1200
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   33
         Left            =   4920
         Picture         =   "frmListado.frx":1AF5D
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Cta. cuotas a compensar"
         Height          =   195
         Index           =   58
         Left            =   240
         TabIndex        =   713
         Top             =   4560
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   22
         Left            =   2160
         Picture         =   "frmListado.frx":1AFE8
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   240
         TabIndex        =   294
         Top             =   5400
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   1320
         Picture         =   "frmListado.frx":1B9EA
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Left            =   240
         TabIndex        =   292
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Index           =   32
         Left            =   2640
         TabIndex        =   286
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   195
         Index           =   27
         Left            =   1320
         TabIndex        =   284
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periódo"
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
         Index           =   30
         Left            =   240
         TabIndex        =   282
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Ini."
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   281
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "Liquidación IVA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   9
         Left            =   1440
         TabIndex        =   280
         Top             =   300
         Width           =   2535
      End
   End
   Begin VB.Frame frameIVA 
      Height          =   5775
      Left            =   60
      TabIndex        =   260
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame Frame4 
         Height          =   2055
         Left            =   360
         TabIndex        =   273
         Top             =   2280
         Visible         =   0   'False
         Width           =   5295
         Begin MSComctlLib.ProgressBar pb5 
            Height          =   375
            Left            =   240
            TabIndex        =   274
            Top             =   960
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label11 
            Caption         =   "Empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   275
            Top             =   360
            Width           =   4695
         End
      End
      Begin VB.CommandButton cmdCertIVA 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   267
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   14
         Left            =   4200
         TabIndex        =   265
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   13
         Left            =   1320
         TabIndex        =   264
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   262
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   4560
         TabIndex        =   269
         Top             =   5040
         Width           =   975
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   266
         Top             =   2950
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         Height          =   1815
         ItemData        =   "frmListado.frx":1C3EC
         Left            =   360
         List            =   "frmListado.frx":1C3EE
         TabIndex        =   277
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de I.V.A."
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
         Index           =   29
         Left            =   360
         TabIndex        =   276
         Top             =   3000
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Index           =   28
         Left            =   360
         TabIndex        =   272
         Top             =   3480
         Width           =   825
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1440
         Picture         =   "frmListado.frx":1C3F0
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Desde"
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
         Left            =   3240
         TabIndex        =   271
         Top             =   2280
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   14
         Left            =   3960
         Picture         =   "frmListado.frx":1CDF2
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Desde"
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
         Left            =   480
         TabIndex        =   270
         Top             =   2280
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   13
         Left            =   1080
         Picture         =   "frmListado.frx":1CE7D
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   27
         Left            =   480
         TabIndex        =   268
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   12
         Left            =   1680
         Picture         =   "frmListado.frx":1CF08
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   26
         Left            =   480
         TabIndex        =   263
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Certificado declaración de I.V.A."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   8
         Left            =   540
         TabIndex        =   261
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame frameBorrarClientes 
      Height          =   4440
      Left            =   120
      TabIndex        =   498
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame FrBorrePorEjercicios 
         Height          =   2535
         Left            =   120
         TabIndex        =   836
         Top             =   720
         Width           =   5175
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   255
            Index           =   75
            Left            =   840
            TabIndex        =   838
            Top             =   1080
            Width           =   3855
         End
         Begin VB.Label Label3 
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
            Height          =   255
            Index           =   74
            Left            =   360
            TabIndex        =   837
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkBorreFactura 
         Caption         =   "Borre por ejercicios"
         Height          =   255
         Left            =   240
         TabIndex        =   835
         Top             =   3360
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin MSComctlLib.ProgressBar pb8 
         Height          =   375
         Left            =   180
         TabIndex        =   518
         Top             =   3840
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Frame frameTapa 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         TabIndex        =   517
         Top             =   720
         Width           =   4995
      End
      Begin VB.TextBox txtNumFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3840
         TabIndex        =   505
         Text            =   "Text1"
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox txtNumFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   504
         Text            =   "Text1"
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   3
         Left            =   3660
         TabIndex        =   503
         Text            =   "Text1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   502
         Text            =   "Text1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrarFacCli 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   509
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   22
         Left            =   4320
         TabIndex        =   511
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   24
         Left            =   3840
         TabIndex        =   507
         Top             =   2820
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   23
         Left            =   1260
         TabIndex        =   506
         Top             =   2835
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   49
         Left            =   360
         TabIndex        =   516
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   48
         Left            =   3060
         TabIndex        =   515
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   72
         Left            =   360
         TabIndex        =   514
         Top             =   1740
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   71
         Left            =   360
         TabIndex        =   513
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   70
         Left            =   3060
         TabIndex        =   512
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Index           =   69
         Left            =   360
         TabIndex        =   510
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   68
         Left            =   360
         TabIndex        =   508
         Top             =   2880
         Width           =   450
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   24
         Left            =   3600
         Picture         =   "frmListado.frx":1CF93
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   65
         Left            =   3060
         TabIndex        =   501
         Top             =   2880
         Width           =   420
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   23
         Left            =   960
         Picture         =   "frmListado.frx":1D01E
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   64
         Left            =   360
         TabIndex        =   500
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Borrar registro clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   16
         Left            =   960
         TabIndex        =   499
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdCancelarAccion 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   5400
      TabIndex        =   736
      Top             =   6480
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameExploCon 
      Height          =   6375
      Left            =   0
      TabIndex        =   630
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkCtaExpCon 
         Caption         =   "Desglosar empresa"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   644
         Top             =   2940
         Width           =   1875
      End
      Begin VB.ListBox List7 
         Enabled         =   0   'False
         Height          =   2400
         ItemData        =   "frmListado.frx":1D0A9
         Left            =   1500
         List            =   "frmListado.frx":1D0AB
         TabIndex        =   654
         Top             =   3300
         Width           =   2775
      End
      Begin VB.Frame FrameCtasExploC 
         Height          =   1035
         Left            =   240
         TabIndex        =   653
         Top             =   1740
         Width           =   5865
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   642
            Top             =   720
            Width           =   1245
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   641
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   640
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   639
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   638
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4320
            TabIndex        =   637
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   2940
            TabIndex        =   636
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   1560
            TabIndex        =   635
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExploC 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   180
            TabIndex        =   634
            Top             =   240
            Value           =   1  'Checked
            Width           =   1185
         End
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   28
         Left            =   240
         TabIndex        =   631
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   19
         ItemData        =   "frmListado.frx":1D0AD
         Left            =   4440
         List            =   "frmListado.frx":1D0AF
         Style           =   2  'Dropdown List
         TabIndex        =   633
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkCtaExpCon 
         Caption         =   "Imprimir acumulados y movimientos del mes"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   643
         Top             =   2940
         Width           =   3495
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   31
         Left            =   5220
         TabIndex        =   647
         Top             =   5820
         Width           =   975
      End
      Begin VB.CommandButton cmdCtaexplcmp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   646
         Top             =   5820
         Width           =   975
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   18
         Left            =   3480
         TabIndex        =   632
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar pb10 
         Height          =   375
         Left            =   180
         TabIndex        =   645
         Top             =   5820
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label2 
         Caption         =   "Label23"
         Height          =   195
         Index           =   28
         Left            =   180
         TabIndex        =   656
         Top             =   5520
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Index           =   86
         Left            =   240
         TabIndex        =   655
         Top             =   3300
         Width           =   825
      End
      Begin VB.Image Image10 
         Height          =   240
         Left            =   1200
         Picture         =   "frmListado.frx":1D0B1
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nivel"
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
         Index           =   87
         Left            =   240
         TabIndex        =   649
         Top             =   1500
         Width           =   405
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   28
         Left            =   1440
         Picture         =   "frmListado.frx":1DAB3
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   89
         Left            =   240
         TabIndex        =   652
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label22 
         Caption         =   "Cuenta de explotación consolidada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   720
         TabIndex        =   651
         Top             =   300
         Width           =   5235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   88
         Left            =   4440
         TabIndex        =   650
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Index           =   85
         Left            =   3480
         TabIndex        =   648
         Top             =   840
         Width           =   330
      End
   End
   Begin VB.Frame FrameBalPersoConso 
      Height          =   7035
      Left            =   120
      TabIndex        =   716
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Frame FrameTapa3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   720
         TabIndex        =   717
         Top             =   2760
         Width           =   3795
      End
      Begin VB.ListBox List8 
         Enabled         =   0   'False
         Height          =   2400
         ItemData        =   "frmListado.frx":1DB3E
         Left            =   1500
         List            =   "frmListado.frx":1DB40
         TabIndex        =   724
         Top             =   3480
         Width           =   4215
      End
      Begin VB.TextBox TextDescBalance 
         BackColor       =   &H80000018&
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
         Index           =   1
         Left            =   1800
         TabIndex        =   728
         Text            =   "Text1"
         Top             =   1140
         Width           =   4035
      End
      Begin VB.TextBox txtNumBal 
         Height          =   315
         Index           =   1
         Left            =   780
         TabIndex        =   718
         Text            =   "Text1"
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   20
         Left            =   2880
         TabIndex        =   720
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   21
         ItemData        =   "frmListado.frx":1DB42
         Left            =   1560
         List            =   "frmListado.frx":1DB44
         Style           =   2  'Dropdown List
         TabIndex        =   722
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   21
         Left            =   2880
         TabIndex        =   723
         Text            =   "Text1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   20
         ItemData        =   "frmListado.frx":1DB46
         Left            =   1560
         List            =   "frmListado.frx":1DB48
         Style           =   2  'Dropdown List
         TabIndex        =   719
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdBalanPersConso 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   726
         Top             =   6420
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   50
         Left            =   4800
         TabIndex        =   727
         Top             =   6420
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   34
         Left            =   180
         TabIndex        =   725
         Top             =   6540
         Width           =   1455
      End
      Begin VB.CheckBox chkBalPerCompaCon 
         Caption         =   "Comparativo"
         Height          =   255
         Left            =   720
         TabIndex        =   721
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Image Image11 
         Height          =   240
         Left            =   1200
         Picture         =   "frmListado.frx":1DB4A
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Index           =   105
         Left            =   240
         TabIndex        =   735
         Top             =   3480
         Width           =   825
      End
      Begin VB.Image ImgNumBal 
         Height          =   240
         Index           =   1
         Left            =   420
         Picture         =   "frmListado.frx":1E54C
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
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
         Index           =   104
         Left            =   180
         TabIndex        =   734
         Top             =   780
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   103
         Left            =   240
         TabIndex        =   733
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   62
         Left            =   720
         TabIndex        =   732
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Balances configurables consolidados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   731
         Top             =   240
         Width           =   5595
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   102
         Left            =   180
         TabIndex        =   730
         Top             =   6240
         Width           =   1200
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   34
         Left            =   1500
         Picture         =   "frmListado.frx":1EF4E
         Top             =   6240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   59
         Left            =   720
         TabIndex        =   729
         Top             =   3000
         Width           =   615
      End
   End
   Begin VB.Frame frameDiarioHco 
      Height          =   5655
      Left            =   120
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chkTotalAsiento 
         Caption         =   "Mostrar totales por asiento"
         Height          =   255
         Left            =   2640
         TabIndex        =   715
         Top             =   3480
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtReemisionDiario 
         Height          =   285
         Left            =   1200
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   4140
         Width           =   3615
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   67
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   74
         Top             =   5040
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   71
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   77
         Top             =   5040
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdAceptarHco 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2520
         TabIndex        =   73
         Top             =   5040
         Width           =   1155
      End
      Begin VB.TextBox txtAsiento 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   65
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   69
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtAsiento 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   66
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   68
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   76
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   75
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   3480
         TabIndex        =   70
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Título"
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
         Index           =   35
         Left            =   360
         TabIndex        =   322
         Top             =   3960
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F. Listado"
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
         Index           =   12
         Left            =   360
         TabIndex        =   149
         Top             =   3240
         Width           =   795
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   8
         Left            =   1200
         Picture         =   "frmListado.frx":1EFD9
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   3120
         Picture         =   "frmListado.frx":1F064
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmListado.frx":1F0EF
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   960
         Picture         =   "frmListado.frx":1FAF1
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmListado.frx":1FB7C
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Impresión de diario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   87
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   86
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Asiento"
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
         Index           =   5
         Left            =   360
         TabIndex        =   85
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   84
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   83
         Top             =   2805
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   82
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
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
         Index           =   4
         Left            =   360
         TabIndex        =   81
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   3
         Left            =   360
         TabIndex        =   80
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   79
         Top             =   2070
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   78
         Top             =   2805
         Width           =   495
      End
   End
   Begin VB.Frame frameAsiento 
      Height          =   2895
      Left            =   60
      TabIndex        =   52
      Top             =   20
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   2400
         TabIndex        =   62
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton AsientoAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   60
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDesAs 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtAS 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtDesAs 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtAS 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Listado de asientos predefinidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   5
         Left            =   360
         TabIndex        =   58
         Top             =   1245
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   4
         Left            =   360
         TabIndex        =   57
         Top             =   765
         Width           =   555
      End
   End
   Begin VB.Frame FramePersa 
      Height          =   2535
      Left            =   60
      TabIndex        =   588
      Top             =   0
      Visible         =   0   'False
      Width           =   3435
      Begin VB.CheckBox Check4 
         Caption         =   "Llevar a diskette"
         Height          =   255
         Left            =   420
         TabIndex        =   594
         Top             =   1200
         Width           =   2415
      End
      Begin VB.OptionButton optActual 
         Caption         =   "Siguiente"
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   593
         Top             =   840
         Width           =   1035
      End
      Begin VB.OptionButton optActual 
         Caption         =   "Actual"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   592
         Top             =   840
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   29
         Left            =   2160
         TabIndex        =   591
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdPersa 
         Caption         =   "&Generar"
         Height          =   375
         Left            =   1080
         TabIndex        =   590
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblPersa2 
         Caption         =   "Label21"
         Height          =   255
         Left            =   120
         TabIndex        =   596
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblPersa 
         Caption         =   "Label21"
         Height          =   255
         Left            =   1020
         TabIndex        =   595
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Traspaso PERSA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   19
         Left            =   420
         TabIndex        =   589
         Top             =   300
         Width           =   2490
      End
   End
   Begin VB.Frame FrameCopyBalan 
      Height          =   3735
      Left            =   0
      TabIndex        =   813
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkCopyBalan 
         Caption         =   "Copiar las cuentas "
         Height          =   255
         Left            =   720
         TabIndex        =   823
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CommandButton cmdCopyBalan 
         Caption         =   "Copiar"
         Height          =   375
         Left            =   3600
         TabIndex        =   822
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   57
         Left            =   4800
         TabIndex        =   820
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox TextDescBalance 
         BackColor       =   &H80000018&
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
         Index           =   3
         Left            =   1740
         TabIndex        =   819
         Text            =   "Text1"
         Top             =   2160
         Width           =   4035
      End
      Begin VB.TextBox txtNumBal 
         Height          =   315
         Index           =   3
         Left            =   720
         TabIndex        =   818
         Text            =   "Text1"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TextDescBalance 
         BackColor       =   &H80000018&
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
         Left            =   1680
         TabIndex        =   815
         Text            =   "Text1"
         Top             =   1080
         Width           =   4035
      End
      Begin VB.TextBox txtNumBal 
         BackColor       =   &H80000018&
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
         Left            =   720
         TabIndex        =   814
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Index           =   117
         Left            =   120
         TabIndex        =   821
         Top             =   1800
         Width           =   555
      End
      Begin VB.Image ImgNumBal 
         Height          =   240
         Index           =   3
         Left            =   360
         Picture         =   "frmListado.frx":2057E
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label17 
         Caption         =   "Copiar balances configurables"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   3
         Left            =   720
         TabIndex        =   817
         Top             =   240
         Width           =   4875
      End
      Begin VB.Image ImgNumBal 
         Height          =   240
         Index           =   2
         Left            =   360
         Picture         =   "frmListado.frx":20F80
         Top             =   1080
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DESTINO"
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
         Index           =   116
         Left            =   120
         TabIndex        =   816
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame FramePresu 
      Height          =   4095
      Left            =   60
      TabIndex        =   213
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtMes 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   218
         Tag             =   "Mes hasta|N|S|1|12|||||"
         Text            =   "Text1"
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdPresupuestos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   220
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   4560
         TabIndex        =   221
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtAno 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   219
         Tag             =   "Año hasta|N|S|1900|10000|||||"
         Text            =   "Text2"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtAno 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   217
         Tag             =   "Año desde|N|S|1900|10000|||||"
         Text            =   "Text2"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtMes 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   216
         Tag             =   "Mes|N|S|1|12|||||"
         Text            =   "Text1"
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   2880
         TabIndex        =   223
         Text            =   "Text5"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   222
         Text            =   "Text5"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   11
         Left            =   1680
         TabIndex        =   215
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   214
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes/Año"
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
         Index           =   22
         Left            =   360
         TabIndex        =   230
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   23
         Left            =   840
         TabIndex        =   229
         Top             =   2835
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   22
         Left            =   840
         TabIndex        =   228
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Listado de presupuestos sobre cuentas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   6
         Left            =   420
         TabIndex        =   227
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   21
         Left            =   360
         TabIndex        =   226
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   21
         Left            =   840
         TabIndex        =   225
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   11
         Left            =   1440
         Picture         =   "frmListado.frx":21982
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListado.frx":22384
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   20
         Left            =   840
         TabIndex        =   224
         Top             =   1125
         Width           =   615
      End
   End
   Begin VB.Frame FrameEvolSaldo 
      Height          =   4815
      Left            =   60
      TabIndex        =   758
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbEjercicios 
         Height          =   315
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   826
         Top             =   2160
         Width           =   4095
      End
      Begin VB.CheckBox chkEvolSalMeses 
         Caption         =   "Mostrar meses sin movimientos"
         Height          =   195
         Left            =   120
         TabIndex        =   782
         Top             =   3960
         Width           =   2895
      End
      Begin VB.CommandButton cmdEvolMensSald 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   772
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   24
         Left            =   1440
         TabIndex        =   761
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2640
         TabIndex        =   779
         Text            =   "Text5"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   23
         Left            =   1440
         TabIndex        =   760
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   2640
         TabIndex        =   776
         Text            =   "Text5"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Frame FrameNivelEvolSaldo 
         Height          =   1035
         Left            =   120
         TabIndex        =   759
         Top             =   2760
         Width           =   5865
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   771
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   770
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   769
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   768
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   767
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   766
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   765
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   764
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   763
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox ChkEvolSaldo 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   762
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "    Nivel     "
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
            Index           =   111
            Left            =   360
            TabIndex        =   773
            Top             =   0
            Width           =   810
         End
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   54
         Left            =   4920
         TabIndex        =   774
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ejercicio"
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
         Index           =   118
         Left            =   240
         TabIndex        =   827
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Label26"
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   781
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   24
         Left            =   1200
         Picture         =   "frmListado.frx":22D86
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   66
         Left            =   600
         TabIndex        =   780
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   65
         Left            =   600
         TabIndex        =   778
         Top             =   1125
         Width           =   465
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   23
         Left            =   1200
         Picture         =   "frmListado.frx":23788
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   112
         Left            =   240
         TabIndex        =   777
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Evolución mensual de saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   21
         Left            =   1320
         TabIndex        =   775
         Top             =   360
         Width           =   4125
      End
   End
   Begin VB.Frame frameComparativo 
      Height          =   3795
      Left            =   60
      TabIndex        =   478
      Top             =   20
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkExploCompaPorcentual 
         Caption         =   "Listado con porcentajes"
         Height          =   255
         Left            =   180
         TabIndex        =   492
         Top             =   3180
         Width           =   2175
      End
      Begin VB.CheckBox chkExploCompa 
         Caption         =   "Mensual"
         Height          =   195
         Left            =   3840
         TabIndex        =   481
         Top             =   1080
         Width           =   1275
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   13
         Left            =   2220
         TabIndex        =   480
         Text            =   "Text1"
         Top             =   1020
         Width           =   855
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   21
         Left            =   4920
         TabIndex        =   494
         Top             =   3120
         Width           =   975
      End
      Begin VB.Frame FrameComp2 
         Height          =   1305
         Left            =   120
         TabIndex        =   495
         Top             =   1500
         Width           =   5865
         Begin VB.CheckBox chkcmp 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   482
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   483
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   484
            Top             =   240
            Width           =   1005
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   485
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   486
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   487
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   491
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   488
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   489
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkcmp 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   490
            Top             =   720
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdComparativo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   493
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   13
         ItemData        =   "frmListado.frx":2418A
         Left            =   900
         List            =   "frmListado.frx":2418C
         Style           =   2  'Dropdown List
         TabIndex        =   479
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta de explotación comparativa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   15
         Left            =   540
         TabIndex        =   497
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Index           =   63
         Left            =   300
         TabIndex        =   496
         Top             =   1080
         Width           =   345
      End
   End
   Begin VB.Frame frameCuentas 
      Height          =   5175
      Left            =   60
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame Frame2 
         Height          =   1605
         Left            =   120
         TabIndex        =   50
         Top             =   2640
         Width           =   5865
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   34
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   35
            Top             =   240
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   36
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   37
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   39
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   40
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   41
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   42
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Remarcar"
            Height          =   195
            Left            =   1260
            TabIndex        =   51
            Top             =   1245
            Width           =   795
         End
      End
      Begin VB.ComboBox cmbCuenta 
         Height          =   315
         ItemData        =   "frmListado.frx":2418E
         Left            =   2520
         List            =   "frmListado.frx":2419B
         Style           =   2  'Dropdown List
         TabIndex        =   739
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CheckBox chkCuetas_x_nombre 
         Caption         =   "Ordenar por nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4920
         TabIndex        =   61
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton AceptarCuentas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   45
         Top             =   4560
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Datos fiscales"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   32
         Top             =   2160
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Informe cuentas"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   31
         Top             =   2160
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   30
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label25 
         Caption         =   "Seleccionar 347:"
         Height          =   255
         Left            =   960
         TabIndex        =   740
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmListado.frx":241B3
         Top             =   960
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmListado.frx":24BB5
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   49
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   48
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Listado de cuentas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   1380
         TabIndex        =   47
         Top             =   300
         Width           =   3555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Left            =   240
         TabIndex        =   46
         Top             =   720
         Width           =   600
      End
   End
   Begin VB.Frame frameCCostSaldos 
      Height          =   4755
      Left            =   60
      TabIndex        =   324
      Top             =   20
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdSaldosCC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   332
         Top             =   3960
         Width           =   1035
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   6
         Left            =   2880
         TabIndex        =   331
         Text            =   "Text1"
         Top             =   3060
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   4
         ItemData        =   "frmListado.frx":255B7
         Left            =   1440
         List            =   "frmListado.frx":255B9
         Style           =   2  'Dropdown List
         TabIndex        =   330
         Top             =   3060
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   5
         Left            =   2880
         TabIndex        =   329
         Text            =   "Text1"
         Top             =   2580
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   3
         ItemData        =   "frmListado.frx":255BB
         Left            =   1440
         List            =   "frmListado.frx":255BD
         Style           =   2  'Dropdown List
         TabIndex        =   327
         Top             =   2580
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   4020
         TabIndex        =   333
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   338
         Text            =   "Text2"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   326
         Text            =   "Text2"
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   336
         Text            =   "Text2"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   325
         Text            =   "Text2"
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   39
         Left            =   240
         TabIndex        =   341
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   31
         Left            =   720
         TabIndex        =   340
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   30
         Left            =   720
         TabIndex        =   339
         Top             =   2580
         Width           =   615
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmListado.frx":255BF
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   29
         Left            =   300
         TabIndex        =   337
         Top             =   1740
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado.frx":25FC1
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   300
         TabIndex        =   335
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         Index           =   38
         Left            =   240
         TabIndex        =   334
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Saldos centros de coste"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   11
         Left            =   960
         TabIndex        =   328
         Top             =   360
         Width           =   3525
      End
   End
   Begin VB.Frame frameCtaConcepto 
      Height          =   4575
      Left            =   60
      TabIndex        =   88
      Top             =   20
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4320
         TabIndex        =   97
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarTotalesCta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   96
         Top             =   3960
         Width           =   975
      End
      Begin VB.CheckBox chkMeses 
         Caption         =   "Desglosar meses"
         Height          =   255
         Left            =   600
         TabIndex        =   107
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtDescConce 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Left            =   1440
         TabIndex        =   95
         Text            =   "Text1"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   91
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   93
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   92
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   3720
         TabIndex        =   94
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2640
         TabIndex        =   90
         Text            =   "Text5"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2640
         TabIndex        =   89
         Text            =   "Text5"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Image ImgConcepto 
         Height          =   240
         Left            =   1140
         Picture         =   "frmListado.frx":269C3
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Index           =   8
         Left            =   240
         TabIndex        =   106
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   104
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   103
         Top             =   2445
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   102
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   101
         Top             =   2445
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   6
         Left            =   3360
         Picture         =   "frmListado.frx":273C5
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   5
         Left            =   1200
         Picture         =   "frmListado.frx":27450
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Totales por cuenta y concepto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   660
         TabIndex        =   100
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   7
         Left            =   240
         TabIndex        =   99
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   6
         Left            =   240
         TabIndex        =   98
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   5
         Left            =   1200
         Picture         =   "frmListado.frx":274DB
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   4
         Left            =   1200
         Picture         =   "frmListado.frx":27EDD
         Top             =   1560
         Width           =   240
      End
   End
   Begin VB.Frame frameLibroDiario 
      Height          =   5535
      Left            =   60
      TabIndex        =   299
      Top             =   20
      Visible         =   0   'False
      Width           =   5715
      Begin VB.TextBox txtLibroOf 
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   305
         Text            =   "Text1"
         Top             =   3480
         Width           =   735
      End
      Begin VB.CheckBox chkRenumerar 
         Height          =   195
         Left            =   2340
         TabIndex        =   303
         Top             =   2640
         Width           =   195
      End
      Begin VB.TextBox txtLibroOf 
         Height          =   285
         Index           =   0
         Left            =   2340
         TabIndex        =   304
         Text            =   "Text1"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   17
         Left            =   2280
         TabIndex        =   302
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   16
         Left            =   3000
         TabIndex        =   301
         Top             =   1200
         Width           =   1155
      End
      Begin MSComctlLib.ProgressBar pb6 
         Height          =   375
         Left            =   120
         TabIndex        =   310
         Top             =   5040
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdLibroDiario 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   308
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   4560
         TabIndex        =   309
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtExplo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2340
         TabIndex        =   307
         Text            =   "Text1"
         Top             =   4380
         Width           =   1335
      End
      Begin VB.TextBox txtExplo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2340
         TabIndex        =   306
         Text            =   "Text1"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   15
         Left            =   900
         TabIndex        =   300
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "Acumulado haber"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   321
         Top             =   4440
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nº página"
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   320
         Top             =   3540
         Width           =   705
      End
      Begin VB.Label Label9 
         Caption         =   "Renumerar"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   319
         Top             =   2580
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nº asiento"
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   318
         Top             =   3060
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha informe"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   317
         Top             =   2170
         Width           =   1050
      End
      Begin VB.Label Label9 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   3000
         TabIndex        =   316
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   900
         TabIndex        =   315
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Opciones informe"
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
         Index           =   36
         Left            =   180
         TabIndex        =   314
         Top             =   1860
         Width           =   1470
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   17
         Left            =   2040
         Picture         =   "frmListado.frx":288DF
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   16
         Left            =   3480
         Picture         =   "frmListado.frx":2896A
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Acumulado debe"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   313
         Top             =   4020
         Width           =   1200
      End
      Begin VB.Label Label14 
         Caption         =   "Libro diario oficial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1320
         TabIndex        =   312
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fechas"
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
         Index           =   34
         Left            =   240
         TabIndex        =   311
         Top             =   660
         Width           =   585
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   15
         Left            =   1380
         Picture         =   "frmListado.frx":289F5
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame frameccporcta 
      Height          =   5235
      Left            =   60
      TabIndex        =   433
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin MSComctlLib.ProgressBar pb7 
         Height          =   375
         Left            =   120
         TabIndex        =   456
         Top             =   4620
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.CommandButton cmdCtapoCC 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   450
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   19
         Left            =   4200
         TabIndex        =   455
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   6
         Left            =   1560
         TabIndex        =   448
         Text            =   "Text2"
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   2640
         TabIndex        =   451
         Text            =   "Text2"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   7
         Left            =   1560
         TabIndex        =   449
         Text            =   "Text2"
         Top             =   3720
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   2640
         TabIndex        =   447
         Text            =   "Text2"
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   20
         Left            =   3720
         TabIndex        =   443
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   19
         Left            =   1440
         TabIndex        =   442
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2640
         TabIndex        =   437
         Text            =   "Text5"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2640
         TabIndex        =   436
         Text            =   "Text5"
         Top             =   1260
         Width           =   2535
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   17
         Left            =   1440
         TabIndex        =   435
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   434
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Label20"
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   457
         Top             =   4320
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         Index           =   57
         Left            =   120
         TabIndex        =   454
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   660
         TabIndex        =   453
         Top             =   3300
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   7
         Left            =   1320
         Picture         =   "frmListado.frx":28A80
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   660
         TabIndex        =   452
         Top             =   3735
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   6
         Left            =   1260
         Picture         =   "frmListado.frx":29482
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   446
         Top             =   2445
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   56
         Left            =   180
         TabIndex        =   445
         Top             =   2220
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   444
         Top             =   2445
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   20
         Left            =   3480
         Picture         =   "frmListado.frx":29E84
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   19
         Left            =   1200
         Picture         =   "frmListado.frx":29F0F
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "Detalle de explotación centro de coste"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   441
         Top             =   360
         Width           =   4875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   55
         Left            =   120
         TabIndex        =   440
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   45
         Left            =   600
         TabIndex        =   439
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   44
         Left            =   600
         TabIndex        =   438
         Top             =   1305
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   1200
         Picture         =   "frmListado.frx":29F9A
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   16
         Left            =   1200
         Picture         =   "frmListado.frx":2A99C
         Top             =   1260
         Width           =   240
      End
   End
   Begin VB.Frame frameExplotacion 
      Height          =   5535
      Left            =   60
      TabIndex        =   150
      Top             =   20
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Frame FrameCtasExplo 
         Height          =   1035
         Left            =   240
         TabIndex        =   174
         Top             =   1800
         Width           =   5865
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   154
            Top             =   240
            Value           =   1  'Checked
            Width           =   1005
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   155
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   156
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   157
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   158
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   159
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   160
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   161
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   162
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   163
            Top             =   720
            Width           =   1245
         End
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   4
         Left            =   3480
         TabIndex        =   152
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar pb3 
         Height          =   375
         Left            =   120
         TabIndex        =   181
         Top             =   5040
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.CommandButton cmdCtaExplo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   169
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5280
         TabIndex        =   170
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtExplo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3840
         TabIndex        =   168
         Text            =   "Text1"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtExplo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   167
         Text            =   "Text1"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtExplo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   166
         Text            =   "Text1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtExplo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   165
         Text            =   "Text1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CheckBox chkExplotacion 
         Caption         =   "Imprimir acumulados y movimientos del mes"
         Height          =   255
         Left            =   240
         TabIndex        =   164
         Top             =   3060
         Width           =   3495
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   2
         ItemData        =   "frmListado.frx":2B39E
         Left            =   4440
         List            =   "frmListado.frx":2B3A0
         Style           =   2  'Dropdown List
         TabIndex        =   153
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   151
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Index           =   31
         Left            =   3480
         TabIndex        =   298
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label9 
         Caption         =   "Finales(Haber)"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   180
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Iniciales(Debe)"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   179
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   178
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Acumuladas"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   177
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " Existencias acumuladas  "
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
         Index           =   16
         Left            =   240
         TabIndex        =   176
         Top             =   3480
         Width           =   2145
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Left            =   120
         Top             =   3600
         Width           =   6255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nivel"
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
         Index           =   15
         Left            =   240
         TabIndex        =   175
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   14
         Left            =   4440
         TabIndex        =   173
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Cuenta de explotación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1500
         TabIndex        =   172
         Top             =   300
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   13
         Left            =   240
         TabIndex        =   171
         Top             =   840
         Width           =   1200
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   9
         Left            =   1440
         Picture         =   "frmListado.frx":2B3A2
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.Frame FrameAce 
      Height          =   5775
      Left            =   60
      TabIndex        =   597
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   4260
         TabIndex        =   623
         Top             =   2280
         Width           =   1695
         Begin VB.OptionButton optAceAcum 
            Caption         =   "Mensual"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   603
            Top             =   660
            Width           =   1095
         End
         Begin VB.OptionButton optAceAcum 
            Caption         =   "Acumulado"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   602
            Top             =   180
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   598
         Text            =   "A12345"
         Top             =   840
         Width           =   1170
      End
      Begin VB.ListBox List6 
         Enabled         =   0   'False
         Height          =   1620
         ItemData        =   "frmListado.frx":2B42D
         Left            =   1380
         List            =   "frmListado.frx":2B42F
         TabIndex        =   619
         Top             =   1980
         Width           =   2775
      End
      Begin VB.Frame Frame7 
         Height          =   1035
         Left            =   180
         TabIndex        =   606
         Top             =   3900
         Width           =   5865
         Begin VB.CheckBox chkAce 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4560
            TabIndex        =   616
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3480
            TabIndex        =   615
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   2400
            TabIndex        =   614
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   1200
            TabIndex        =   613
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   612
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   4560
            TabIndex        =   611
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   3480
            TabIndex        =   610
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   2400
            TabIndex        =   609
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   608
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox chkAce 
            Caption         =   "Último:  "
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   607
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "    Nivel     "
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
            Index           =   81
            Left            =   120
            TabIndex        =   617
            Top             =   0
            Width           =   810
         End
      End
      Begin VB.OptionButton optAce 
         Caption         =   "Siguiente"
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   601
         Top             =   1380
         Width           =   1155
      End
      Begin VB.OptionButton optAce 
         Caption         =   "Actual"
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   600
         Top             =   1380
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.CommandButton cmdACE 
         Caption         =   "&Generar"
         Height          =   375
         Left            =   3780
         TabIndex        =   618
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   30
         Left            =   4860
         TabIndex        =   620
         Top             =   5160
         Width           =   975
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   18
         ItemData        =   "frmListado.frx":2B431
         Left            =   1380
         List            =   "frmListado.frx":2B433
         Style           =   2  'Dropdown List
         TabIndex        =   599
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "BM Code"
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
         Index           =   84
         Left            =   180
         TabIndex        =   622
         Top             =   900
         Width           =   720
      End
      Begin VB.Image Image9 
         Height          =   240
         Left            =   1080
         Picture         =   "frmListado.frx":2B435
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Index           =   83
         Left            =   180
         TabIndex        =   621
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Traspaso ACE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   20
         Left            =   1440
         TabIndex        =   605
         Top             =   300
         Width           =   2805
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes "
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
         Index           =   82
         Left            =   180
         TabIndex        =   604
         Top             =   1380
         Width           =   390
      End
   End
   Begin VB.Frame frameCCxCta 
      Height          =   6435
      Left            =   60
      TabIndex        =   370
      Top             =   20
      Visible         =   0   'False
      Width           =   5655
      Begin VB.OptionButton optCCxCta 
         Caption         =   "SIN cen. reparto"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   751
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optCCxCta 
         Caption         =   "Centros de reparto"
         Height          =   195
         Index           =   1
         Left            =   3720
         TabIndex        =   750
         Top             =   4260
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optCCxCta 
         Caption         =   "Todo"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   749
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CheckBox chkCC_Cta 
         Caption         =   "Ver movimientos posteriores"
         Height          =   195
         Left            =   2880
         TabIndex        =   398
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2520
         TabIndex        =   394
         Text            =   "Text5"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   2520
         TabIndex        =   393
         Text            =   "Text5"
         Top             =   1260
         Width           =   2535
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   15
         Left            =   1320
         TabIndex        =   372
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   14
         Left            =   1320
         TabIndex        =   371
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton cmdCCxCta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   380
         Top             =   5820
         Width           =   1035
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   10
         Left            =   2640
         TabIndex        =   378
         Text            =   "Text1"
         Top             =   4380
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   8
         ItemData        =   "frmListado.frx":2BE37
         Left            =   1320
         List            =   "frmListado.frx":2BE39
         Style           =   2  'Dropdown List
         TabIndex        =   375
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   9
         Left            =   2640
         TabIndex        =   376
         Text            =   "Text1"
         Top             =   3900
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   9
         ItemData        =   "frmListado.frx":2BE3B
         Left            =   1320
         List            =   "frmListado.frx":2BE3D
         Style           =   2  'Dropdown List
         TabIndex        =   377
         Top             =   4380
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   4440
         TabIndex        =   381
         Top             =   5820
         Width           =   975
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   2280
         TabIndex        =   383
         Text            =   "Text2"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   374
         Text            =   "Text2"
         Top             =   3000
         Width           =   795
      End
      Begin VB.TextBox txtDCost 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   2280
         TabIndex        =   382
         Text            =   "Text2"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtCCost 
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   373
         Text            =   "Text2"
         Top             =   2520
         Width           =   795
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   10
         ItemData        =   "frmListado.frx":2BE3F
         Left            =   1380
         List            =   "frmListado.frx":2BE41
         Style           =   2  'Dropdown List
         TabIndex        =   379
         Top             =   5160
         Width           =   1215
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
         Index           =   109
         Left            =   3720
         TabIndex        =   748
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   46
         Left            =   180
         TabIndex        =   397
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   41
         Left            =   480
         TabIndex        =   396
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   40
         Left            =   480
         TabIndex        =   395
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   15
         Left            =   1020
         Picture         =   "frmListado.frx":2BE43
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   14
         Left            =   1020
         Picture         =   "frmListado.frx":2C845
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fechas"
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
         Index           =   45
         Left            =   180
         TabIndex        =   392
         Top             =   3600
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   39
         Left            =   660
         TabIndex        =   391
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   38
         Left            =   660
         TabIndex        =   390
         Top             =   3960
         Width           =   615
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   5
         Left            =   1020
         Picture         =   "frmListado.frx":2D247
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   37
         Left            =   480
         TabIndex        =   389
         Top             =   3120
         Width           =   465
      End
      Begin VB.Image imgCCost 
         Height          =   240
         Index           =   4
         Left            =   1020
         Picture         =   "frmListado.frx":2DC49
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   36
         Left            =   480
         TabIndex        =   388
         Top             =   2580
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Centro de coste"
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
         Index           =   44
         Left            =   180
         TabIndex        =   387
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Centros de coste por cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   13
         Left            =   780
         TabIndex        =   386
         Top             =   360
         Width           =   4185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes cálculo"
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
         Index           =   43
         Left            =   240
         TabIndex        =   385
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label15"
         Height          =   315
         Index           =   26
         Left            =   180
         TabIndex        =   384
         Top             =   5940
         Width           =   2835
      End
   End
   Begin VB.Frame frameResumen 
      Height          =   6135
      Left            =   60
      TabIndex        =   400
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtNumRes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   406
         Text            =   "Text2"
         Top             =   2580
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   18
         Left            =   4200
         TabIndex        =   403
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   4980
         TabIndex        =   421
         Top             =   5460
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Height          =   1035
         Left            =   120
         TabIndex        =   419
         Top             =   4200
         Width           =   5865
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   410
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   1260
            TabIndex        =   411
            Top             =   240
            Width           =   1065
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   2340
            TabIndex        =   412
            Top             =   240
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   3540
            TabIndex        =   413
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   4680
            TabIndex        =   414
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   540
            TabIndex        =   415
            Top             =   720
            Width           =   1365
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   1860
            TabIndex        =   416
            Top             =   720
            Width           =   1125
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3240
            TabIndex        =   417
            Top             =   720
            Width           =   1125
         End
         Begin VB.CheckBox ChkNivelRes 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4440
            TabIndex        =   418
            Top             =   720
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdDiarioRes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3900
         TabIndex        =   420
         Top             =   5460
         Width           =   975
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   12
         ItemData        =   "frmListado.frx":2E64B
         Left            =   1380
         List            =   "frmListado.frx":2E64D
         Style           =   2  'Dropdown List
         TabIndex        =   404
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   12
         Left            =   2820
         TabIndex        =   405
         Text            =   "Text1"
         Top             =   1500
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   11
         ItemData        =   "frmListado.frx":2E64F
         Left            =   1380
         List            =   "frmListado.frx":2E651
         Style           =   2  'Dropdown List
         TabIndex        =   401
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   11
         Left            =   2820
         TabIndex        =   402
         Text            =   "Text1"
         Top             =   1020
         Width           =   855
      End
      Begin VB.TextBox txtNumRes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   407
         Text            =   "Text2"
         Top             =   2520
         Width           =   1035
      End
      Begin VB.TextBox txtNumRes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1440
         TabIndex        =   408
         Text            =   "Text2"
         Top             =   3600
         Width           =   1035
      End
      Begin VB.TextBox txtNumRes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4200
         TabIndex        =   409
         Text            =   "Text2"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   25
         Left            =   180
         TabIndex        =   432
         Top             =   5580
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número asiento"
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
         Index           =   54
         Left            =   180
         TabIndex        =   431
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   18
         Left            =   5520
         Picture         =   "frmListado.frx":2E653
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   53
         Left            =   4200
         TabIndex        =   430
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label17 
         Caption         =   "Diario resumen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   0
         Left            =   1680
         TabIndex        =   429
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   43
         Left            =   660
         TabIndex        =   428
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   42
         Left            =   660
         TabIndex        =   427
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   52
         Left            =   180
         TabIndex        =   426
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número página"
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
         Index           =   51
         Left            =   2820
         TabIndex        =   425
         Top             =   2220
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Acumulado "
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
         Index           =   50
         Left            =   180
         TabIndex        =   424
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Debe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   660
         TabIndex        =   423
         Top             =   3660
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Haber"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   48
         Left            =   3540
         TabIndex        =   422
         Top             =   3660
         Width           =   435
      End
   End
   Begin VB.Frame frameConsolidado 
      Height          =   6915
      Left            =   120
      TabIndex        =   519
      Top             =   0
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CheckBox chkDesgloseEmpresa 
         Caption         =   "Desglose por empresas"
         Height          =   255
         Left            =   240
         TabIndex        =   552
         Top             =   5640
         Width           =   5835
      End
      Begin MSComctlLib.ProgressBar pb9 
         Height          =   375
         Left            =   240
         TabIndex        =   550
         Top             =   6360
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   15
         Left            =   2880
         TabIndex        =   529
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   15
         ItemData        =   "frmListado.frx":2E6DE
         Left            =   1560
         List            =   "frmListado.frx":2E6E0
         Style           =   2  'Dropdown List
         TabIndex        =   528
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Index           =   14
         Left            =   2880
         TabIndex        =   527
         Text            =   "Text1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
         Height          =   315
         Index           =   14
         ItemData        =   "frmListado.frx":2E6E2
         Left            =   1560
         List            =   "frmListado.frx":2E6E4
         Style           =   2  'Dropdown List
         TabIndex        =   526
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2880
         TabIndex        =   543
         Text            =   "Text5"
         Top             =   1260
         Width           =   2895
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2880
         TabIndex        =   542
         Text            =   "Text5"
         Top             =   780
         Width           =   2895
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   19
         Left            =   1560
         TabIndex        =   525
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   18
         Left            =   1560
         TabIndex        =   524
         Top             =   780
         Width           =   1215
      End
      Begin VB.CommandButton cmdConsolidado 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   539
         Top             =   6360
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   24
         Left            =   5100
         TabIndex        =   540
         Top             =   6360
         Width           =   975
      End
      Begin VB.Frame Frame32 
         Height          =   1095
         Left            =   240
         TabIndex        =   523
         Top             =   4380
         Width           =   5865
         Begin VB.CheckBox ChkConso 
            Caption         =   "9º nivel"
            Height          =   210
            Index           =   9
            Left            =   4200
            TabIndex        =   538
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "8º nivel"
            Height          =   210
            Index           =   8
            Left            =   3120
            TabIndex        =   537
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "7º nivel"
            Height          =   210
            Index           =   7
            Left            =   1800
            TabIndex        =   536
            Top             =   720
            Width           =   1065
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "6º nivel"
            Height          =   210
            Index           =   6
            Left            =   420
            TabIndex        =   535
            Top             =   720
            Width           =   1185
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "5º nivel"
            Height          =   210
            Index           =   5
            Left            =   4680
            TabIndex        =   534
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "4º nivel"
            Height          =   210
            Index           =   4
            Left            =   3420
            TabIndex        =   533
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "3º nivel"
            Height          =   210
            Index           =   3
            Left            =   2280
            TabIndex        =   532
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "2º nivel"
            Height          =   210
            Index           =   2
            Left            =   1200
            TabIndex        =   531
            Top             =   240
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox ChkConso 
            Caption         =   "1er nivel"
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   530
            Top             =   240
            Value           =   1  'Checked
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "    Nivel     "
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
            Index           =   67
            Left            =   180
            TabIndex        =   541
            Top             =   0
            Width           =   810
         End
      End
      Begin VB.ListBox List4 
         Enabled         =   0   'False
         Height          =   1425
         ItemData        =   "frmListado.frx":2E6E6
         Left            =   1440
         List            =   "frmListado.frx":2E6E8
         TabIndex        =   520
         Top             =   2820
         Width           =   2775
      End
      Begin VB.Label Label21 
         Caption         =   "Label21"
         Height          =   255
         Left            =   240
         TabIndex        =   551
         Top             =   6060
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Index           =   74
         Left            =   240
         TabIndex        =   549
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   53
         Left            =   720
         TabIndex        =   548
         Top             =   2340
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   52
         Left            =   720
         TabIndex        =   547
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   73
         Left            =   240
         TabIndex        =   546
         Top             =   540
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   51
         Left            =   720
         TabIndex        =   545
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   50
         Left            =   720
         TabIndex        =   544
         Top             =   825
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   19
         Left            =   1320
         Picture         =   "frmListado.frx":2E6EA
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   18
         Left            =   1320
         Picture         =   "frmListado.frx":2F0EC
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Balance consolidado de empresas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   17
         Left            =   780
         TabIndex        =   522
         Top             =   180
         Width           =   5085
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Index           =   66
         Left            =   240
         TabIndex        =   521
         Top             =   2820
         Width           =   825
      End
      Begin VB.Image Image7 
         Height          =   240
         Left            =   1080
         Picture         =   "frmListado.frx":2FAEE
         Top             =   2820
         Width           =   240
      End
   End
   Begin VB.Frame frameListFacCli 
      Height          =   6975
      Left            =   120
      TabIndex        =   182
      Top             =   0
      Visible         =   0   'False
      Width           =   5355
      Begin VB.TextBox txtPag2 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   189
         Text            =   "Text1"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CheckBox ChkListFac 
         Caption         =   "Mostrar tipo IVA"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   626
         Top             =   6000
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkMostrarRetencion 
         Caption         =   "Mostrar retención"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   625
         Top             =   5994
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Frame FrameClientesCons 
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   2295
         Left            =   240
         TabIndex        =   744
         Top             =   840
         Width           =   4935
         Begin VB.ListBox List10 
            Enabled         =   0   'False
            Height          =   1620
            ItemData        =   "frmListado.frx":304F0
            Left            =   1320
            List            =   "frmListado.frx":304F2
            TabIndex        =   745
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Empresas"
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
            Index           =   107
            Left            =   0
            TabIndex        =   746
            Top             =   480
            Width           =   825
         End
         Begin VB.Image Image13 
            Height          =   240
            Left            =   960
            Picture         =   "frmListado.frx":304F4
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   31
         Left            =   3960
         TabIndex        =   193
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtNpag2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   192
         Text            =   "Text2"
         Top             =   4560
         Width           =   795
      End
      Begin VB.CheckBox ChkListFac 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   627
         Top             =   6000
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox ChkListFac 
         Caption         =   "Agrupar por cuenta"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   628
         Top             =   6480
         Width           =   2115
      End
      Begin VB.OptionButton optListFac 
         Caption         =   "Fecha liquidacion"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   629
         Top             =   5520
         Width           =   1695
      End
      Begin VB.OptionButton optListFac 
         Caption         =   "Fecha factura"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   624
         Top             =   5520
         Width           =   1515
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   360
         TabIndex        =   296
         Top             =   720
         Width           =   4635
         Begin VB.Label Label2 
            Caption         =   "Proveedores"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   240
            TabIndex        =   297
            Top             =   120
            Width           =   1950
         End
      End
      Begin VB.OptionButton optListFac 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   212
         Top             =   5160
         Width           =   1635
      End
      Begin VB.OptionButton optListFac 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   211
         Top             =   5160
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2400
         TabIndex        =   197
         Text            =   "Text5"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   196
         Text            =   "Text5"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   11
         Left            =   3960
         TabIndex        =   191
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   188
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   10
         Left            =   1200
         TabIndex        =   190
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   187
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdFactCli 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   194
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   3960
         TabIndex        =   195
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   183
         Text            =   "Text1"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   184
         Text            =   "Text1"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtNumFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   185
         Text            =   "Text1"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtNumFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   186
         Text            =   "Text1"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmListado.frx":30EF6
         ToolTipText     =   "Filtrar NIF"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N.I.F."
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
         Index           =   122
         Left            =   360
         TabIndex        =   843
         Top             =   3240
         Width           =   405
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   31
         Left            =   3720
         Picture         =   "frmListado.frx":318F8
         Top             =   4620
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha informe"
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
         Index           =   99
         Left            =   2400
         TabIndex        =   707
         Top             =   4620
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden"
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
         Index           =   98
         Left            =   360
         TabIndex        =   706
         Top             =   5160
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pág:"
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
         Index           =   97
         Left            =   360
         TabIndex        =   705
         Top             =   4560
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   9
         Left            =   960
         Picture         =   "frmListado.frx":31983
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   8
         Left            =   960
         Picture         =   "frmListado.frx":32385
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   20
         Left            =   360
         TabIndex        =   210
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
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
         Index           =   19
         Left            =   360
         TabIndex        =   209
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Listado facturas clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   5
         Left            =   120
         TabIndex        =   208
         Top             =   360
         Width           =   5175
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   10
         Left            =   960
         Picture         =   "frmListado.frx":32D87
         Top             =   3900
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   11
         Left            =   3720
         Picture         =   "frmListado.frx":32E12
         Top             =   3900
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   207
         Top             =   3945
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   360
         TabIndex        =   206
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   205
         Top             =   3945
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   360
         TabIndex        =   204
         Top             =   2445
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Index           =   18
         Left            =   360
         TabIndex        =   203
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   17
         Left            =   2040
         TabIndex        =   202
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   201
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   17
         Left            =   360
         TabIndex        =   200
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   199
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   198
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.Menu mnP1 
      Caption         =   "p1"
      Visible         =   0   'False
      Begin VB.Menu mnPrueba 
         Caption         =   "Prueba F1"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '1 .- Listado consultas extractos, listado de MAYOR
    '2 .- Listado de cuentas
    '3 .- Listado de asientos
    '4 .- Totales cuenta concepto
    '5 .- balance de sumas y saldos
    '6 .- Reemision de diario
    '7 .- Cuentas de explotacion
    '8 .- Listado facturas clientes
    '9 .- Presupuestos
    
    '10 .- Balance presupuestario
    '11 .- Certificado declaración de IVA
    '12 .- Liquidacion IVA
    '13 .- Listado facturas proveedores    LO CAMBIAMOS 19/FEB/2004
    '14 .- Libro diario oficial
    
    '       Centros de coste
    '15 .- Acumulados y saldos
    '16 .- Cuenta explotacion centro de coste
    '17 .- Centro de coste por cuenta
        
    '18 .- Diario resumen
    '19 .- Cta explotacion por cta
    
    '20 .- Modelo 347
    '21 .- Cuenta de explotacion comparativa
    
    '22 .- Borre facturas clientes
    '23 .- "         "    proveedores
    '24 .- Balance consolidado de empresas
    
    '25 .- Balances personalizados
    '26 .-   "          "           Perdeterminado Situacion
    '27 .-   "          "           Perdeterminado Py g
    
    
    '28 .- Modelo 349
    '29 .- Traspaso PERSA
    '30 .- Traspaso ACE
    
    '31 .- Cuenta explotacion CONSIOLIDADA
    
    
    '----------------------------------- Legalizacion de libros
    ' 32.- Diario Normal. Como el 14
    ' 33.- Diario resumen. Como el 18
    ' 34.- Consulta extracots
    ' 35.- Inventario inicial
    ' 36.- Balance sumas y saldos
    ' 37.- Listado facturas clientes
    ' 38.- Listado Facturas proveedores
    ' 39.- Balance pyG
    ' 40.- Balance Situacion
    ' 41.- Inventario final
    
    ' Antiguos 39 y 40
    ' 50 .- Balance perosnalizados, consolidados. PyG
    ' 51 .- "           "               "        SITUACION
    
    ' 52 .- Facturas proveedor Consolidadas
    ' 53 .- Facturas CLIENTES   consolidada

    ' 54 .- Evolucion mensual de saldos

    ' 55 .- Relacion de clientes por cuenta gastos/ventas
    ' 56 .-  "          proveedores   ""           ""

    ' 57 .- Copiar balance configurables
    ' 58 .- Modelo 340
    
Public EjerciciosCerrados As Boolean
    'En algunos informes me servira para utilizar unas tablas u otras
Public Legalizacion As String   'Datos para la legalizacion
    
Dim Tablas As String
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim Cont As Long
Dim I As Integer

Dim Importe As Currency

'Para los balcenes frameBalance
' Cuando este trbajando con cerrado
' Para poder sbaer cuando empezaba el año del ejercicio a listar
Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date


Dim HanPulsadoSalir As Boolean

'Para cancelar
Dim PulsadoCancelar As Boolean


'Private Sub PonFoco(ByRef T1 As TextBox)
'    T1.SelStart = 0
'    T1.SelLength = Len(T1.Text)
'End Sub

'Private Sub KEYPRESS(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
'End Sub




Private Sub AceptarCuentas_Click()
Dim J As Integer
'Imprimir el listado, segun sea
'    If txtCta(3).Text <> "" And txtCta(2).Text <> "" Then
'        If Val(txtCta(3).Text) > Val(txtCta(2).Text) Then
'            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarCuentas(3, 2) Then Exit Sub
    Screen.MousePointer = vbHourglass
    'Selecion cuenta desde hasta
    SQL = ""
    RC = "" 'Sera el nexo
    If txtCta(3).Text <> "" Then
        SQL = SQL & "codmacta >=""" & txtCta(3).Text & """"
        RC = " AND "
    End If
        
    If txtCta(2).Text <> "" Then
        SQL = SQL & RC & "codmacta<=""" & txtCta(2).Text & """"
        RC = " AND "
    End If
    
    If SQL <> "" Then
        Cad = "(" & SQL & ")"
    Else
        Cad = ""
    End If
    
    RC = ""
    If Option1(0).Value Then
        For I = 1 To Check1.Count - 2  'El 10 k es el ultimo nivel no lo quiero
            If Check1(I).visible Then
                If Check1(I).Value = 1 Then
                    SQL = ""
                    J = DigitosNivel(I)
                    For Cont = 1 To J
                        SQL = SQL & "_"
                    Next Cont
                    If RC <> "" Then RC = RC & " OR "
                    RC = RC & " (codmacta like '" & SQL & "')"
                End If
            End If
        Next I
        If Check1(10).Value Then
            If RC <> "" Then RC = RC & " OR "
            RC = RC & " apudirec='S'"
        End If
    End If
    If RC <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "(" & RC & ")"
    End If
    
    
    If Option1(1).Value Then
        RC = " apudirec='S' "
        If Me.cmbCuenta.ListIndex > 0 Then
            RC = RC & " AND model347 = " & Me.cmbCuenta.ItemData(Me.cmbCuenta.ListIndex)

        End If
        If RC <> "" Then
            If Cad <> "" Then Cad = Cad & " AND "
            Cad = Cad & RC
        End If
    End If
    If Not GenerarDatosCuentas(Cad) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If chkCuetas_x_nombre.Value = 1 Then
        J = 62  'Pq los listados ordenados por nombre son el 67 y el 68
    Else
        J = 0
    End If
    If Option1(0).Value Then
    '1ER informe
            frmImprimir.Opcion = 5 + J 'informe 1 Ctas
            If Combo2.ListIndex >= 0 Then
                frmImprimir.NumeroParametros = 1
                frmImprimir.OtrosParametros = "Resaltar= " & Combo2.ItemData(Combo2.ListIndex) & "|"
            Else
                frmImprimir.NumeroParametros = 0
            End If
        Else
            '2º informe
            frmImprimir.Opcion = 6 + J 'informe 1 Ctas ultmimo nu¡ivel
            frmImprimir.NumeroParametros = 0
    End If
    frmImprimir.FormulaSeleccion = "{ado.codusu}=  " & vUsu.Codigo
    frmImprimir.SoloImprimir = False
    frmImprimir.Show vbModal
End Sub



Private Sub AsientoAceptar_Click()
'Imprimir el listado, segun sea
    SQL = ""
    If txtAS(1).Text <> "" And txtAS(1).Text <> "" Then
        If Val(txtAS(0).Text) > Val(txtAS(1).Text) Then
            MsgBox "Asiento desde mayor que asiento hasta.", vbExclamation
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    Cad = ""
    If txtAS(0).Text <> "" Then
        SQL = "t1.numaspre>=" & txtAS(0).Text
        Cad = "Desde " & txtAS(0).Text & " - " & txtDesAs(0).Text
        I = 1
    End If
    If txtAS(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & "t1.numaspre <=" & txtAS(1).Text
        If Cad <> "" Then
            Cad = Cad & "    h"
        Else
            Cad = Cad & "H"
        End If
        Cad = Cad & "asta " & txtAS(1).Text & " - " & txtDesAs(1).Text
    End If
    If IAsientosPre(SQL) Then
        If Cad <> "" Then
            Cad = "DesdeHasta= """ & Cad & """|"
            I = 1
        Else
            I = 0
        End If
        
        frmImprimir.Opcion = 7
        frmImprimir.NumeroParametros = I
        frmImprimir.OtrosParametros = Cad
        frmImprimir.FormulaSeleccion = "{ado.codusu} = " & vUsu.Codigo
        frmImprimir.SoloImprimir = False
        frmImprimir.Show vbModal
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then HacerF1
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub




Private Sub Check2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then HacerF1
End Sub

Private Sub Check2_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then HacerF1
End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub Check4_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then HacerF1
End Sub

Private Sub chk347_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkAce_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub chkAgrupacionCtasBalance_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkApaisado_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkApertura_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub






Private Sub chkBalIncioEjercicio_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 112 Then HacerF1
End Sub

Private Sub chkBalIncioEjercicio_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkBalPerCompa_Click()
    FrameTapa2.visible = Me.chkBalPerCompa.Value = 0
End Sub

Private Sub chkBalPerCompaCon_Click()
    FrameTapa3.visible = Me.chkBalPerCompaCon.Value = 0
    cmbFecha(21).visible = Not FrameTapa3.visible
    txtAno(21).visible = Not FrameTapa3.visible
End Sub

Private Sub chkBalPerCompaCon_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkBorreFactura_Click()
    Me.FrBorrePorEjercicios.visible = Me.chkBorreFactura.Value = 1
End Sub

Private Sub chkCC_Cta_KeyPress(KeyAscii As Integer)
        ListadoKEYpress KeyAscii
End Sub



Private Sub chkCliproxCtalineas_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkcmp_Click(Index As Integer)
    If chkcmp(Index).Value = 1 Then
        For I = 1 To 10
            If I <> Index Then chkcmp(I).Value = 0
        Next I
    End If
End Sub

Private Sub chkcmp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub chkcmp_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub ChkConso_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkCopyBalan_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub





Private Sub chkCtaExpCC_Click(Index As Integer)
    If Index = 1 Then
         FrameCCComparativo.visible = chkCtaExpCC(1).Value = 1
    End If
End Sub

Private Sub chkCtaExpCC_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkCtaExpCon_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub chkCtaExpCon_KeyPress(Index As Integer, KeyAscii As Integer)
     ListadoKEYpress KeyAscii
End Sub

Private Sub chkCtaExplo_Click(Index As Integer)
    If chkCtaExplo(Index).Value = 1 Then
        For I = 1 To 10
            If I <> Index Then chkCtaExplo(I).Value = 0
        Next I
    End If
End Sub

Private Sub chkCtaExplo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = 112 Then HacerF1
End Sub

Private Sub chkCtaExplo_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub chkCtaExploC_Click(Index As Integer)
    If chkCtaExploC(Index).Value = 1 Then
        For I = 1 To 9
            If I <> Index Then chkCtaExploC(I).Value = 0
        Next I
    End If
End Sub

Private Sub chkCtaExploC_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub ChkCtaPre_Click(Index As Integer)
'    If ChkCtaPre(Index).Value = 1 Then
'        For I = 1 To 10
'            If I <> Index Then ChkCtaPre(I).Value = 0
'        Next I
'    End If
End Sub

Private Sub chkCuetas_x_nombre_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub chkDesgloseBasexCta_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkDesgloseEmpresa_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub ChkEvolSaldo_Click(Index As Integer)
    If ChkEvolSaldo(Index).Value = 1 Then
        For I = 1 To 10
            If I <> Index Then ChkEvolSaldo(I).Value = 0
        Next I
    End If

End Sub

Private Sub ChkEvolSaldo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub ChkEvolSaldo_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkExploCompa_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkExploCompaPorcentual_KeyPress(KeyAscii As Integer)
  ListadoKEYpress KeyAscii
End Sub

Private Sub chkExplotacion_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub chkIVAdetallado_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkLiqDefinitiva_KeyPress(KeyAscii As Integer)
  ListadoKEYpress KeyAscii
End Sub


Private Sub ChkListFac_KeyPress(Index As Integer, KeyAscii As Integer)
     ListadoKEYpress KeyAscii
End Sub

Private Sub chkMeses_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkMostrarRetencion_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkMovimientos_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub ChkNivelRes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub ChkNivelRes_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub





Private Sub chkPreAct_KeyPress(KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub chkPreMensual_KeyPress(KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub chkPresu3Digit_KeyPress(KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub chkQuitaCierre_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub chkRenumerar_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub chkResetea6y7_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub cmbFecha_Click(Index As Integer)
    If Not PrimeraVez Then
        If Index = 0 Then ComprobarFechasBalanceQuitar6y7
    End If
End Sub

Private Sub cmbFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub cmbFecha_KeyPress(Index As Integer, KeyAscii As Integer)

    ListadoKEYpress KeyAscii
    
End Sub



Private Sub cmd347_Click()
Dim B As Boolean
Dim B2 As Boolean

    If Text3(21).Text = "" Or Text3(22).Text = "" Then
        MsgBox "Introduce las fechas de consulta.", vbExclamation
        Exit Sub
    End If

    If Not ComprobarFechas(21, 22) Then Exit Sub
    
    
    If Year(CDate(Text3(21).Text)) <> Year(CDate(Text3(22).Text)) Then
        MsgBox "Esta abarcando dos años. Se considera el año: " & Year(CDate(Text3(22).Text)), vbExclamation
    End If
    If Combo5.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de informe.", vbExclamation
        Exit Sub
    End If
    
    
    If Combo5.ListIndex = 0 And Text347(0).Text = "" Then
            MsgBox "Escriba el nombre del responsable.", vbExclamation
            Exit Sub
    End If
            
            
    If Combo5.ListIndex = 2 Then
        MsgBox "Impresión del modelo oficial NO esta disponible", vbExclamation
        Exit Sub
    End If
    
    
    If Combo5.ListIndex >= 2 And vUsu.Nivel > 2 Then
        MsgBox "No tiene permiso", vbExclamation
        Exit Sub
    End If
    
    
    
    If Combo5.ListIndex = 3 Then
        'Enero 2012
        'Tiene que ser una año exacto
        If Month(CDate(Text3(21).Text)) <> 1 Or Month(CDate(Text3(21).Text)) <> 1 Then
            MsgBox "Año natural. Enero diciembre", vbExclamation
            Exit Sub
        End If
        If Month(CDate(Text3(22).Text)) <> 12 Or Day(CDate(Text3(22).Text)) <> 31 Then
            MsgBox "Año natural. Hasta 31 diciembre", vbExclamation
            Exit Sub
        End If
        
    End If
    
    If Text347(1).Text = "" Then
        MsgBox "Importe limite en blanco", vbExclamation
        Exit Sub
    End If
    
    If Me.chk347(2).Value = 0 And Me.chk347(3).Value = 0 Then
        MsgBox "Seleccione una procedencia de datos para el 347.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Modificacion de 26 Marzo 2007
    '------------------------------------
    'Hay una tabla auxiliar donde se guardan datos externos de 347.
    'Cuando voy a imprimir los datos pedire si de una y/o de la otra
    
    SQL = "DELETE FROM Usuarios.z347 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    SQL = "DELETE FROM Usuarios.z347trimestral where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    
    Set miRsAux = New ADODB.Recordset
    
    'El de siempre
    If Me.chk347(2).Value = 1 Then
        B = ComprobarCuentas347_
        If Not B Then Exit Sub
    End If
    
    
    
    
    'Los datos externos de la tabla externa de 347.  Estos importes van todos
    If Me.chk347(3).Value = 1 Then
        B = ComprobarCuentas347DatosExternos
        If Not B Then Exit Sub
    End If

    
    'Cobros efectivo
    'Updatearemos a cero los metalicos que no llegen al minimo
    SQL = "Select ImporteMaxEfec340 from parametros "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = DBLet(miRsAux!ImporteMaxEfec340, "N")
    miRsAux.Close
    If Val(SQL) > 0 Then
        SQL = "UPDATE usuarios.z347trimestral set metalico=0  WHERE codusu = " & vUsu.Codigo & " AND metalico < " & TransformaComasPuntos(SQL)
         Conn.Execute SQL
    End If
     
     
    'Ahora borramos todas las entrdas k no superan el importe limite
    Label2(31).Caption = "Comprobar importes"
    Label2(31).Refresh
    Importe = ImporteFormateado(Text347(1).Text)
    SQL = "Delete from Usuarios.z347 where codusu = " & vUsu.Codigo & " AND Importe  <" & TransformaComasPuntos(CStr(Importe))
    Conn.Execute SQL
    
    
    
    
    'Comprobare si hay datos
    'Comprobamos si hay datos
    SQL = "Select count(*) FROM Usuarios.z347 where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            Cont = miRsAux.Fields(0)
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If Cont = 0 Then
        MsgBox "No se ha devuelto ningun dato", vbExclamation
        Exit Sub
    End If
    
    'Precomprobacion de NIFs
    If Not ComprobarNifs347 Then Exit Sub
    
    
    Label2(31).Caption = ""
    Label2(30).Caption = ""
    DoEvents
    Screen.MousePointer = vbDefault
    

    
    If B Then
            Select Case Combo5.ListIndex
            Case 0
                'La carta
                Cad = "¿ Desea imprimir también los proveedores ?"
                If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then
                    'Cad = " AND {ado1.cliprov} = 0"
                    'ahora
                    Cad = " AND {ado1.cliprov} = " & Asc(0)
                Else
                    Cad = ""
                End If
                CargaEncabezadoCarta 0, Text347(0).Text
                SQL = "Preimpreso= " & Abs(chk347(1).Value) & "|"
                With frmImprimir
                    .OtrosParametros = SQL
                    .NumeroParametros = 1
                    SQL = "{ado.codusu}=" & vUsu.Codigo
                    SQL = SQL & Cad
                    .FormulaSeleccion = SQL
                    .SoloImprimir = False
                    'Opcion dependera del combo
                    .Opcion = 43
                    .Show vbModal
                End With
                
                
            Case 2, 3
                'Si es impresion y el numero de registros es superior a 25 no
                'puede imprimirse
                Cont = 0
                SQL = ""
                If Combo5.ListIndex = 2 Then
                    Set RS = New ADODB.Recordset
                    SQL = "Select count(*) from usuarios.z347 where codusu =" & vUsu.Codigo
                    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not RS.EOF Then Cont = DBLet(RS.Fields(0), "N")
                    RS.Close
                    Set RS = Nothing
                    If Cont > 25 Then
                        SQL = "El numero de registros supera los 25." & vbCrLf & _
                            "Se debe presentar por diskette o via internet."
                        MsgBox SQL, vbExclamation
                        Exit Sub
                    End If
                End If
                
                'Modelo de hacienda
                B2 = Modelo347(Combo5.ListIndex = 2, Year(CDate(Text3(22).Text)))
                If B2 And (Combo5.ListIndex = 3) Then CopiarFicheroHacienda (True)
            Case Else
            
                'LISTADO
                '-----------------------------------------------------------------
                If chk347(3).Value = 1 Then
                    'Volcar datos sobre tabla ztmpctaexplo
                    'para que asi salgan en columnas
                    If Not Volcar347TablaTmp2 Then Exit Sub
                End If
                    
            
            
                
                SQL = "Desde " & Text3(21).Text & "      hasta  " & Text3(22).Text
                SQL = "Fechas= """ & SQL & """|"
                'Imprime un listadito
                With frmImprimir
                    .OtrosParametros = SQL
                    .NumeroParametros = 1
                    .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                    .SoloImprimir = False
                    'Opcion dependera del combo
                    If chk347(3).Value = 1 Then
                        .Opcion = 80
                    Else
                        .Opcion = 42
                    End If
                    .Show vbModal
                End With
            End Select
            
        
    End If
    
End Sub



Private Sub cmd349_Click()
Dim B As Boolean
Dim ConCli As Integer 'Clientes
Dim ConPro As Integer  'proveedores
Dim Periodo As String

    If Text3(26).Text = "" Or Text3(27).Text = "" Then
        MsgBox "Introduce las fechas de consulta.", vbExclamation
        Exit Sub
    End If

    If Not ComprobarFechas(26, 27) Then Exit Sub


    If chk349.Value = 1 Then
        'Si ha marcado la presentacion, la fecha la necesito. Para coger el año desde ahi
        If Text3(35).Text = "" Then
            MsgBox "Escriba la fecha de presentación", vbExclamation
            Exit Sub
        End If
        
        If Combo6.ListIndex < 0 Then
            MsgBox "Indique periodo", vbExclamation
            Exit Sub
        End If
        
        
        If vUsu.Nivel > 2 Then
            MsgBox "No tiene permiso", vbExclamation
            Exit Sub
        End If
  
        
        
    End If


    Screen.MousePointer = vbHourglass
    B = ComprobarCuentas349(ConCli, ConPro)
    Screen.MousePointer = vbDefault
    If Not vParam.Presentacion349Mensual Then
        
        'Trimestral
        If Combo6.ListIndex < 4 Then
            RC = CStr(Combo6.ListIndex + 1) & "T"
        Else
            RC = "0A"
        End If
    Else
        'Mensual
        If Combo6.ListIndex < 12 Then
            RC = Format(Combo6.ListIndex + 1, "00")
        Else
            RC = "0A"
        End If
    
    End If
    Periodo = RC
    If B Then
    

    
    
        'Comprobamos si va mas de una empresa
        Cad = vEmpresa.nomempre
        If List5.ListCount > 1 Then vEmpresa.nomempre = "CONSOLIDADO"
            
        
    
        
        
        'Desde hastas Abril 2012
        RC = "Fechas: " & Text3(26).Text & " - " & Text3(27).Text
        RC = RC & "       Periodo: " & Combo6.Text
        RC = "pdh1= """ & RC & """|"
        
        'Las que habian
        RC = "ContadorLinCli= " & ConCli & "|ContadorLinPRO= " & ConPro & "|" & RC
        With frmImprimir
            .OtrosParametros = RC
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .Opcion = 56
            .Show vbModal
        End With
        vEmpresa.nomempre = Cad
    
    
        'Impresion del modelo oficial
        If chk349.Value = 1 Then
            RC = Periodo
            If MODELO349(True, RC, CDate(Text3(35).Text)) Then CopiarFicheroHacienda (False)               'Modelo de hacienda
        End If
    End If
End Sub

Private Sub cmdACE_Click()
Dim B As Boolean
    Text1(1).Text = Trim(Text1(1).Text)
    If Text1(1).Text = "" Then
        MsgBox "Escriba el BM Code", vbExclamation
        Exit Sub
    End If
    If List6.ListCount = 0 Then
        MsgBox "Seleccione al menos una empresa.", vbExclamation
        Exit Sub
    End If
    
    If cmbFecha(18).ListIndex < 0 Then
        MsgBox "Seleccione el mes del cálculo.", vbExclamation
        Exit Sub
    End If
    
    
    
    'falta meter en cont el numero de digitos
    
    Cont = 0
    Cad = ""
    For I = 1 To chkAce.Count
        If chkAce(I).visible Then
            If chkAce(I).Value = 1 Then
                'Cont = I  'Para saber k nivel es el seleccionado
                If I = 10 Then
                    Cont = 10
                Else
                    Cont = DigitosNivel(I)
                End If
                Cad = Cad & "1"
            End If
        End If
    Next I
    
    If Len(Cad) <> 1 Then
        MsgBox "Seleccione uno, y solo uno de los niveles contables.", vbExclamation
        Exit Sub
    End If
    
    'El numero de digitos si ha seleccionado el nivel 10 esta en digitosultimonivle
    If Cont = 10 Then Cont = vEmpresa.DigitosUltimoNivel
    
    'Llegados aqui ya podemos realizar el traspaso
    Screen.MousePointer = vbHourglass
    
    'Segun sea actual o siguiente
    If optAce(0).Value Then
        I = 0
    Else
        I = 1
    End If
    FechaIncioEjercicio = DateAdd("yyyy", I, vParam.fechaini)
    FechaFinEjercicio = DateAdd("yyyy", I, vParam.fechafin)
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.ztmppresu1 WHERE codusu = " & vUsu.Codigo
    'INSERT INTO ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (1, 0, NULL, NULL, 0, 0, NULL)
    FijarValoresACE Text1(1).Text, Me.cmbFecha(18).ListIndex + 1, FechaIncioEjercicio, FechaFinEjercicio, optAceAcum(1).Value, CInt(Cont)
    B = True
    For I = 0 To List6.ListCount - 1
        'Generaremos para la empresa tal
         B = GenerarACE2(List6.ItemData(I))
         If Not B Then Exit For
    Next I
    If B Then
       If GeneraFicheroAce Then
        CopiarACE
        Unload Me
    End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAceptarHco_Click()
    If Text3(4).Text = "" Or Text3(5).Text = "" Then
        MsgBox "Introduce las fechas de consulta.", vbExclamation
        Exit Sub
    End If
'    If Text3(4).Text <> "" And Text3(5).Text <> "" Then
'        If CDate(Text3(4).Text) > CDate(Text3(5).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(4, 5) Then Exit Sub

    'Llegados aqui montamos las cadenas
    SQL = ""
    Tablas = ""
    If EjerciciosCerrados Then Tablas = "1"
    
    'Fechas intervalor
    SQL = "Fechas= ""Desde " & Text3(4).Text & " hasta " & Text3(5).Text & """|"
    Cad = "(hcabapu" & Tablas & ".fechaent>= '" & Format(Text3(4).Text, FormatoFecha) & "')"
    Cad = Cad & " AND (hcabapu" & Tablas & ".fechaent<= '" & Format(Text3(5).Text, FormatoFecha) & "')"
    'Formula de seleccion
    RC = ""
    If Me.txtAsiento(0).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "hcabapu" & Tablas & ".numasien >=" & txtAsiento(0).Text
        
        RC = "Desde asiento nº: " & txtAsiento(0).Text
    End If
    
    If Me.txtAsiento(1).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "hcabapu" & Tablas & ".numasien <=" & txtAsiento(1).Text
        
        If RC <> "" Then
            RC = RC & "   h"
        Else
            RC = "H"
        End If
        RC = RC & "asta asiento nº: " & txtAsiento(1).Text
    End If
    
    
    If Me.txtDiario(0).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "hcabapu" & Tablas & ".numdiari >=" & txtDiario(0).Text
        If RC <> "" Then RC = RC & "    "
        RC = RC & "Desde diario nº: " & txtDiario(0).Text
    End If
    
    If Me.txtDiario(1).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "hcabapu" & Tablas & ".numdiari <=" & txtDiario(1).Text
        If RC <> "" Then
            RC = RC & "    h"
            Else
            RC = "H"
        End If
        RC = RC & "asta diario nº: " & txtDiario(1).Text
    End If
    
    SQL = SQL & "Cuenta= """ & RC & """|"
    'El titulo
    If txtReemisionDiario.Text = "" Then
        RC = "REEMISION DEL DIARIO"
    Else
        RC = txtReemisionDiario.Text
    End If
    SQL = SQL & "Titulo= """ & RC & """|"
    
    'Fecha impresion
    If Text3(8).Text = "" Then Text3(8).Text = Format(Now, "dd/mm/yyyy")
    SQL = SQL & "FechaIMP= """ & Text3(8).Text & """|"
    
    Screen.MousePointer = vbHourglass
    If chkTotalAsiento.Value Then
        I = 12
    Else
        I = 65
    End If
    If IHcoApuntes(Cad, Tablas) Then
        If Opcion <> 35 Then
                With frmImprimir
                    .OtrosParametros = SQL
                    .NumeroParametros = 4
                    .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                    .SoloImprimir = False
                    'Opcion dependera del combo
                    .Opcion = I
                    .Show vbModal
                End With
        Else
              GeneraLegalizaPRF SQL, 4
              CadenaDesdeOtroForm = "OK"
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAceptarTotalesCta_Click()
Dim Aux As String
    'Para enviar a totales
    'Ciertas comprobaciones
'    If txtCta(5).Text <> "" And txtCta(4).Text <> "" Then
'        If Val(txtCta(5).Text) > Val(txtCta(4).Text) Then
'            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarCuentas(5, 4) Then Exit Sub

    If Text3(6).Text = "" Or Text3(3).Text = "" Then
        MsgBox "Introduce las fechas de consulta.", vbExclamation
        Exit Sub
    End If
'    If Text3(6).Text <> "" And Text3(3).Text <> "" Then
'        If CDate(Text3(6).Text) > CDate(Text3(3).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(6, 3) Then Exit Sub
    If txtConcepto.Text = "" Then
        SQL = "No ha puesto concepto. Si continua mostrará todos los conceptos(Puede llevar mucho tiempo)"
        SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If

    SQL = ""
    
    Aux = "hlinapu"
    If EjerciciosCerrados Then Aux = Aux & "1"
    
    
    'Er conceto
    If txtConcepto.Text = "" Then
        SQL = SQL & " 1 = 1"
        RC = "Concepto: TODOS" & "    "
    Else
        SQL = SQL & "codconce = " & txtConcepto.Text
        RC = "Concepto: " & txtConcepto.Text & "   " & Me.txtDescConce.Text & "    "
    End If
    txtCta(5).Tag = ""
    If txtCta(5).Text <> "" Then
        SQL = SQL & " AND " & Aux & ".codmacta>=""" & txtCta(5).Text & """"
        txtCta(5).Tag = txtCta(5).Tag & " desde " & txtCta(5).Text
    End If
    If txtCta(4).Text <> "" Then
        SQL = SQL & " AND " & Aux & ".codmacta <=""" & txtCta(4).Text & """"
        txtCta(5).Tag = txtCta(5).Tag & " hasta " & txtCta(4).Text
    End If

    Text3(6).Tag = "Fechas: "
    'Fecha desde
    SQL = SQL & " AND fechaent >= '" & Format(Text3(6).Text, FormatoFecha) & "'"
    Text3(6).Tag = Text3(6).Tag & " desde " & Text3(6).Text
    'hasta
    SQL = SQL & " AND  fechaent <= '" & Format(Text3(3).Text, FormatoFecha) & "'"
    Text3(6).Tag = Text3(6).Tag & " hasta " & Text3(3).Text
    
    'Numero de parep¡metros: 2
    Aux = "Cuenta= """ & txtCta(5).Tag & """|"
    Text3(6).Tag = RC & "   " & Text3(6).Tag
        
    Aux = Aux & "Fechas= """ & Text3(6).Tag & """|"
    
    Tablas = "hlinapu"
    If EjerciciosCerrados Then Tablas = Tablas & "1"
    
    
    Screen.MousePointer = vbHourglass
    If ITotalesCtaConcepto(SQL, Tablas) Then
        With frmImprimir
            .OtrosParametros = Aux
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            If txtConcepto.Text = "" Then
                'TOTAL
                If chkMeses.Value = 0 Then
                    .Opcion = 74
                Else
                    .Opcion = 75
                End If
            Else
                If chkMeses.Value = 0 Then
                    .Opcion = 13
                Else
                    .Opcion = 14
                End If
            End If
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBalance_Click()
'Ciertas comprobaciones
'   If txtCta(6).Text <> "" And txtCta(7).Text <> "" Then
'        If txtCta(6).Text > txtCta(7).Text Then
'            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
'            Exit Sub
'        End If
'    End If

    Cont = 0
    For I = 1 To 10
        If Check2(I).Value = 1 Then Cont = Cont + 1
    Next I
    If Cont = 0 Then
        MsgBox "Seleccione como mínimo un nivel contable", vbExclamation
        Exit Sub
    End If



    'Febrero 2009
    '-----------------------------------------------------
    'Balance a inicio de ejerecicio
    If Me.chkBalIncioEjercicio.Value = 1 Then
        'Es un balance especial.
        Screen.MousePointer = vbHourglass
        HacerBalanceInicio
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If Not ComprobarCuentas(6, 7) Then Exit Sub


    'Ahora, si ha puesto desde hasta cuenta, no puede seleccionar
    'un desde.
    If Me.txtCta(6).Text <> "" Or txtCta(7).Text <> "" Then
        If Cont > 1 Then
            If vUsu.Nivel < 2 Then
                Cad = "debe"
            Else
                Cad = "puede"
            End If
            Cad = "No " & Cad & " pedir un balance a distintos niveles poniendo desde/hasta cuenta"
            MsgBox Cad, vbExclamation
            If vUsu.Nivel > 1 Then Exit Sub
        End If
    End If


    If txtAno(0).Text = "" Or txtAno(1).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    
    If Not ComparaFechasCombos(0, 1, 0, 1) Then Exit Sub
'    If txtAno(0).Text <> "" And txtAno(1).Text <> "" Then
'        If Val(txtAno(0).Text) > Val(txtAno(1).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        Else
'            If Val(txtAno(0).Text) = Val(txtAno(1).Text) Then
'                If Me.cmbFecha(0).ListIndex > Me.cmbFecha(1).ListIndex Then
'                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
'
    
    If Abs(Val(txtAno(1).Text) - Val(txtAno(0).Text)) > 2 Then
        MsgBox "Fechas pertenecen a ejercicios distintos.", vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    


    'Trabajaresmos contra ejercicios cerrados
    'Si el mes es mayor o igual k el de inicio, significa k la feha
    'de inicio de aquel ejercicio fue la misma k ahora pero de aquel año
    'si no significa k fue la misma de ahora pero del año anterior
    I = cmbFecha(0).ListIndex + 1
    If I >= Month(vParam.fechaini) Then
        Cont = Val(txtAno(0).Text)
    Else
        Cont = Val(txtAno(0).Text) - 1
    End If
    Cad = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Cont
    FechaIncioEjercicio = CDate(Cad)
    
    I = cmbFecha(1).ListIndex + 1
    If I <= Month(vParam.fechafin) Then
        Cont = Val(txtAno(1).Text)
    Else
        Cont = Val(txtAno(1).Text) + 1
    End If
    Cad = Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & Cont
    FechaFinEjercicio = CDate(Cad)

    
    'Veamos si pertenecen a un mismo año
    If Abs(DateDiff("d", FechaFinEjercicio, FechaIncioEjercicio)) > 365 Then
        MsgBox "Las fechas son incorrectas. Abarca mas de un ejercicio", vbExclamation
        Exit Sub
    End If
 
 
 
 
 
    'Si tiene otro grupo de
    
    If Me.chkResetea6y7.Value = 1 Then
       If Val(txtAno(0).Text) > Year(vParam.fechaini) Then
            If vParam.grupoord <> "" And vParam.Automocion <> "" Then
                If Check2(1).Value = 1 Or Check2(2).Value = 1 Then
                    MsgBox "La cuenta de exclusion del grupoord de la analitica no esta inlcuida en el balance", vbExclamation
                End If
            End If
        End If
    End If
    'Fecha informe
    If Text3(7).Text = "" Then Text3(7).Text = Format(Now, "dd/mm/yyyy")

    EmpezarBalance -1, pb2  'La -1 es balance normal
End Sub

Private Sub cmdBalances_Click()

    'Comprobamos datos
    If Me.txtNumBal(0).Text = "" Then
        MsgBox "Número de balance incorrecto", vbExclamation
        Exit Sub
    End If
    
    
    'Año 1
    If txtAno(16).Text = "" Then
        MsgBox "Año no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    If Val(txtAno(16).Text) < 1900 Then
        MsgBox "No se permiten años anteriores a 1900", vbExclamation
        Exit Sub
    End If
    
    If chkBalPerCompa.Value = 1 Then
        If txtAno(17).Text = "" Then
            MsgBox "Año no puede estar en blanco", vbExclamation
            Exit Sub
        End If
        If Val(txtAno(17).Text) < 1900 Then
            MsgBox "No se permiten años anteriores a 1900", vbExclamation
            Exit Sub
        End If
    End If

    'Fecha informe
    If Text3(25).Text = "" Then
        MsgBox "Fecha informe incorrecta.", vbExclamation
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    I = -1
    If chkBalPerCompa.Value = 1 Then
        I = Val(cmbFecha(17).ListIndex)
        I = I + 1
        If I = 0 Then I = -1
    End If
    GeneraDatosBalanceConfigurable CInt(txtNumBal(0).Text), Me.cmbFecha(16).ListIndex + 1, CInt(txtAno(16).Text), I, Val(txtAno(17).Text), False, -1
    
    
    'Para saber k informe abriresmos
    Cont = 1
    RC = 1 'Perdidas y ganancias
    Set RS = New ADODB.Recordset
    SQL = "Select * from sbalan where numbalan=" & Me.txtNumBal(0).Text
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then

            If DBLet(RS!Aparece, "N") = 0 Then
                Cont = 3
            Else
                Cont = 1
            End If

        RC = RS!perdidas
    End If
    RS.Close
    Set RS = Nothing
        
        
    'Si es comarativo o no
    If Me.chkBalPerCompa.Value = 1 Then Cont = Cont + 1
        
    'Textos
    RC = "perdidasyganancias= " & RC & "|"
          
    SQL = RC & "FechaImp= """ & Text3(25).Text & """|"
    SQL = SQL & "Titulo= """ & Me.TextDescBalance(0).Text & """|"
    'PGC 2008 SOlo pone el año, NO el mes
    If vParam.NuevoPlanContable Then
        RC = ""
    Else
        RC = cmbFecha(16).List(cmbFecha(16).ListIndex)
    End If
    RC = RC & " " & txtAno(16).Text
    RC = "fec1= """ & RC & """|"
    SQL = SQL & RC
    
    
    
    
    If Me.chkBalPerCompa.Value = 1 Then
            'PGC 2008 SOlo pone el año, NO el mes
            If vParam.NuevoPlanContable Then
                RC = ""
            Else
                RC = cmbFecha(17).List(cmbFecha(17).ListIndex)
            End If
            RC = RC & " " & txtAno(17).Text
            RC = "Fec2= """ & RC & """|"
            SQL = SQL & RC
            

    Else
        'Pong el nombre del mes
        RC = UCase(Mid(cmbFecha(16).Text, 1, 1)) & Mid(cmbFecha(16).Text, 2, 2)
        RC = "vMes= """ & RC & """|"
        SQL = SQL & RC
    End If
    SQL = SQL & "Titulo= """ & Me.TextDescBalance(0).Text & """|"
        
    If Opcion < 39 Or Opcion > 40 Then
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 4
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            
            'La opcion sera si esta marcado apaisado
            If chkApaisado.Value = 1 Then
                .Opcion = 82 + Cont   'El 83 es el primero en la de apisado que es para el PGC2008
            Else
                .Opcion = 48 + Cont   'El 49 es el primero de los rpt de balance
            End If
            .Show vbModal
        End With
    Else
        GeneraLegalizaPRF SQL, 6
        CadenaDesdeOtroForm = "OK"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBalanPersConso_Click()
Dim J As Integer
Dim vE As Cempresa
Dim EstaLaEmpresaActual As Boolean
Dim Contabilidades As String
    
 'Comprobamos datos
    If Me.txtNumBal(1).Text = "" Then
        MsgBox "Número de balance incorrecto", vbExclamation
        Exit Sub
    End If
    
    
    'Año 1
    If txtAno(20).Text = "" Then
        MsgBox "Año no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    If Val(txtAno(20).Text) < 1900 Then
        MsgBox "No se permiten años anteriores a 1900", vbExclamation
        Exit Sub
    End If
    
    If chkBalPerCompaCon.Value = 1 Then
        If txtAno(21).Text = "" Then
            MsgBox "Año no puede estar en blanco", vbExclamation
            Exit Sub
        End If
        If Val(txtAno(21).Text) < 1900 Then
            MsgBox "No se permiten años anteriores a 1900", vbExclamation
            Exit Sub
        End If
    End If

    'Fecha informe
    If Text3(34).Text = "" Then
        MsgBox "Fecha informe incorrecta.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    For J = 0 To List8.ListCount - 1
        Cad = Cad & "1"
    Next J
    
    If Cad = "" Then
        MsgBox "Seleccione alguna empresa para el balance", vbExclamation
        Exit Sub
    End If
    
    'AHORA HAY K VER si tienen la misma
    '     Fechas de ejercicio
    '     Mismo numero de nivles contables
    '
    '
    
    
    'Primero veremos si esta la empresa actual
    Cad = ""
    For J = 0 To List8.ListCount - 1
        If List8.ItemData(J) = vEmpresa.codempre Then
            Cad = "SI"
            Exit For
        End If
    Next J
    
    Set vE = New Cempresa
    If Cad <> "" Then
        'Tenemos la empresa actual seleccionada
        'Con lo cual la asignamos
        Set vE = vEmpresa
        Cad = ""
        EstaLaEmpresaActual = True
        FechaIncioEjercicio = vParam.fechaini
        
        
    Else
        EstaLaEmpresaActual = False
        'Cojemos la primera empresa y asignamos los valores
        J = List8.ItemData(0)
        If vE.Leer(CStr(J)) = 1 Then
            Cad = "Error leyendo datos empresa: " & List8.List(0)
            MsgBox Cad, vbExclamation
        End If
        
        'Ahora pongo la fecha inicio
        SQL = "Select * from Conta" & List8.ItemData(0) & ".parametros"
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        FechaIncioEjercicio = "0:00:00"
        If Not RS.EOF Then
            FechaIncioEjercicio = DBLet(RS!fechaini, "F")
        End If
        RS.Close
        Set RS = Nothing
    End If
    If Cad <> "" Then Exit Sub
    
    For J = 0 To List8.ListCount - 1
        If List8.ItemData(J) <> vE.codempre Then
            'Solo comparamos si la empresa no es
            'la que hemos puesto en VE
            If Not CompararEmpresasBlancePerson(List8.ItemData(J), vE, FechaIncioEjercicio) Then
                Cad = List8.List(J) & vbCrLf & Cad
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
        End If
    Next J
    
    
    
    I = -1
    If chkBalPerCompa.Value = 1 Then
        I = Val(cmbFecha(21).ListIndex)
        I = I + 1
        If I = 0 Then I = -1
    End If
    
    
    Screen.MousePointer = vbHourglass
    'Cargamos los datos de la empresa
    Cad = ""
    For J = 0 To List8.ListCount - 1
        Cad = Cad & List8.ItemData(J) & "|"
    Next J
    GeneraDatosBalanceConfigurable CInt(txtNumBal(1).Text), Me.cmbFecha(20).ListIndex + 1, CInt(txtAno(20).Text), I, Val(txtAno(21).Text), False, Cad

    
    
    Cont = 1
    RC = 1 'Perdidas y ganancias
    Set RS = New ADODB.Recordset
    SQL = "Select * from sbalan where numbalan=" & Me.txtNumBal(1).Text
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS!Aparece, "N") = 0 Then
            Cont = 3
        Else
            Cont = 1
        End If
        RC = RS!perdidas
    End If
    RS.Close
    Set RS = Nothing
        
        
    'Si es comarativo o no
    If Me.chkBalPerCompaCon.Value = 1 Then Cont = Cont + 1
        
    'Textos
    RC = "perdidasyganancias= " & RC & "|"
    SQL = RC & "FechaImp= """ & Text3(34).Text & """|"
    SQL = SQL & "Titulo= """ & Me.TextDescBalance(1).Text & """|"
    
    If Val(txtAno(20).Text) < 2008 Then
        MsgBox "  Error. Avise soporte tecnico. CLAVE: Cambio", vbExclamation
        RC = cmbFecha(20).List(cmbFecha(20).ListIndex) & " " & txtAno(20).Text
        RC = "fec1= """ & RC & """|"
    Else
        'fec1= " 2011"|vMes= "Dic"
        RC = "27/" & cmbFecha(20).ListIndex + 1 & "/2000"
        RC = Format(RC, "mmm")
        
        RC = "fec1= """ & txtAno(20).Text & """|" & "vMes= """ & RC & """|"
        
    End If
    SQL = SQL & RC
    If Me.chkBalPerCompaCon.Value = 1 Then
            RC = cmbFecha(21).List(cmbFecha(21).ListIndex) & " " & txtAno(21).Text
            RC = "Fec2= """ & RC & """|"
            SQL = SQL & RC
    End If
    'SQL = SQL & "Titulo= """ & Me.TextDescBalance(0).Text & """|"
        
    RC = vEmpresa.nomempre
    vEmpresa.nomempre = "CONSOLIDADO:"
    With frmImprimir
        .OtrosParametros = SQL
        .NumeroParametros = 3
        .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
        .SoloImprimir = False
        'Opcion dependera del combo
        .Opcion = 48 + Cont   'El 49 es el primero de los rpt de balance
        .Show vbModal
    End With
    vEmpresa.nomempre = RC
    Screen.MousePointer = vbDefault
    
    
    
    
    
    
    
    
    'Mostrar listado
    With frmImprimir
        .OtrosParametros = ""
        .NumeroParametros = 0
        .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
        .SoloImprimir = False
        'Opcion dependera del combo
        .Opcion = 49   'El 49 es el primero de los rpt de balance
        .Show vbModal
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBalPre_Click()

    If Not ComprobarCuentas(12, 13) Then Exit Sub
    
    SQL = ""
    For I = 1 To Me.ChkCtaPre.Count
        If Me.ChkCtaPre(I).Value Then SQL = SQL & "&"
    Next I
    If Len(SQL) <> 1 Then
        If chkPreMensual.Value = 1 Then
            MsgBox "Seleccione uno, y solo uno, de los niveles contables.", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If txtMes(2).Text <> "" And Me.chkPreMensual.Value = 0 Then
        
        MsgBox "Si indica el mes debe marcar la opcion ""mensual""", vbExclamation
        Exit Sub
    End If
    
    If txtMes(2).Text <> "" Then
        If Val(txtMes(2).Text) < 1 Or Val(txtMes(2).Text) > 12 Then
            MsgBox "Mes incorrecto: " & txtMes(2).Text, vbExclamation
            Exit Sub
        End If
    End If
    
    
    'Solo podemos quitar el asiento de apertura para ejercicio actual
    I = 0
    If chkQuitarApertura.Value = 1 Then
        I = 1
        'ejer siguiente
        If chkPreAct.Value = 1 Then
            I = 0
        Else
            'Si es mensual y el mes NO es uno tampoco lo quita
            If chkPreMensual.Value = 1 Then
                If Val(txtMes(2).Text) > 1 Then I = 0
            End If
        End If
    End If
    chkQuitarApertura.Value = I
        
    
    
    
    SQL = ""
    RC = ""
    If txtCta(12).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        RC = "Desde " & txtCta(12).Text & " - " & DtxtCta(12).Text
        SQL = SQL & "presupuestos.codmacta >= '" & txtCta(12).Text & "'"
    End If
    
    
    If txtCta(13).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        If RC <> "" Then
            RC = RC & "       h"
        Else
            RC = "H"
        End If
        RC = RC & "asta " & txtCta(13).Text & " - " & DtxtCta(13).Text
        SQL = SQL & "presupuestos.codmacta <= '" & txtCta(13).Text & "'"
    End If

    If SQL <> "" Then SQL = SQL & " AND"
    I = Year(vParam.fechaini)
    If chkPreAct.Value Then I = I + 1
    SQL = SQL & " anopresu =" & I
    
    
    If RC <> "" Then RC = """ + chr(13) +""" & RC
    If chkPreMensual.Value = 1 Then
        If txtMes(2).Text <> "" Then RC = "** " & Format("01/" & txtMes(2).Text & "/1999", "mmmm") & " ** " & RC
        RC = "  MENSUAL " & RC
    End If
    
    
    
    RC = "Año: " & I & RC
    CadenaDesdeOtroForm = ""
    
    For Cont = 1 To 10
        If ChkCtaPre(Cont).Value = 1 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "- " & Cont
    Next

    RC = RC & " Digitos: " & Mid(CadenaDesdeOtroForm, 2)
    
    If chkQuitarApertura.Value = 1 Then RC = RC & "     Sin apertura"
    CadenaDesdeOtroForm = "CampoSeleccion= """ & RC & """|"

    RC = ""
    For Cont = 1 To 9
        If ChkCtaPre(Cont).Value = 1 Then
            If RC = "" Then RC = Cont
        End If
    Next
    If RC = "" Then RC = "11"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Remarcar= " & RC & "|"
    


    If GeneraBalancePresupuestario() Then
        With frmImprimir
            .OtrosParametros = CadenaDesdeOtroForm
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .Opcion = 23 + Me.chkPreMensual.Value
            .Show vbModal
        End With
    End If
    pb4.visible = False
End Sub

Private Sub cmdBorrarFacCli_Click()

    If Me.chkBorreFactura.Value = 0 Then
        If Text3(24).Text = "" Then
            MsgBox "Fechas 'HASTA' es obligada.", vbExclamation
            Exit Sub
        End If
        If Not ComprobarFechas(23, 24) Then Exit Sub
        If txtNumFac(0).Text <> "" And txtNumFac(1).Text <> "" Then
            If Val(txtNumFac(0).Text) > Val(txtNumFac(1).Text) Then
                MsgBox "Número factura inicio mayor que fin", vbExclamation
                Exit Sub
            End If
        End If
    
        If txtSerie(2).Text <> "" And txtSerie(3).Text <> "" Then
            If txtSerie(2).Text > txtSerie(3).Text Then
                MsgBox "Serie desde mayor que seria hasta", vbExclamation
                Exit Sub
            End If
        End If
        
        'Fecha
        FechaIncioEjercicio = CDate(Text3(24).Text)
        If vParam.fechaini < FechaIncioEjercicio Then
            MsgBox "La fecha final es posterior al inicio de ejercicio actual", vbExclamation
            Exit Sub
        End If
        
        
        'Fecha hasta
        If DateDiff("yyyy", CDate(Text3(24).Text), vParam.fechaini) < 5 Then
            Cad = "Debe conservar, al menos, los ultimos 5 años." & vbCrLf & vbCrLf
            Cad = Cad & "¿Desea continuar igualmente?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
        End If
    Else
        If Me.Label3(75).Tag = "" Then
            MsgBox "No es posible eliminar datos de las facturas", vbExclamation
            Exit Sub
        End If
  
        'Por ejercicios
        If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
            MsgBox "En años partidos, no puede efectuar el borre directo", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    
    
    'pASSSWORD MOMENTANEO
    Cad = InputBox("Escriba password de seguridad", "CLAVE")
    If LCase(Cad) <> "ariadna" Then
        If Cad <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    If Me.chkBorreFactura.Value = 0 Then
        HacerBorreFacturas
    Else
        HacerBorreFacturasEjercicio
    End If
    pb8.visible = False
    Screen.MousePointer = vbDefault
    Unload Me


End Sub

Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub cmdCanListExtr_Click(Index As Integer)
    If Me.cmdCancelarAccion.visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub

Private Sub cmdCCxCta_Click()

    '// Centros de coste por cuenta de explotacion

    If txtCCost(4).Text <> "" And txtCCost(5).Text <> "" Then
        If txtCCost(5).Text > txtCCost(5).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtAno(9).Text = "" Or txtAno(10).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    
    If txtAno(9).Text <> "" And txtAno(10).Text <> "" Then
        If Val(txtAno(9).Text) > Val(txtAno(10).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Sub
        Else
            If Val(txtAno(9).Text) = Val(txtAno(10).Text) Then
                If Me.cmbFecha(8).ListIndex > Me.cmbFecha(9).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    End If

    
    If Me.cmbFecha(10).ListIndex < 0 Then
        MsgBox "Seleccione un mes de cálculo", vbExclamation
        Exit Sub
    End If
    
    
    'Comprobamos que el total de meses no supera el año
    I = Val(txtAno(9).Text)
    Cont = Val(txtAno(10).Text)
    Cont = Cont - I
    I = 0
    If Cont > 1 Then
       I = 1  'Ponemos a uno para luego salir del bucle
    Else
        If Cont = 1 Then
            'Se diferencian un año, luego el mes fin tienes k ser menor a mes inicio
            If Me.cmbFecha(9).ListIndex >= Me.cmbFecha(8).ListIndex Then I = 1
        End If
    End If
    If I <> 0 Then
        MsgBox "El intervalo tiene que ser de un año como máximo", vbExclamation
        Exit Sub
    End If
    
    
    
    Screen.MousePointer = vbHourglass
    If GeneraCCxCtaExplotacion Then
        
        Label2(26).Caption = ""
        'Vamos a poner los textos
        SQL = "Mes cálculo: " & UCase(cmbFecha(10).List(cmbFecha(10).ListIndex))
        SQL = SQL & "   Desde : " & cmbFecha(8).ListIndex + 1 & " / " & txtAno(9).Text
        SQL = SQL & "   Hasta : " & cmbFecha(9).ListIndex + 1 & " / " & txtAno(10).Text
        
        
        Cad = ""
        If txtCta(14).Text <> "" Then Cad = "Desde cta:" & txtCta(14).Text
        If txtCta(15).Text <> "" Then
            If Cad <> "" Then Cad = Cad & "    "
            Cad = Cad & "Hasta cta: " & txtCta(15).Text
        End If
        If Cad <> "" Then SQL = SQL & "  " & Cad
        
        
        
        
        RC = ""
        'Centros de coste
        If Me.txtCCost(4).Text <> "" Then _
            RC = "Desde CC: " & Me.txtCCost(4).Text & " - " & Me.txtDCost(4).Text
        If Me.txtCCost(5).Text <> "" Then
            If RC <> "" Then RC = RC & "  "
            RC = RC & "Hasta CC: " & Me.txtCCost(5).Text & " - " & Me.txtDCost(5).Text
        End If
        
        
        If Me.chkCC_Cta.Value = 1 Then
            'Solo hay una linea
            I = 0
            Cont = 1
            If RC <> "" Then SQL = SQL & "     " & RC
            RC = ""
        Else
            'Hay dos lineas para poner todo
            I = 1
            Cont = 0
        End If
        
        
        'Si se ha marcado solo repartos lo marco en el informe
        If Me.optCCxCta(1).Value Then SQL = SQL & "   (C. REPARTO)"
        RC = "Fechas= """ & RC & """|"
        SQL = "Cuenta= """ & SQL & """|"
        SQL = SQL & RC
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = I + 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 36 + Cont
            .Show vbModal
        End With
    End If
    Label15.Caption = ""
    Label2(26).Caption = ""
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdCertIVA_Click()
Dim J As Integer

    If Text3(12).Text = "" Then
        MsgBox "Fecha informe no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    
'    If Text3(13).Text <> "" And Text3(14).Text <> "" Then
'        If CDate(Text3(13).Text) > CDate(Text3(14).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(13, 14) Then Exit Sub
    
    If Me.List1.ListCount = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Sub
    End If
    
    If Combo4.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de IVA", vbExclamation
        Exit Sub
    End If
    
    'Borramos lo k haya
    Screen.MousePointer = vbHourglass
    Conn.Execute "DELETE FROM Usuarios.zcertifiva where codusu =" & vUsu.Codigo
    
    'Empezamos
    Frame4.visible = True
    J = 0
    Cont = 1
    Do
        Label11.Caption = List1.List(J)
        Label11.Refresh
        Cad = "conta" & List1.ItemData(J)
        ProcesaRegstrosParaBD Cad, CInt(Combo4.ItemData(Combo4.ListIndex))
        J = J + 1
    Loop Until J >= List1.ListCount
    Frame4.visible = False
    Me.Refresh
    
    'Ponemos todos los parametros
    SQL = PonerDatosEmpresa
'
    
    If Cont = 1 Then
        MsgBox "Ningun dato se ha generado", vbExclamation
    Else
        'Ahora mostramos el informe
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 7
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .Opcion = 29
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdComparativo_Click()
Dim F As Date
    If txtAno(13).Text = "" Then
        MsgBox "Introduce el año de  consulta", vbExclamation
        Exit Sub
    End If
    
    If Me.cmbFecha(13).ListIndex < 0 Then
        MsgBox "Seleccione un mes de cálculo", vbExclamation
        Exit Sub
    End If
    
    FechaFinEjercicio = "25/" & cmbFecha(13).ListIndex + 1 & "/" & txtAno(13).Text
    F = UltimaFechaHcoCabapu
    SQL = ""
    If EjerciciosCerrados Then
        If vParam.fechaini < FechaFinEjercicio Then SQL = "Fecha pertenece a ejercicicio actual o siguiente"
    Else
        If vParam.fechaini > FechaFinEjercicio Then SQL = "Fecha pertenece a ejercicios cerrados"
    End If
'    If SQL <> "" Then
'        MsgBox SQL, vbExclamation
'        Exit Sub
'    End If
'
    SQL = ""
    For I = 1 To Me.chkcmp.Count
        If Me.chkcmp(I).Value Then
            SQL = SQL & "&"
            Cont = I
        End If
    Next I
    If Len(SQL) <> 1 Then
        MsgBox "Seleccione uno, y solo uno, de los niveles contables.", vbExclamation
        Exit Sub
    End If
    'Si keremos ultimo nivel
    If Cont = 10 Then Cont = -1
    
    Screen.MousePointer = vbHourglass
    
    'Antes del 4 de Mayo de 2005
    'Cambiamos el dato de ultimafecha hlinapu
    'If GeneraCtaExplComparativa(cmbFecha(13).ListIndex + 1, (Me.chkExploCompa.Value = 0),  CInt(txtAno(13).Text), FechaIncioEjercicio, CInt(Cont)) Then
    If GeneraCtaExplComparativa(cmbFecha(13).ListIndex + 1, (Me.chkExploCompa.Value = 0), CInt(txtAno(13).Text), F, CInt(Cont)) Then
        'Generamos la tabla para imprimir
        If ImprimirCtaExploComp Then
            I = Val(txtAno(13).Text)
            SQL = "Anyo1= """ & I - 1 & """|"
            SQL = SQL & "Anyo2= """ & I & """|"
            'Periodo
            RC = UCase(cmbFecha(13).List(cmbFecha(13).ListIndex))
            If Me.chkExploCompa.Value = 0 Then RC = "Acumulado hasta " & RC
            SQL = SQL & "Periodo= """ & RC & """|"
            'Impresion
            With frmImprimir
                .OtrosParametros = SQL
                .NumeroParametros = 3
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = 44 + chkExploCompaPorcentual.Value
                .Show vbModal
            End With
        Else
            MsgBox "Ningún registro a mostrar", vbExclamation
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdConsolidado_Click()
Dim I As Integer
     If Not ComprobarCuentas(18, 19) Then Exit Sub
     If Not ComparaFechasCombos(14, 15, 14, 15) Then Exit Sub
     
     'Comprobar un nivel solicitado
     Cont = 0
     For I = 1 To 9
        If ChkConso(I).visible Then
            If ChkConso(I).Value = 1 Then Cont = Cont + 1
        End If
     Next I
     If Cont = 0 Then
        MsgBox "Debe seleccionar, por lo menos, un nivel contable,", vbExclamation
        Exit Sub
    End If
    
    
    If Abs(Val(txtAno(7).Text) - Val(txtAno(8).Text)) > 2 Then
        MsgBox "Fechas pertenecen a ejercicios distintos.", vbExclamation
        Exit Sub
    End If
 
    'Trabajaresmos contra ejercicios cerrados
    'Si el mes es mayor o igual k el de inicio, significa k la feha
    'de inicio de aquel ejercicio fue la misma k ahora pero de aquel año
    'si no significa k fue la misma de ahora pero del año anterior
    I = cmbFecha(14).ListIndex + 1
    If I >= Month(vParam.fechaini) Then
        Cont = Val(txtAno(14).Text)
    Else
        Cont = Val(txtAno(14).Text) - 1
    End If
    Cad = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Cont
    FechaIncioEjercicio = CDate(Cad)
    
    I = cmbFecha(15).ListIndex + 1
    If I <= Month(vParam.fechafin) Then
        Cont = Val(txtAno(15).Text)
    Else
        Cont = Val(txtAno(15).Text) + 1
    End If
    Cad = Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & Cont
    FechaFinEjercicio = CDate(Cad)

    
    
    
    'Veamos si pertenecen a un mismo año
    If Abs(DateDiff("d", FechaFinEjercicio, FechaIncioEjercicio)) > 365 Then
        MsgBox "Las fechas son incorrectas. Abarca mas de un ejercicio", vbExclamation
        Exit Sub
    End If
    
    
    'Comprobar que para los niveles señalados, para la empresa
    If Not ComprobarNivelesEmpresa Then Exit Sub
    Screen.MousePointer = vbHourglass
    'Llegados aqui, haremos la entrada
    'Para cade empresa cogeremos los datos
    Label21.visible = True
    Label21.Caption = "Borrando datos anteriores"
    Label21.Refresh
    SQL = "DELETE FROM Usuarios.ztmpbalanceconsolidado where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    'El progress
    pb9.Value = 0
    pb9.visible = True
    PulsadoCancelar = False
    'Para cada empresa seleccionada
    For I = 0 To List4.ListCount - 1
        pb9.Value = 0
        Label21.Caption = List4.List(I)
        Me.Refresh
        EmpezarBalance List4.ItemData(I), pb9
        If PulsadoCancelar Then
            Label21.Caption = ""
            Exit Sub
        End If
    Next I
    
    Set RS = New ADODB.Recordset
    
    'Hacer rectificado de cuentas a no consolidar
    Label21.Caption = "Exclusion"
    Label21.Refresh
    SQL = "Select count(codmacta) from ctaexclusion"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cont = 0
    If Not RS.EOF Then Cont = DBLet(RS.Fields(0), "N")
    RS.Close
    
    If Cont > 0 Then
        pb9.Value = 0
        pb9.visible = True
        Me.Refresh
        SQL = "Select codmacta from ctaexclusion"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        I = 1
        While Not RS.EOF
            pb9 = CInt((I / Cont) * 1000)
            SQL = "DELETE FROM Usuarios.ztmpbalanceconsolidado where codusu = " & vUsu.Codigo
            SQL = SQL & " AND cta = '" & RS!codmacta & "'"
            Conn.Execute SQL
            espera 0.25
            RS.MoveNext
            I = I + 1
        Wend
        RS.Close
    End If
    
    'Vamos acabando antes de mostrar el informe
    Label21.visible = False
    pb9.visible = False
    
    'Comprobamos si hay datos
    
    SQL = "SELECT codusu FROM Usuarios.ztmpbalanceconsolidado where codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    If Not RS.EOF Then I = 1
    RS.Close
    Set RS = Nothing
    If I > 0 Then
        'Vemos las empresas
        SQL = "Empresas= ""Empresas:"
        Cont = 0
        Do
            SQL = SQL & """ + chr(13) + ""        " & List4.List(Cont)
            Cont = Cont + 1
        Loop Until Cont = List4.ListCount
        SQL = SQL & """"
        With frmImprimir
            .OtrosParametros = "Cuenta= """"|" & SQL & "|"
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 46 + chkDesgloseEmpresa.Value
            .Show vbModal
        End With
    Else
        MsgBox "No hay ningún dato a visualizar", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Function ComprobarNivelesEmpresa() As Boolean
    ComprobarNivelesEmpresa = True
    If List4.ListCount = 1 Then Exit Function
    
    Set RS = New ADODB.Recordset
    For I = 0 To List4.ListCount - 1
        SQL = "Select * from Conta" & List4.ItemData(I) & ".empresa"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        If Not RS.EOF Then
            For Cont = 4 To RS.Fields.Count - 2  'Menos dos pq empieza en 0 y el ultimo nivel no cuenta
                Cad = Cad & DBLet(RS.Fields(Cont), "N")
            Next Cont
        End If
        'Ponemos a 0 el digito correspondiente al ultimo nivel
        Cont = RS.Fields(3)
        Cad = Mid(Cad, 1, Cont - 1) & "0" & Mid(Cad, Cont + 1)
        
        RS.Close
        
        If I = 0 Then
            'La primera asignamos cad a rc, para comparar en el futuro
            RC = Cad
            Tablas = List4.List(0)
        Else
            'Comparamos la cadena con RC, si es igual, son iguales los niveles
            If RC <> Cad Then
                MsgBox "La empresa " & List4.List(I) & " NO tiene los mismos digitos de nivel que la empresa " & Tablas, vbExclamation
                ComprobarNivelesEmpresa = False
                Exit For
            End If
        End If
    Next I
    Set RS = Nothing
End Function

Private Sub cmdCopyBalan_Click()
    If txtNumBal(3).Text = "" Then
        MsgBox "Seleccione el balance origen", vbExclamation
        Exit Sub
    End If
    
    
    Cad = "Va a copiar los datos del balance: " & vbCrLf & vbCrLf
    Cad = Cad & txtNumBal(3).Text & " - " & Me.TextDescBalance(3).Text & vbCrLf
    Cad = Cad & " sobre " & vbCrLf
    Cad = Cad & txtNumBal(2).Text & " - " & Me.TextDescBalance(2).Text & vbCrLf
    Cad = Cad & vbCrLf & vbCrLf & "Los datos del balance destino seran eliminados"
    Cad = Cad & vbCrLf & vbCrLf & "¿Desea continuar?"
    If MsgBox(Cad, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    
    SQL = "aparece"
    Cad = DevuelveDesdeBD("perdidas", "sbalan", "numbalan", txtNumBal(3).Text, "N", SQL)
    If Cad = "" Then
        MsgBox "Error leyendo datos: " & txtNumBal(3).Text
        Exit Sub
    End If
    
    Cad = "UPDATE sbalan SET perdidas=" & Cad & ",Aparece= " & SQL & " WHERE numbalan=" & txtNumBal(2).Text
    Conn.Execute Cad
    
    Cad = "DELETE FROM sperdi2 WHERE numbalan=" & txtNumBal(2).Text
    Conn.Execute Cad
    Cad = "DELETE FROM sperdid WHERE numbalan=" & txtNumBal(2).Text
    Conn.Execute Cad
    Cad = "INSERT INTO sperdid (NumBalan, Pasivo, codigo, padre, Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita, A_Cero, Pintar, LibroCD)"
    Cad = Cad & " SELECT " & txtNumBal(2).Text & ", Pasivo, codigo, padre, Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita, A_Cero, Pintar, LibroCD FROM"
    Cad = Cad & " sperdid WHERE numbalan = " & txtNumBal(3).Text
    Conn.Execute Cad
    
    
    If Me.chkCopyBalan.Value = 1 Then
        'COpio los datos tb
       ' NumBalan, Pasivo, codigo, codmacta, tipsaldo, Resta
        Cad = "INSERT INTO sperdi2 ( NumBalan, Pasivo, codigo, codmacta, tipsaldo, Resta)"
        Cad = Cad & " SELECT " & txtNumBal(2).Text & ", Pasivo, codigo, codmacta, tipsaldo, Resta FROM"
        Cad = Cad & " sperdi2 WHERE numbalan = " & txtNumBal(3).Text
        Conn.Execute Cad
    End If
    Unload Me
End Sub

Private Sub cmdCtaExpCC_Click()

    If txtCCost(2).Text <> "" And txtCCost(3).Text <> "" Then
        If txtCCost(2).Text > txtCCost(3).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtAno(7).Text = "" Or txtAno(8).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    
    If Me.cmbFecha(7).ListIndex < 0 Then
        MsgBox "Seleccione un mes de cálculo", vbExclamation
        Exit Sub
    End If
    
    If Not ComparaFechasCombos(7, 8, 5, 6) Then Exit Sub
     
    
    'Comprobamos que el total de meses no supera el año
    I = Val(txtAno(7).Text)
    Cont = Val(txtAno(8).Text)
    Cont = Cont - I
    I = 0
    If Cont > 1 Then
       I = 1  'Ponemos a uno para luego salir del bucle
    Else
        If Cont = 1 Then
            'Se diferencian un año, luego el mes fin tienes k ser menor a mes inicio
            If Me.cmbFecha(6).ListIndex >= Me.cmbFecha(5).ListIndex Then I = 1
        End If
    End If
    If I <> 0 Then
        MsgBox "El intervalo tiene que ser de un año como máximo", vbExclamation
        Exit Sub
    End If


    'No puede pedir movimientos posteriores y comparativo
    If chkCtaExpCC(0).Value = 1 And chkCtaExpCC(1).Value = 1 Then
        MsgBox "No puede pedir comparativo y movimientos posteriores", vbExclamation
        Exit Sub
    End If


    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    If GeneraCtaExplotacionCC Then
        Label15.Caption = ""
        'Vamos a poner los textos
        If chkCtaExpCC(1).Value = 1 And optCCComparativo(0).Value Then
            SQL = ""
        Else
            SQL = "Mes cálculo: " & UCase(cmbFecha(7).List(cmbFecha(7).ListIndex)) & " "
        End If
        
        If chkCtaExpCC(2).Value = 1 Then SQL = Trim(SQL & "   Solo reparto") & "  "
        
        'If Not (chkCCComparativo.Value = 1 And optCCComparativo(1).Value) Then
        SQL = SQL & "Desde : " & cmbFecha(5).ListIndex + 1 & " / " & txtAno(7).Text
        SQL = SQL & " hasta : " & cmbFecha(6).ListIndex + 1 & " / " & txtAno(8).Text
        
        
        Cad = ""
        'Si han puesto desde hasta cuenta
        If txtCta(29).Text <> "" Then Cad = " Desde cta: " & txtCta(29).Text ' & " " & Mid(DtxtCta(29).Text, 1, 13) & "..."
        If txtCta(30).Text <> "" Then Cad = Cad & " hasta cta: " & txtCta(30).Text ' & " " & Mid(DtxtCta(30).Text, 1, 13) & "..."
        Cad = Trim(Cad)
    
    

  
        
        If Me.chkCtaExpCC(0).Value = 1 Then
            'Solo hay una linea
            RC = ""
            I = 0
            If Me.txtCCost(2).Text <> "" Then _
                SQL = SQL & "Desde CC: " & Me.txtCCost(2).Text & " - " & Me.txtDCost(2).Text
            If Me.txtCCost(3).Text <> "" Then _
                SQL = SQL & " Hasta CC: " & Me.txtCCost(3).Text & " - " & Me.txtDCost(3).Text
                
                
            'Cont = 1
            Cont = 35
        Else
        
                'Hay dos lineas para poner todo
                I = 1
                RC = ""
                If Me.txtCCost(2).Text <> "" Then _
                    RC = " Desde CC: " & Me.txtCCost(2).Text & " - " & Me.txtDCost(2).Text
                If Me.txtCCost(3).Text <> "" Then _
                    RC = RC & " Hasta CC: " & Me.txtCCost(3).Text & " - " & Me.txtDCost(3).Text
                'Cont = 0
                Cont = 34
                If chkCtaExpCC(1).Value = 1 Then   '2013  Octubre 28.  Habia chkCtaExpCC(0)
                    'Comparativo
                    Cont = 90
                    If optCCComparativo(0).Value Then Cont = Cont + 1
                End If

        End If
        SQL = Trim(SQL & "    " & Cad)
        RC = "Fechas= """ & RC & """|"
        SQL = "Cuenta= """ & SQL & """|"
        SQL = SQL & RC
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = I + 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            '.Opcion = 34 + Cont
            .Opcion = Cont
            .Show vbModal
        End With
    End If
    Label15.Caption = ""
    Set miRsAux = Nothing
    Set RS = Nothing
    Screen.MousePointer = vbDefault


End Sub





Private Sub cmdCtaexplcmp_Click()
Dim C As Integer
Dim Dig As Integer

    'Mes de cálculo
    If cmbFecha(19).ListIndex < 0 Then
        MsgBox "Seleccion un mes para el cálculo.", vbExclamation
        Exit Sub
    End If
    
    ' Uno y solo uno de los niveles tiene que estar marcado
    Cad = ""
    For I = 1 To 9
        If Me.chkCtaExploC(I).visible Then
            If Me.chkCtaExploC(I).Value = 1 Then
                Cad = Cad & "1"
                Dig = DigitosNivel(I)
            End If
        End If
    Next I
    If Len(Cad) <> 1 Then
        MsgBox "Seleccione uno(y solo uno) de los niveles para el informe.", vbExclamation
        Exit Sub
    End If
    
    If txtAno(18).Text = "" Then
        MsgBox "Ponga el año para el listado.", vbExclamation
        Exit Sub
    End If
        
    'Nº empresas
    If List7.ListCount = 0 Then
        MsgBox "Seleccione al menos una empresa.", vbExclamation
        Exit Sub
    End If
   
    'Comprobamos datos
    If Text3(28).Text = "" Then Text3(28).Text = Format(Now, "dd/mm/yyyy")
    
    'Lo primero k haremos sera borrar los datos
    Screen.MousePointer = vbHourglass
    Conn.Execute "Delete FROM Usuarios.ztmpctaexplotacionC where codusu= " & vUsu.Codigo
    Label2(28).Caption = ""
    Label2(28).visible = True
    For C = 0 To List7.ListCount - 1
        Label2(28).Caption = List7.List(C)
        Label2(28).Refresh
        
        'Generaremos para la empresa tal los datos
        ListadoExplotacion Dig, List7.ItemData(C)
         
        'Una vez creados los datos insertaremos en la tmpde consolidado
        'INSERT INTO ztmpctaexplotacionc (codusu, cta, codempre, empresa, nomcta, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES (1, '1', 1, NULL, '0', NULL, NULL, NULL, NULL, NULL, NULL)
        SQL = "INSERT INTO Usuarios.ztmpctaexplotacionc (codusu, cta, codempre, empresa, nomcta, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) "
        SQL = SQL & " SELECT " & vUsu.Codigo & ", cta," & List7.ItemData(C) & ",'" & List7.List(C) & "', nomcta, acumAntD, acumAntH, acumPerD, acumPerH,"
        SQL = SQL & "TotalD, TotalH FROM Usuarios.ztmpctaexplotacion where codusu =" & vUsu.Codigo
        Conn.Execute SQL
    Next C
    
    Label2(28).visible = False
    pb10.visible = False
    Me.Refresh
    'Si tiene mas de unregistro mostraremos
    'Comprobamos k existen los registros
    Set RS = Nothing
    Set RS = New ADODB.Recordset
    SQL = "Select count(*) from Usuarios.ztmpctaexplotacionc WHERE codusu =" & vUsu.Codigo
    I = 0
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    Set RS = Nothing
    If I > 0 Then
            'Con movimientos
            Cont = 2 * Abs(chkCtaExpCon(0).Value)
            '0,1, con desglose
            Cont = Cont + Abs(chkCtaExpCon(1).Value)
            Cont = Cont + 59
            
            
            
            'Numero de parámetros: 2
            Cad = "Cuenta= ""Mes cálculo:  " & UCase(cmbFecha(19).List(cmbFecha(19).ListIndex)) & "    Año: " & txtAno(18).Text & """|"
            Cad = Cad & "FechaIMP= """ & Text3(28).Text & """|"
            'Vemos las empresas
            SQL = "Empresas= ""Empresas:"
            I = 0
            Do
                SQL = SQL & """ + chr(13) + ""        " & List7.List(I)
                I = I + 1
            Loop Until I = List7.ListCount
            SQL = SQL & """"
            'Los dos juntos
            SQL = Cad & SQL & "|"
            With frmImprimir
                .OtrosParametros = SQL
                .NumeroParametros = 3
                .FormulaSeleccion = "{ado.codusu}=  " & vUsu.Codigo
                .SoloImprimir = False
                'Opcion dependera del combo
                .Opcion = Cont
                .Show vbModal
            End With

    
    
    
    Else
        MsgBox "No se ha genereado ningun dato", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCtaExplo_Click()

'Mes de cálculo
If cmbFecha(2).ListIndex < 0 Then
    MsgBox "Seleccion un mes para el cálculo.", vbExclamation
    Exit Sub
End If

' Uno y solo uno de los niveles tiene que estar marcado
Cad = ""
For I = 1 To 10
    If Me.chkCtaExplo(I).visible Then
        If Me.chkCtaExplo(I).Value = 1 Then
            If I < 10 Then
                Cont = DigitosNivel(I)
            Else
                Cont = vEmpresa.DigitosUltimoNivel
            End If
            Cad = Cad & "1"
        End If
    End If
Next I
If Len(Cad) <> 1 Then
    MsgBox "Seleccione uno(y solo uno) de los niveles para el informe.", vbExclamation
    Exit Sub
End If

If txtAno(4).Text = "" Then
    MsgBox "Ponga el año para el listado.", vbExclamation
    Exit Sub
End If

    If vParam.grupoord <> "" And vParam.Automocion <> "" Then
        If CDate("01/" & cmbFecha(2).ListIndex + 1 & "/" & txtAno(4).Text) > vParam.fechafin Then
            'Ha seleccionado a uno o dos digitos
            If chkCtaExplo(1).Value = 1 Or chkCtaExplo(2).Value = 1 Then
                MsgBox "La cuenta de exclusion del grupoord de la analitica no esta inlcuida en el balance", vbExclamation
            End If
        End If
    End If
    'Comprobamos datos
    If Text3(9).Text = "" Then Text3(9).Text = Format(Now, "dd/mm/yyyy")

    Screen.MousePointer = vbHourglass
    If ListadoExplotacion(CInt(Cont)) Then

        If chkExplotacion.Value = 1 Then
            Cont = 19
        Else
            Cont = 20
        End If
        'Aqui mostraremos los informes
        pb3.visible = False
        'Numero de parámetros: 2
        SQL = "Cuenta= ""Mes cálculo:  " & UCase(cmbFecha(2).List(cmbFecha(2).ListIndex)) & "    Año: " & txtAno(4).Text & """|"
        SQL = SQL & "FechaIMP= """ & Text3(9).Text & """|"
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=  " & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .Opcion = Cont
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCtapoCC_Click()
Dim F As Date
'    If txtCta(16).Text <> "" And txtCta(17).Text <> "" Then
'        If Val(txtCta(16).Text) > Val(txtCta(17).Text) Then
'            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarCuentas(16, 17) Then Exit Sub

    If txtCCost(6).Text <> "" And txtCCost(7).Text <> "" Then
        If txtCCost(6).Text > txtCCost(7).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    If Not (Text3(19).Text <> "" And Text3(20).Text <> "") Then
        MsgBox "Debe introducir las fechas.", vbExclamation
        Exit Sub
    End If
'
'    If Text3(19).Text <> "" And Text3(20).Text <> "" Then
'        If CDate(Text3(19).Text) > CDate(Text3(20).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(19, 20) Then Exit Sub
    
    '-------------------------------------------------
    'INtervalo coja un año
    'Veamos siocupa mas de un año
    If Abs(DateDiff("d", CDate(Text3(19).Text), CDate(Text3(20).Text))) > 365 Then
        MsgBox "Las fechas son incorrectas. Abarca mas de un ejercicio", vbExclamation
        Exit Sub
    End If
    
    
    'Vamos a ver si coje un mismo año contable
    'Para ello situamos las fechas de inicio y fin de ejercicio
    'en funcion de la primera fecha
    F = CDate(Text3(19).Text)
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Años naturales
        FechaIncioEjercicio = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Year(F))
        FechaFinEjercicio = CDate(Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & Year(F))
        Else
            'Años partidos
            'vemos si la donde entra le fecha de inicio
            'Auxiliarmente usamos este var
            FechaFinEjercicio = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Year(F))
            I = Year(F)
                            'Es del años siguiente
            If F > FechaFinEjercicio Then I = I + 1
            
            'Ahora fijamos la de fin de jercicio y inicio
            FechaIncioEjercicio = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & I)
            FechaFinEjercicio = CDate(Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & I + 1)
    End If
    
    'Como era en funcion de la fecha de incio, comprobaremos la fecha fin
    F = CDate(Text3(20).Text)
    If F > FechaFinEjercicio Then
        MsgBox "Las fechas no estan dentro del mismo ejercicio contable", vbExclamation
        Exit Sub
    End If
    
    
    'Vemos si trabajamos con ejercicios cerrados
    F = UltimaFechaHcoCabapu
    If F >= FechaIncioEjercicio Then
        EjerciciosCerrados = True
    Else
        EjerciciosCerrados = False
    End If
    
    Screen.MousePointer = vbHourglass
    PulsadoCancelar = False
    Me.cmdCancelarAccion.visible = True
    If ObtenerDatosCCCtaExp Then
        'Las cadenas
        SQL = "Desde " & Text3(19).Text & "  hasta  " & Text3(20).Text
            
        RC = ""
        If txtCta(16).Text <> "" Then RC = "Desde cuenta: " & txtCta(16).Text
        If txtCta(17).Text <> "" Then
        If RC = "" Then
                RC = "H"
            Else
                RC = RC & "  h"
            End If
            RC = RC & "asta cuenta: " & txtCta(17).Text
        End If
        
        
        If txtCCost(6).Text <> "" Then
            If RC <> "" Then RC = RC & "     "
            RC = RC & "Desde Centro coste: " & txtCCost(6).Text
        End If
                
        
        If txtCCost(7).Text <> "" Then
            If RC <> "" Then RC = RC & "     "
            RC = RC & "Hasta Centro coste: " & txtCCost(7).Text
        End If
        
        SQL = """" & SQL & """"
        If RC <> "" Then
            RC = " """ & RC & """"
            SQL = SQL & RC
        End If
        RC = "Cuenta= " & SQL & "|"
        
        With frmImprimir
            .OtrosParametros = RC
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 41
            .Show vbModal
        End With
    End If
    Me.cmdCancelarAccion.visible = False
    Label2(27).visible = False
    pb7.visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDatosCarta_Click()
    'Para ver los datos de las carta
    Screen.MousePointer = vbHourglass
    frmMensajes.Opcion = 11
    frmMensajes.Show vbModal
End Sub

Private Sub cmdDiarioRes_Click()

    cmdCancelarAccion.visible = False
    

    Label2(25).Caption = ""
    If txtAno(11).Text = "" Or txtAno(12).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    If Me.cmbFecha(11).ListIndex < 0 Then
       MsgBox "Seleccione mes consulta desde", vbExclamation
       Exit Sub
    End If
    If Me.cmbFecha(12).ListIndex < 0 Then
       MsgBox "Seleccione mes consulta hasta", vbExclamation
       Exit Sub
    End If
    
    If Not ComparaFechasCombos(11, 12, 11, 12) Then Exit Sub
'    If txtAno(11).Text <> "" And txtAno(12).Text <> "" Then
'        If Val(txtAno(11).Text) > Val(txtAno(12).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        Else
'            If Val(txtAno(11).Text) = Val(txtAno(12).Text) Then
'                If Me.cmbFecha(11).ListIndex > Me.cmbFecha(12).ListIndex Then
'                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
    If Text3(18).Text <> "" Then
        If Not IsDate(Text3(18).Text) Then
            MsgBox "Fecha impresión incorrecta", vbExclamation
            Text3(18).SetFocus
        End If
    End If
    
    
    If Abs(Val(txtAno(11).Text) - Val(txtAno(12).Text)) > 2 Then
        MsgBox "Fechas pertenecen a ejercicios distintos.", vbExclamation
        Exit Sub
    End If


    'Fechas
    'Trabajaresmos contra ejercicios cerrados
    'Si el mes es mayor o igual k el de inicio, significa k la feha
    'de inicio de aquel ejercicio fue la misma k ahora pero de aquel año
    'si no significa k fue la misma de ahora pero del año anterior
    I = cmbFecha(11).ListIndex + 1
    If I >= Month(vParam.fechaini) Then
        Cont = Val(txtAno(11).Text)
    Else
        Cont = Val(txtAno(11).Text) - 1
    End If
    Cad = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Cont
    FechaIncioEjercicio = CDate(Cad)
    
    I = cmbFecha(12).ListIndex + 1
    If I <= Month(vParam.fechafin) Then
        Cont = Val(txtAno(12).Text)
    Else
        Cont = Val(txtAno(12).Text) + 1
    End If
    Cad = Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & Cont
    FechaFinEjercicio = CDate(Cad)

    
    
    
    
    
    'Veamos si pertenecen a un mismo año
    If Abs(DateDiff("d", FechaFinEjercicio, FechaIncioEjercicio)) > 365 Then
        MsgBox "Las fechas son incorrectas. Abarca mas de un ejercicio", vbExclamation
        Exit Sub
    End If


    'AHora, si ha puesto importes, entonces veremos
    'Si :  -importes correctos.
    '      -si exite importe, que no sea mes inicio ejerecicio
    txtNumRes(3).Text = Trim(txtNumRes(3).Text)
    txtNumRes(4).Text = Trim(txtNumRes(4).Text)
    If txtNumRes(3).Text <> "" Or txtNumRes(4).Text <> "" Then
       If cmbFecha(11).ListIndex + 1 = Month(FechaIncioEjercicio) Then
            MsgBox "No puede poner importes para el mes de inicio de ejerecicio", vbExclamation
            Exit Sub
        End If
    End If
    
    'Solo un nivel seleccionado
    Cont = 0
    For I = 1 To 9
        If ChkNivelRes(I).visible = True Then
            If ChkNivelRes(I).Value Then Cont = Cont + 1
        End If
    Next I
    If Cont <> 1 Then
        MsgBox "Seleccione uno, y solo uno, de los niveles para mostrar el informe", vbExclamation
        Exit Sub
    End If
    




    Screen.MousePointer = vbHourglass
    If GenerarLibroResumen Then
        If txtNumRes(1).Text <> "" Then
            Cont = Val(txtNumRes(1).Text) - 1
        Else
            Cont = 0
        End If
        Cad = "npag= " & Cont & "|"
        'Fecha impresuion
        If Text3(18).Text = "" Then
            RC = Format(Now, "dd/mm/yyyy")
        Else
            RC = Format(Text3(18).Text, "dd/mm/yyyy")
        End If
        RC = """" & RC & """"
        RC = "FechaIMP= " & RC & "|"
        SQL = Cad & RC
        If Opcion = 18 Then
            frmImprimir.Opcion = 40
            frmImprimir.NumeroParametros = 2
            frmImprimir.OtrosParametros = SQL
            frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            frmImprimir.SoloImprimir = False
            frmImprimir.Show vbModal
        Else
            'LEGALIZACION
            GeneraLegalizaPRF SQL, 2
            CadenaDesdeOtroForm = "OK"
        End If
    End If
    Screen.MousePointer = vbDefault
    Label2(25).Caption = ""
End Sub

Private Sub cmdEvolMensSald_Click()
    
    
    SQL = ""
    For I = 1 To 10
        If Me.ChkEvolSaldo(I).visible Then
            If Me.ChkEvolSaldo(I).Value = 1 Then SQL = SQL & "1"
        End If
    Next I
    
    If Len(SQL) <> 1 Then
        MsgBox "Eliga un nivel (y solo uno) para el listado de  evolución mesual de saldos", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    SQL = "DELETE FROM Usuarios.ztmpconextcab where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "DELETE FROM Usuarios.ztmpconext where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    DoEvents 'Para que no bloquee la pantalla
    Label2(29).Caption = "Leyendo datos BD"
    Label2(29).Refresh
    If ListadoEvolucionMensual Then
        
        
        'Lo imprimimos
        SQL = "Fechas= """"|"
    
        'Cuentas
        RC = cmbEjercicios(0).List(cmbEjercicios(0).ListIndex)
        RC = "Ejercicio " & Mid(RC, 1, 23) & "  "
        If txtCta(23).Text <> "" Then RC = RC & " desde " & txtCta(23).Text & " -" & DtxtCta(23).Text
        If txtCta(24).Text <> "" Then RC = RC & " hasta " & txtCta(24).Text & " -" & DtxtCta(24).Text
        If RC <> "" Then RC = "Cuentas: " & RC
        SQL = SQL & "Cuenta= """ & RC & """|"
        
        'Fecha impresion
        SQL = SQL & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
        
        
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .Opcion = 76
            .Show vbModal
        End With

        
    
    End If
    Screen.MousePointer = vbDefault
End Sub

'///////////////////////////////////////////////////////////////////////////
'
'     FACTURAS CLIENTES

Private Sub cmdFacProv_Click()
Dim B As Boolean
Dim Orde As String
Dim OtroCampo As Byte
Dim TipoDeIVA As Integer
    If Not ComprobarCuentas(20, 21) Then Exit Sub
    
    Tablas = "cabfactprov"
    
    SQL = ""
    Cad = ""
    txtCta(1).Tag = ""
    If txtCta(20).Text <> "" Then
        SQL = " " & Tablas & ".codmacta >= '" & txtCta(20).Text & "'"
        txtCta(1).Tag = "Desde " & txtCta(20).Text & " " & DtxtCta(20).Text
    End If
    
    If txtCta(21).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " " & Tablas & ".codmacta <= '" & txtCta(21).Text & "'"
        If txtCta(1).Tag <> "" Then
            txtCta(1).Tag = txtCta(1).Tag & "     h"
        Else
            txtCta(1).Tag = "H"
        End If
        txtCta(1).Tag = txtCta(1).Tag & "asta " & txtCta(21).Text & " " & DtxtCta(21).Text
    End If
    txtCta(0).Tag = SQL  'Para las cuentas
    Cad = "Cuenta= """ & txtCta(1).Tag & """|"

    'Las fechas
'    If Text3(29).Text <> "" And Text3(30).Text <> "" Then
'        If CDate(Text3(29).Text) > CDate(Text3(30).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(29, 30) Then Exit Sub
    
    
    'Como podemos ordenar por fecha registro o fecha liquidacion
    If optSelFech(1).Value Then
        Tablas = "fecliqpr"
    Else
        Tablas = "fecrecpr"
    End If
    
    SQL = ""
    Text3(1).Tag = ""
    If Text3(29).Text <> "" Then
        SQL = Tablas & " >= '" & Format(Text3(29).Text, FormatoFecha) & "'"
        Text3(1).Tag = "Desde " & Text3(29).Text
    End If
    
    If Text3(30).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Tablas & " <= '" & Format(Text3(30).Text, FormatoFecha) & "'"
        If Text3(1).Tag <> "" Then
            Text3(1).Tag = Text3(1).Tag & "     h"
        Else
            Text3(1).Tag = "H"
        End If
        Text3(1).Tag = Text3(1).Tag & "asta " & Text3(30).Text
    End If
    Text3(0).Tag = SQL  'Para las fechas
    
    
    
    Cad = Cad & "Fechas= """ & Text3(1).Tag


    If txtPag2(3).Text <> "" Then Cad = Trim(Cad & "   NIF: " & txtPag2(3).Text)
    
    TipoDeIVA = -1
    If Combo8.ListIndex > 0 Then
        TipoDeIVA = CInt(Combo8.ItemData(Combo8.ListIndex))
        Cad = Trim(Cad & "      IVA: " & Combo8.Text)
    End If
  
    txtSerie(1).Tag = ""



    'Numero factura
    Tablas = " numregis "
    SQL = ""
    txtNumFac(1).Tag = ""
    If txtNumFac(4).Text <> "" Then
        SQL = Tablas & " >=" & txtNumFac(4).Text
        txtNumFac(1).Tag = "Desde nº registro " & txtNumFac(4).Text
    End If
    If txtNumFac(5).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Tablas & " <=" & txtNumFac(5).Text
        If txtNumFac(1).Tag = "" Then
            txtNumFac(1).Tag = "H"
        Else
            txtNumFac(1).Tag = txtNumFac(1).Tag & "    h"
        End If
        txtNumFac(1).Tag = txtNumFac(1).Tag & "asta nº registro  " & txtNumFac(5).Text
    End If
    txtNumFac(0).Tag = SQL
    If txtNumFac(1).Tag <> "" Then If Cad <> "" Then Cad = Cad & """ + chr(13) + """ & txtNumFac(1).Tag
    Cad = Cad & """|" 'Para los parematros en el formulario de imprimir
    
    SQL = ""
    SQL = txtNumFac(0).Tag
    If txtCta(0).Tag <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & txtCta(0).Tag
    End If
    
    If txtSerie(0).Tag <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & txtSerie(0).Tag
    End If
    
    If Text3(0).Tag <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Text3(0).Tag
    End If
    
    
    
    'Nuevo Abril 2006
    '--------------------
    If Text4(0).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " numfacpr >= '" & DevNombreSQL(Text4(0).Text) & "'"
    End If
    
    If Text4(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " numfacpr <= '" & DevNombreSQL(Text4(1).Text) & "'"
    End If
    
     'Si selecicona la opcion de ordenar por fecha de emision, pero no marca ver
    ' fecha emision, en lugar de numero factura, se lo avisio
    If optListFacP(1).Value Then
        If optMostrarFecha(0).Value Then
            RC = "Ha selecionado ordenar por fecha emision factura, pero no ha marcado ver la fecha emision."
            RC = RC & vbCrLf & Space(30) & "¿Continuar ?"
            If MsgBox(RC, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        End If
    End If
    
    
    
    RC = ""
    For I = 0 To optListFacP.Count - 1
        If optListFacP(I).Value Then RC = RC & I
    Next I

    If Len(RC) <> 1 Then
        MsgBox "Seleccione un tipo de ordenación para el listado", vbExclamation
        Exit Sub
    End If
    I = Val(RC)
    
    Select Case I
    Case 0
           RC = "numregis"
    Case 1
            RC = "fecfacpr,numregis"
    Case 2
           If optSelFech(1).Value Then
                RC = "fecliqpr,numregis"
            Else
                RC = "fecrecpr,numregis"
            End If
    End Select
    
    
    OtroCampo = 0
    For Cont = 0 To 2
        If optMostrarFecha(Cont).Value Then OtroCampo = Cont
    Next Cont
    
    'Si es consilidado
    If Opcion = 52 Then
        If List9.ListCount = 0 Then
            MsgBox "Seleccione alguna empresa.", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    'Renumerar-. Modificacion renumera todo el mundo
    Cont = -1
    'If vParam.Constructoras Then
        If ChkListFac(2).Value = 1 Then
            Orde = "registro"
            Cont = DevuelveRegistrosProveedores
            Orde = "Escriba el numero de " & Orde
            Orde = InputBox(Orde, "Renumerar", Cont)
            Orde = Trim(Orde)
            If Orde = "" Then Exit Sub
            If Val(Orde) = 0 Then
                MsgBox "Numero de renumeracion no válido: " & Orde, vbExclamation
                Exit Sub
            End If
            Cont = CLng(Orde)
        End If
    'End If
    
    
    Screen.MousePointer = vbHourglass

     If Opcion = 52 Then
        
        B = ListadoFacturasProveedoresConsolidado2(SQL, RC, Cont, (ChkListFac(3).Value = 1), OtroCampo, optSelFech(1).Value, True, Trim(txtPag2(3).Text))
        
     Else
        B = ListadoFacturasProveedores(SQL, RC, Cont, (ChkListFac(3).Value = 1), OtroCampo, optSelFech(1).Value, Me.chkMostrarRetencion(1).Value = 1, Trim(txtPag2(3).Text), TipoDeIVA)
     End If
     
    'Para los textos de f.recepcion, liquidacion o factura
    Select Case OtroCampo
    Case 1
        RC = "numfac= ""Fecha Fact.""|"
    Case 2
        If optSelFech(0).Value Then
            RC = "Liquida."
        Else
            RC = "Recepción"
        End If
        RC = "numfac= """ & RC & """|"
    Case Else
        RC = "numfac= ""Nº Fact.""|"
    End Select
    
    If optListFacP(1).Value Then
        RC = RC & "fecrec= ""Fec. Fact""|"
    Else
        If optSelFech(1).Value Then
            
            RC = RC & "fecrec= ""Liquidación""|"
        Else
            RC = RC & "fecrec= ""Recepción""|"
        End If
    End If
    Cad = Cad & RC
    
    
    'Para el informe
    If Opcion = 52 Then
        'CONSOLIDADAS
        RC = ""
        For I = 1 To List9.ListCount
            If RC <> "" Then RC = RC & "' + Chr(13) + '"
            RC = RC & List9.List(I - 1)
        Next I
        RC = "Empresas= " & "'" & RC & "'|"
        Cad = Cad & RC

        If optListFacP(0).Value Then
            I = 69
        Else
            I = 70
        End If
        
    Else
        
         If ChkListFac(3).Value = 1 Then
             I = 58
         Else
             I = 31
         End If

    End If


    'El numero de pagina
    If txtNpag2(0).Text <> "" Then
        RC = Val(txtNpag2(0).Text)
    Else
        RC = "-1"   'Ponemos un menos 1 para k asi aprezca elcontador normal (ejemplo):  2 de 35
    End If
    Cad = Cad & "Numpag= " & RC & "|"


        
    'Fecha inofrme
    If Text3(32).Text = "" Then
        RC = Format(Now, "dd/mm/yyyy")
    Else
        RC = Text3(32).Text
    End If
    Cad = Cad & "FechaImp= """ & RC & """|"
    
    If Opcion = 13 Then
        If TipoDeIVA >= 0 Then Cad = Cad & "MostrarNIF= 1|"
    End If
    
    'Imprimimos
    If B Then
        If Opcion = 13 Or Opcion = 52 Then
            
            With frmImprimir
            
                Cad = Cad & "MostrarTipoIVA= " & Me.ChkListFac(5).Value & "|"
                If Opcion = 13 Then
                    .OtrosParametros = Cad & "MostrarRetencion= " & Me.chkMostrarRetencion(1).Value & "|"
                    .NumeroParametros = 9
                Else
                    .OtrosParametros = Cad
                    .NumeroParametros = 8 'De fecfac,numfac  y el uno de numpag y fechainofrme
                End If
                
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = I
                .Show vbModal
            End With
        Else
            GeneraLegalizaPRF Cad, 6
            CadenaDesdeOtroForm = "OK"
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

'///////////////////////////////////////////////////////////////////////////
'
'     FACTURAS CLIENTES
Private Sub cmdFactCli_Click()
Dim B As Boolean
Dim Orde As String


    If Not ComprobarCuentas(8, 9) Then Exit Sub
    
    
    If Opcion = 53 Then
        If List10.ListCount = 0 Then
            MsgBox "Seleccione alguna empresa", vbExclamation
            Exit Sub
        End If
    End If
    
    
    Tablas = "cabfact"
    SQL = ""
    Cad = ""
    txtCta(1).Tag = ""
    If txtCta(8).Text <> "" Then
        SQL = " " & Tablas & ".codmacta >= '" & txtCta(8).Text & "'"
        txtCta(1).Tag = "Desde " & txtCta(8).Text & " " & DtxtCta(8).Text
    End If
    
    If txtCta(9).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " " & Tablas & ".codmacta <= '" & txtCta(9).Text & "'"
        If txtCta(1).Tag <> "" Then
            txtCta(1).Tag = txtCta(1).Tag & "     h"
        Else
            txtCta(1).Tag = "H"
        End If
        txtCta(1).Tag = txtCta(1).Tag & "asta " & txtCta(9).Text & " " & DtxtCta(9).Text
    End If
    txtCta(0).Tag = SQL  'Para las cuentas
    Cad = "Cuenta= """ & txtCta(1).Tag & """|"

    'Las fechas
'    If Text3(10).Text <> "" And Text3(11).Text <> "" Then
'        If CDate(Text3(10).Text) > CDate(Text3(11).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(10, 11) Then Exit Sub
    
    
    If optListFac(3).Value Then
        Tablas = "fecliqcl"
    Else
        Tablas = "fecfaccl"
    End If
    
    SQL = ""
    Text3(1).Tag = ""
    If Text3(10).Text <> "" Then
        SQL = Tablas & " >= '" & Format(Text3(10).Text, FormatoFecha) & "'"
        Text3(1).Tag = "Desde " & Text3(10).Text
    End If
    
    If Text3(11).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Tablas & " <= '" & Format(Text3(11).Text, FormatoFecha) & "'"
        If Text3(1).Tag <> "" Then
            Text3(1).Tag = Text3(1).Tag & "     h"
        Else
            Text3(1).Tag = "H"
        End If
        Text3(1).Tag = Text3(1).Tag & "asta " & Text3(11).Text
    End If
    Text3(0).Tag = SQL  'Para las fechas
    
    Cad = Cad & "Fechas= """ & Text3(1).Tag

    SQL = ""
    txtSerie(1).Tag = ""
    If txtSerie(0).Text <> "" Then
        SQL = " numserie >= '" & txtSerie(0).Text & "'"
        txtSerie(1).Tag = "Serie desde " & txtSerie(0).Text
    End If
    
    If txtSerie(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " numserie <= '" & txtSerie(1).Text & "'"
        If txtSerie(1).Tag <> "" Then
            txtSerie(1).Tag = txtSerie(1).Tag & "     serie "
        Else
            txtSerie(1).Tag = "Serie "
        End If
        txtSerie(1).Tag = txtSerie(1).Tag & "hasta " & txtSerie(1).Text
    End If
    txtSerie(0).Tag = SQL  'Para las cuentas
    If txtSerie(1).Tag <> "" Then If Cad <> "" Then Cad = Cad & "     " & txtSerie(1).Tag
    'Cad = Cad & """|"


    If txtPag2(2).Text <> "" Then Cad = Trim(Cad & "   NIF: " & txtPag2(2).Text)
    
        
    Tablas = " codfaccl "
    SQL = ""
    txtNumFac(1).Tag = ""
    If txtNumFac(0).Text <> "" Then
        SQL = Tablas & " >=" & txtNumFac(0).Text
        txtNumFac(1).Tag = "Desde factura " & txtNumFac(0).Text
    End If
    If txtNumFac(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Tablas & " <=" & txtNumFac(1).Text
        If txtNumFac(1).Tag = "" Then
            txtNumFac(1).Tag = "H"
        Else
            txtNumFac(1).Tag = txtNumFac(1).Tag & "    h"
        End If
        txtNumFac(1).Tag = txtNumFac(1).Tag & "asta factura " & txtNumFac(1).Text
    End If
    txtNumFac(0).Tag = SQL
    If txtNumFac(1).Tag <> "" Then If Cad <> "" Then Cad = Cad & """+ chr(13) +""" & txtNumFac(1).Tag
    Cad = Cad & """|" 'Para los parematros en el formulario de imprimir
    
    SQL = ""
    SQL = txtNumFac(0).Tag
    If txtCta(0).Tag <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & txtCta(0).Tag
    End If
    
    If txtSerie(0).Tag <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & txtSerie(0).Tag
    End If
    
    If Text3(0).Tag <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Text3(0).Tag
    End If
    
    RC = ""
    For I = 0 To optListFac.Count - 1
        If optListFac(I).Value Then RC = RC & I
    Next I
    If Len(RC) <> 1 Then
        MsgBox "Seleccione un tipo de ordenación para el listado", vbExclamation
        Exit Sub
    End If
    I = Val(RC)
    
    Select Case I
    Case 3
    
            RC = "fecliqcl"
    
    Case 0
        'Numero
    
        RC = "numserie,codfaccl"
    
    Case 1
        
    
        RC = "fecfaccl,codfaccl"
        

    Case 2
        RC = "fecfacpr"
    End Select
    
    
    'Renumerar. Modificacion renumera todo el mundo
    Cont = -1
    'If vParam.Constructoras Then
        If ChkListFac(1).Value = 1 Then
            If Opcion = 8 Then
                Orde = "factura cliente"
                'CLIENTES. Solo podemos renumerar serie a serie
                If txtSerie(1).Text = "" Then
                    MsgBox "Se renumera serie a serie", vbExclamation
                    Exit Sub
                End If
                If txtSerie(1).Text <> txtSerie(0).Text Then
                    MsgBox "Series distintas. Se renumera serie a serie", vbExclamation
                    Exit Sub
                 End If
            Else
                Orde = "registro"
            End If
            Cont = DevuelveRegistrosProveedores
            Orde = "Escriba el numero de " & Orde
            Orde = InputBox(Orde, "Renumerar", Cont)
            Orde = Trim(Orde)
            If Orde = "" Then Exit Sub
            If Val(Orde) = 0 Then
                MsgBox "Numero de renumeracion no válido: " & Orde, vbExclamation
                Exit Sub
            End If
            Cont = CLng(Orde)
        End If
    'End If
    Screen.MousePointer = vbHourglass
    If Opcion = 13 Then
'            If ChkListFac(0).Value = 1 Then
'                i = 58
'            Else
'                i = 31
'            End If
'
            MsgBox "Opcion incorrecta. Consulte soporte técnico", vbExclamation
        Else
        
            If Opcion = 8 Then
                B = ListadoFacturasClientes(SQL, RC, Cont, (ChkListFac(0).Value = 1), optListFac(3).Value, Me.chkMostrarRetencion(0).Value = 1, Trim(txtPag2(2).Text))
                
                If ChkListFac(0).Value = 1 Then
                    I = 57
                Else
                    I = 21
                End If
                
            Else
            
                B = ListadoFacturasProveedoresConsolidado2(SQL, RC, 0, False, 0, False, False, Trim(txtPag2(2).Text))
                
            End If
            
    End If
        'Imprimimos
    If B Then

        If optListFac(3).Value Then
            RC = "fecfac= ""Liquidación""|"
        Else
            RC = "fecfac= ""Fecha""|"
        End If
        Cad = Cad & RC
        
        
        'El numero de pagina
        If txtNpag2(1).Text <> "" Then
            RC = Val(txtNpag2(1).Text)
        Else
            RC = "1"   'Ponemos un menos 1 para k asi aprezca elcontador normal (ejemplo):  2 de 35
        End If
        Cad = Cad & "Numpag= " & RC & "|"
    
    
    
        'Fecha inofrme
        If Text3(31).Text = "" Then
            RC = Format(Now, "dd/mm/yyyy")
        Else
            RC = Text3(31).Text
        End If
        Cad = Cad & "FechaImp= """ & RC & """|"
    
    
    
        If Opcion = 53 Then
            RC = ""
            For I = 1 To List10.ListCount
                If RC <> "" Then RC = RC & "' + Chr(13) + '"
                RC = RC & List10.List(I - 1)
            Next I
            RC = "Empresas= " & "'" & RC & "'|"
            Cad = Cad & RC
    
            If Not optListFac(0).Value Then
                I = 72
            Else
                I = 71
            End If
        End If
    
    
        If Opcion = 8 Or Opcion = 53 Then
            
            Cad = Cad & "MostrarTipoIVA= " & Me.ChkListFac(4).Value & "|"
            RC = 7
            If Opcion = 8 Then
                RC = RC = RC + 1
                Cad = Cad & "MostrarRetencion= " & Me.chkMostrarRetencion(0).Value & "|"
            End If
            
                
        
            With frmImprimir
                .OtrosParametros = Cad
                .NumeroParametros = Val(RC)
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = I
                .Show vbModal
            End With
        Else
            GeneraLegalizaPRF Cad, 6
            CadenaDesdeOtroForm = "OK"
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdLibroDiario_Click()
'    If Text3(15).Text <> "" And Text3(16).Text <> "" Then
'        If CDate(Text3(15).Text) > CDate(Text3(16).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(15, 16) Then Exit Sub
    
    If chkRenumerar.Value = 1 Then
        'Si ha pedido renumerar, es obligado el numero de asiento
        If txtLibroOf(0).Text = "" Then
            MsgBox "Tiene que poner el numero del primer asiento", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtExplo(4).Text <> "" Then
        If Val(txtExplo(4).Text) = 0 Then txtExplo(4).Text = ""
    End If
    
    If txtExplo(5).Text <> "" Then
        If Val(txtExplo(5).Text) = 0 Then txtExplo(5).Text = ""
    End If
    

'        With frmImprimir
'            .OtrosParametros = SQL
'            .NumeroParametros = 4
'            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
'            .SoloImprimir = False
'            'Opcion dependera del combo
'            .Opcion = 14 + Opcion
'            .Show vbModal

    Screen.MousePointer = vbHourglass
    Set RS = New ADODB.Recordset
    
    If GeneraDiarioOficial Then
    
        'Los campos del informe
        '------------------------------------------------------------------
        
        'Fechas
        SQL = ""
        If Text3(15).Text <> "" Then SQL = "Del " & Text3(15).Text
        
        
        If Text3(16).Text <> "" Then
            If SQL <> "" Then
                SQL = SQL & "   al "
            Else
                SQL = "Hasta "
            End If
            SQL = SQL & Text3(16).Text
        End If
        SQL = "Fechas= """ & SQL & """|"
        
        
        'Fecha de impresion
        SQL = SQL & "FechaImp= """ & Text3(17).Text & """|"
        'Numero de hoja
        If txtLibroOf(1).Text <> "" Then
            I = Val(txtLibroOf(1).Text)
        Else
            I = 0
        End If
        SQL = SQL & "Numhoja= " & I & "|"
    
    
        'Acumulados anteriores
        If txtExplo(5).Text <> "" Or txtExplo(4).Text <> "" Then
            I = 0  'En el informe diremos k si se muestra
        Else
            I = 1
        End If
        SQL = SQL & "TieneAcumulados= " & I & "|"
    
        If I = 1 Then
            SQL = SQL & "AntD= 0|"
            SQL = SQL & "AntH= 0|"
        Else
            SQL = SQL & "AntD= " & TransformaComasPuntos(txtExplo(4).Text) & "|"
            SQL = SQL & "AntH= " & TransformaComasPuntos(txtExplo(5).Text) & "|"
        End If
        
        If Opcion = 14 Then
            With frmImprimir
                 .OtrosParametros = SQL
                 .NumeroParametros = 6
                 .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                 .SoloImprimir = False
                 'Opcion dependera del combo
                 .Opcion = 32
                 .Show vbModal
            End With
        Else
            'Opcion 32. Lgaliza libros
            'Generaremos el pdf
            GeneraLegalizaPRF SQL, 6
            CadenaDesdeOtroForm = "OK"
        End If
    End If
    Me.cmdCancelarAccion.visible = False
    pb6.visible = False
    Set RS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdListExtCta_Click()
Dim AnyoInicioEjercicio As String
Dim CuentasNIF As String

    'Ciertas comprobaciones
'    If txtCta(0).Text <> "" And txtCta(1).Text <> "" Then
'        If Val(txtCta(0).Text) > Val(txtCta(1).Text) Then
'            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarCuentas(0, 1) Then Exit Sub

    If Text3(0).Text = "" Or Text3(1).Text = "" Then
        MsgBox "Introduce las fechas de consulta de extractos", vbExclamation
        Exit Sub
    End If
'    If Text3(0).Text <> "" And Text3(1).Text <> "" Then
'        If CDate(Text3(0).Text) > CDate(Text3(1).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarFechas(0, 1) Then Exit Sub
    SQL = ""
    'Llegados aqui. Vemos la fecha y demas
    If Text3(0).Text <> "" Then
        SQL = " fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
        'lblFecha.Caption = "Desde " & Text3(0).Text
    End If
    
    If Text3(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
        'If lblFecha.Caption <> "" Then lblFecha.Caption = lblFecha.Caption & "     "
        'lblFecha.Caption = lblFecha.Caption & " hasta " & Text3(1).Text
    End If
    
    'Si solo quiere punteada o pendiente
    If Combo1.ListIndex > 1 Then
        If Combo1.ListIndex = 2 Then
            'Solo kiere punteadas
            SQL = SQL & " AND punteada = 1"
        Else
            'Solo kiere PENDIENTES DE PUNTEAR
            SQL = SQL & " AND punteada = 0"
        End If
    End If
    
    Text3(0).Tag = SQL  'Para las fechas
    
    If Text3(2).Text = "" Then
        MsgBox "Seleccione una fecha de impresón del informe", vbExclamation
        Exit Sub
    End If

    
    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    
    
    
    'Si es ejercicioscerrados
    AnyoInicioEjercicio = ""
    If EjerciciosCerrados Then
        FechaIncioEjercicio = CDate(Text3(0).Text)
        I = Month(FechaIncioEjercicio)
        If I >= Month(vParam.fechaini) Then
            Cad = Year(FechaIncioEjercicio)
        Else
            Cad = Year(FechaIncioEjercicio) - 1
        End If
        AnyoInicioEjercicio = Cad
    End If
        
     Set RS = New ADODB.Recordset
    
    
    
    
        
    'Febrero 2014
    'Solo si tiene ese NIF
    CuentasNIF = ""
    If Me.txtPag2(1).Text <> "" Then
        'Significa que ha puesto un NIF
        Cad = "select codmacta from cuentas WHERE APUDIREC='S' AND nifdatos='" & DevNombreSQL(txtPag2(1).Text) & "'"
        RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        Cad = ""
        While Not RS.EOF
            Cad = Cad & ", '" & DevNombreSQL(CStr(RS!codmacta)) & "'"
            RS.MoveNext
        Wend
        RS.Close
        
        If Cad = "" Then
            MsgBox "Ninguna cuenta con ese NIF", vbExclamation
            Exit Sub
        End If
        CuentasNIF = Mid(Cad, 2)
    End If
        
    Screen.MousePointer = vbHourglass
    
    'Hacemos el select y si tiene resultados mostramos los valores
    Cad = " SELECT hlinapu" & Tablas & ".codmacta From hlinapu" & Tablas
    Cad = Cad & " WHERE (((fechaent)>='" & Format(Text3(0).Text, FormatoFecha)
    Cad = Cad & "') AND ((fechaent)<='" & Format(Text3(1).Text, FormatoFecha) & "')"
    If txtCta(0).Text <> "" Then Cad = Cad & " AND ((codmacta)>='" & txtCta(0).Text & "')"
    If txtCta(1).Text <> "" Then Cad = Cad & " AND ((codmacta)<='" & txtCta(1).Text & "')"
    Cad = Cad & ")"
    'Si kiere solo punteadas o no
    If Combo1.ListIndex > 1 Then
        If Combo1.ListIndex = 2 Then
            'Solo kiere punteadas
           Cad = Cad & " AND punteada =1 "
        Else
            'Solo kiere PENDIENTES DE PUNTEAR
           Cad = Cad & " AND punteada =0 "
        End If
    End If
    
    
    'FEBRERO 2013
    'Ha pedido NIF
    If CuentasNIF <> "" Then Cad = Cad & " AND codmacta IN (" & CuentasNIF & ")"
    
    Cad = Cad & " GROUP BY hlinapu" & Tablas & ".codmacta "
    
    RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If RS.EOF Then
        'NO hay registros a mostrar
        MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
        Cont = -1
    Else
        'Mostramos el frame de resultados
        Cont = 0
        While Not RS.EOF
            Cont = Cont + 1
            RS.MoveNext
        Wend
        If Cont > 32000 Then Cont = 32000
        pb1.Max = Cont + 1
        pb1.visible = True
        Label12.visible = True
        pb1.Value = 0
        Me.Refresh
        
        'Borramos los temporales
        Label12.Caption = "Eliminando datos temporal 1"
        Label12.Refresh
        SQL = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        'La de informe
        SQL = "DELETE from Usuarios.ztmpconextcab where codusu= " & vUsu.Codigo
        Conn.Execute SQL
    
        Label12.Caption = "Eliminando datos temporal 2"
        Label12.Refresh
        SQL = "DELETE from tmpconext where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        SQL = "DELETE from usuarios.ztmpconext where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        
        'Lo mismo k el balance, como puede tardar mucho
        frameListadoCuentas.Enabled = False
        
        'Me.cmdCancelarAccion.Visible = True
        Me.cmdCancelarAccion.visible = Legalizacion = ""
        HanPulsadoSalir = False
        PulsadoCancelar = False
        
        
        pb1.Value = pb1.Value + 1
        Me.Refresh
        RS.MoveFirst
        While Not RS.EOF
            Label12.Caption = RS!codmacta
            Label12.Refresh
            SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RS!codmacta, "T")
            If EjerciciosCerrados Then
                CargaDatosConExtCerrados RS!codmacta, Text3(0).Text, Text3(1).Text, Text3(0).Tag, SQL, AnyoInicioEjercicio
            Else
                CargaDatosConExt RS!codmacta, Text3(0).Text, Text3(1).Text, Text3(0).Tag, SQL
            End If
            GeneraraExtractosListado RS!codmacta
            'Progress
            pb1.Value = pb1.Value + 1
            pb1.Refresh
            
            If (pb1.Value Mod 25) = 0 Then
                Me.Refresh
                
            End If
            DoEvents
            If PulsadoCancelar Then RS.MoveLast
                
            
            'Siguiente cta
            RS.MoveNext
        Wend
    End If
    RS.Close
    
    Label12.Caption = ""
    pb1.Value = 0
    
    Me.frameListadoCuentas.Enabled = True
    Me.cmdCancelarAccion.visible = False
    
    HanPulsadoSalir = True
    Me.Refresh
    If PulsadoCancelar Then
        Label12.Caption = ""
        pb1.visible = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    'Eliminamos datos si procede, segun la indiacion del combo7
    'Que dice listar todo, solo con saldo o cuentas saldadas
    If Cont > 0 Then
        If Combo7.ListIndex > 0 Then
            SQL = "from Usuarios.ztmpconextcab where codusu= " & vUsu.Codigo
            'SQL = SQL & " AND acumtotT"
            SQL = SQL & " AND acumperT"
            If Combo7.ListIndex = 1 Then
                SQL = SQL & "=0"
            Else
                SQL = SQL & "<>0"
            End If
            RS.Open "Select count(cta) " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 0
            If Not RS.EOF Then
                I = DBLet(RS.Fields(0), "N")
            End If
            RS.Close
            pb1.Value = 0
            If I > 0 Then
                pb1.Max = I
                RS.Open "Select cta " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    I = 0
                    Label12.Caption = "Cálculo por saldo"
                    Label12.Refresh
                    SQL = "WHERE codusu =" & vUsu.Codigo & " AND cta ='"
                    While Not RS.EOF
                        I = I + 1
                        pb1.Value = I
                        'Borramos las lineas
                        RC = "DELETE FROM Usuarios.ztmpconext " & SQL & RS.Fields(0) & "'"
                        Conn.Execute RC
                        
                        'Las cabeceras
                        RC = "DELETE FROM Usuarios.ztmpconextcab " & SQL & RS.Fields(0) & "'"
                        Conn.Execute RC
                        
                        Cont = Cont - 1
                        'Siguiente
                        RS.MoveNext
                    Wend
                End If
                RS.Close
            End If
            
            'Si conta vale 0.. no queda para imprimir
            If Cont < 1 Then MsgBox "No hay datos con esos parametros", vbExclamation
         
        End If
    End If
    Set RS = Nothing
    
    'Quitamos progress
    pb1.Value = 0
    pb1.visible = False
    Label12.visible = False
    Me.Refresh
    
    If Cont > 0 Then
            'Si la opcion del listado es EXTENDIDO tengo k grabar un par de tablas mas
        
        If Combo1.ListIndex = 4 Then
             Label12.visible = True
              
            'EXTENDIDO. Guardare en una tabla los datos de la contrapartida
            'para que se vea en el listado extendido
            DatosConsultaExtractoExtendida
            
            Label12.visible = False
        End If
    
    End If
    DoEvents
    'Si hay datos los mostramos
    If Cont > 0 Then ImprimirListadoCuentas

    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPersa_Click()
    Screen.MousePointer = vbHourglass
    If TraspasoPERSA(optActual(0).Value) Then CopiarPersa
    Me.lblPersa.Caption = ""
    Me.lblPersa2.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPresupuestos_Click()
    RC = ""
    Cad = ""
    'Ciertas comprobaciones
'    If txtCta(10).Text <> "" And txtCta(11).Text <> "" Then
'        If Val(txtCta(10).Text) > Val(txtCta(11).Text) Then
'            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
'            Exit Sub
'        End If
'    End If
    If Not ComprobarCuentas(10, 11) Then Exit Sub
    
    'COMPROBAR MES Y AÑO
    For I = 0 To 1
        If Not ComprobarObjeto(txtAno(I + 2)) Then Exit Sub
        If Not ComprobarObjeto(txtMes(I)) Then Exit Sub
        If txtAno(I + 2).Text = "" Xor txtMes(I).Text = "" Then
            If I = 0 Then
                Cad = "'desde'"
            Else
                Cad = "'hasta'"
            End If
            MsgBox "Tiene que poner valor MES/AÑO para " & Cad, vbExclamation
            Exit Sub
        End If
    Next I
    
    SQL = ""
    RC = ""
    If txtCta(10).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        RC = "Desde " & txtCta(10).Text & " - " & DtxtCta(10).Text
        SQL = SQL & "presupuestos.codmacta >= '" & txtCta(10).Text & "'"
    End If
    
    
    If txtCta(11).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        If RC <> "" Then
            RC = RC & "       h"
        Else
            RC = "H"
        End If
        RC = RC & "asta " & txtCta(11).Text & " - " & DtxtCta(11).Text
        SQL = SQL & "presupuestos.codmacta <= '" & txtCta(11).Text & "'"
    End If
    Cad = RC
    RC = ""
    '--------------------------------- MES /AÑO ----------------
    If txtMes(0).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        RC = "Desde mes/año: " & txtMes(0).Text & "/" & txtAno(2).Text
        SQL = SQL & "mespresu >= " & txtMes(0).Text & " AND anopresu >=" & txtAno(2).Text
    End If
        
    If txtMes(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        If RC <> "" Then
            RC = RC & "   h"
        Else
            RC = "H"
        End If
        RC = RC & "asta mes/año: " & txtMes(1).Text & "/" & txtAno(3).Text
        SQL = SQL & "mespresu <= " & txtMes(1).Text & " AND anopresu <=" & txtAno(3).Text
    End If
    If RC <> "" Then
        If Cad <> "" Then Cad = Cad & "' + Chr(13) + '"
        Cad = Cad & RC
    End If
    RC = Cad
    
    'Llegados aqui generamos los registros
    If GneraListadoPresupuesto() Then
    
        'Param
        If RC <> "" Then
            RC = "CampoSeleccion= '" & RC & "'|"
            I = 1
        End If
        'Mandamos a imprimir
        With frmImprimir
            .OtrosParametros = RC
            .NumeroParametros = I
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            
            .Opcion = 22
            .Show vbModal
        End With
    End If
End Sub

Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function


Private Sub cmdRelacion_x_Ctabases_Click()
Dim B As Boolean

    If Me.chkCliproxCtalineas.Value = 1 And Me.chkDesgloseBasexCta.Value = 1 Then
        MsgBox "No puede seleccionar las dos opciones a la vez.", vbExclamation
        Exit Sub
    End If
        
    
    If Not ComprobarCuentas(25, 26) Then Exit Sub
    
    If Not ComprobarCuentas(27, 28) Then Exit Sub
        
    If Not ComprobarFechas(36, 37) Then Exit Sub
    
    
    If Me.chkCliproxCtalineas.Value = 1 Then
        If Text3(36).Text = "" Or Text3(37).Text = "" Then
            MsgBox "Si marca la opcion de comparativo  debe indicar las fechas", vbExclamation
            Exit Sub
        End If
        
        If DateDiff("yyyy", CDate(Text3(36).Text), CDate(Text3(37).Text)) > 0 Then
            MsgBox "Si marca la opcion de comparativo , el periodo no puede ser superior a un año", vbExclamation
            Exit Sub
        End If
        
    End If
    
    
    '--------------------------------------------------
    Screen.MousePointer = vbHourglass
    Me.cmdCancelarAccion.visible = True
    PulsadoCancelar = False
    cmdCanListExtr(55).visible = False
    Me.Refresh
    DoEvents
    'Obtener datos
    B = GeneraRelacionCta_x_Bases
    
    Label2(24).Caption = ""
    
    Me.cmdCancelarAccion.visible = False
    cmdCanListExtr(55).visible = True
    Screen.MousePointer = vbDefault
    If B Then
    
        'OK. TIene datos
        RC = ""
        
        
        If Text3(36).Text <> "" Then RC = RC & " desde " & Text3(36).Text
        If Text3(37).Text <> "" Then RC = RC & " hasta " & Text3(37).Text
        If Me.chkCliproxCtalineas.Value = 1 Then RC = RC & "   COMPARATIVO"
        If RC <> "" Then RC = "Fecha: " & RC
        SQL = RC
        RC = ""
        If Opcion = 55 Then
            Cad = "CLI."
        Else
            Cad = "PRO."
        End If
        
        If txtCta(27).Text <> "" Then RC = RC & " desde " & txtCta(27).Text & " -" & DtxtCta(27).Text
        If txtCta(28).Text <> "" Then RC = RC & " hasta " & txtCta(28).Text & " -" & DtxtCta(28).Text
        If RC <> "" Then RC = " Cta " & Cad & RC
        RC = SQL & RC
        CadenaDesdeOtroForm = "Fechas= """ & RC & """|"
    
    
        'Titulo
        If Opcion = 55 Then
            SQL = "Clientes"
            Cad = "ventas"
        Else
            
            SQL = "Prov."
            Cad = "compras"
        End If
        Tablas = "Relación de " & SQL & " por cta. " & Cad
        Tablas = "Titulo= """ & Tablas & """|"
        
    
    
        'Cuentas
        RC = ""
        If txtCta(25).Text <> "" Then RC = RC & " desde " & txtCta(25).Text & " -" & DtxtCta(25).Text
        If txtCta(26).Text <> "" Then RC = RC & " hasta " & txtCta(26).Text & " -" & DtxtCta(26).Text
        If RC <> "" Then RC = "Cuentas bases: " & RC
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Cuenta= """ & RC & """|"
        SQL = CadenaDesdeOtroForm
        'Fecha impresion
        SQL = SQL & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
        SQL = SQL & Tablas
        
        I = 77
        If Me.chkDesgloseBasexCta.Value = 1 Then I = 78
        If Me.chkCliproxCtalineas = 1 Then I = 82
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 4
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With

     
    
    
    
    
    
    
    
    
    
    End If
    CadenaDesdeOtroForm = ""
End Sub

Private Sub cmdSaldosCC_Click()
Dim MesFin1 As Integer
Dim AnoFin1 As Integer

    On Error GoTo ESaldosCC
    If txtCCost(0).Text <> "" And txtCCost(1).Text <> "" Then
        If txtCCost(0).Text > txtCCost(1).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtAno(5).Text = "" Or txtAno(6).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Sub
    End If
    If Not ComparaFechasCombos(5, 6, 3, 4) Then Exit Sub
'    If txtAno(5).Text <> "" And txtAno(6).Text <> "" Then
'        If Val(txtAno(5).Text) > Val(txtAno(6).Text) Then
'            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'            Exit Sub
'        Else
'            If Val(txtAno(5).Text) = Val(txtAno(6).Text) Then
'                If Me.cmbFecha(4).ListIndex > Me.cmbFecha(4).ListIndex Then
'                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
    
    
    'Llegamos aqui y hacemos el sql, Para ello, y por si acaso piden de cerrados
    'tenemos k comprobar cual es el ultimo mes en saldosanal1
    UltimoMesAnyoAnal1 MesFin1, AnoFin1
    
    
    'Si años consulta iguales
    If txtAno(5).Text = txtAno(6).Text Then
         Cad = " anoccost=" & txtAno(5).Text & " AND mesccost>=" & Me.cmbFecha(3).ListIndex + 1
         Cad = Cad & " AND mesccost<=" & Me.cmbFecha(4).ListIndex + 1
         
    Else
        'Años disitintos
        'Inicio
        Cad = "( anoccost=" & txtAno(5).Text & " AND mesccost>=" & Me.cmbFecha(3).ListIndex + 1 & ")"
        Cad = Cad & " OR ( anoccost=" & txtAno(6).Text & " AND mesccost<=" & Me.cmbFecha(4).ListIndex + 1 & ")"
        'Por si la diferencia es mas de un año
        If Val(txtAno(6).Text) - Val(txtAno(5).Text) > 1 Then
            Cad = Cad & " OR (anoccost >" & txtAno(5).Text & " AND anoccost < " & txtAno(6).Text & ")"
        End If
    End If
    Cad = " AND (" & Cad & ")"
    
    RC = ""
    If txtCCost(0).Text <> "" Then RC = " cabccost.codccost >='" & txtCCost(0).Text & "'"
    If txtCCost(1).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " cabccost.codccost <='" & txtCCost(1).Text & "'"
    End If
    
    
    'Borramos temporal
    Screen.MousePointer = vbHourglass
    Conn.Execute "Delete from Usuarios.zsaldoscc  where codusu = " & vUsu.Codigo
    
    
    'Haremos las inserciones
    SQL = "INSERT INTO Usuarios.zsaldoscc (codusu, codccost, nomccost, ano, mes, impmesde, impmesha) SELECT "
    SQL = SQL & vUsu.Codigo & ",cabccost.codccost,nomccost,anoccost,mesccost,sum(debccost),sum(habccost) from hsaldosanal,cabccost where"
    SQL = SQL & " cabccost.codccost =hsaldosanal.codccost "
    If RC <> "" Then SQL = SQL & " AND " & RC
    SQL = SQL & Cad
    SQL = SQL & " group by codccost,anoccost,mesccost"
    Conn.Execute SQL
    
    
    
    
    'Haremos las inserciones desde hsaldosanal 1, es decir, ejercicios traspasados
    'si la fecha de incio de los calculos es  menor k la ultima fecha k haya en hco 1
    ' EN i tneemos el año y en mesfin1 el ultimo mes grabado en saldosanal1
    Tablas = ""
    If Val(txtAno(5).Text) < AnoFin1 Then
        Tablas = "SI"
    Else
        If Val(txtAno(5).Text) = AnoFin1 Then
            'Dependera del mes
            If MesFin1 >= (Me.cmbFecha(4).ListIndex + 1) Then Tablas = "OK"
        End If
    End If
    
    If Tablas <> "" Then
        SQL = "INSERT INTO Usuarios.zsaldoscc (codusu, codccost, nomccost, ano, mes, impmesde, impmesha) SELECT "
        SQL = SQL & vUsu.Codigo & ",cabccost.codccost,nomccost,anoccost,mesccost,sum(debccost),sum(habccost) from hsaldosanal1,cabccost where"
        SQL = SQL & " cabccost.codccost =hsaldosanal1.codccost "
        If RC <> "" Then SQL = SQL & " AND " & RC
        SQL = SQL & Cad
        SQL = SQL & " group by codccost,anoccost,mesccost"
        Conn.Execute SQL
    End If
    
    Set miRsAux = New ADODB.Recordset
    SQL = "Select count(mes) from Usuarios.zsaldoscc  where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    If Not miRsAux.EOF Then
        I = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    
    If I = 0 Then
        MsgBox "Ningun dato con esos valores", vbExclamation
    Else
        'Titulitos
        RC = ""
        If txtCCost(0).Text <> "" Then RC = "Desde " & txtCCost(0).Text & " - " & txtDCost(0).Text
        If txtCCost(1).Text <> "" Then
            If RC = "" Then
                RC = "H"
            Else
                RC = RC & "     h"
            End If
            RC = RC & "asta " & txtCCost(1).Text & " - " & txtDCost(1).Text
        End If
        
        Cont = cmbFecha(3).ListIndex
        SQL = "Desde " & cmbFecha(3).List(Cont) & " - " & txtAno(5).Text
        Cont = cmbFecha(4).ListIndex
        SQL = SQL & "     hasta " & cmbFecha(4).List(Cont) & " - " & txtAno(6).Text
        
        If RC = "" Then
             RC = SQL
             SQL = ""
        End If
        Cad = "Cuenta= """ & RC & """|"
        Cad = Cad & "Fechas= """ & SQL & """|"
      
      
        
        
        With frmImprimir
            .OtrosParametros = Cad
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .Opcion = 33
            .Show vbModal
        End With

        
    End If
    
    
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ESaldosCC:
    MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then HacerF1
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
     ListadoKEYpress KeyAscii
End Sub


Private Sub Combo7_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 112 Then HacerF1
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub Command1_Click()

    'IVA IVA IVA IVA
    'MUEVO 1 AGOSTO 2005
    '------------------------------------------------
    ' No mostraba los recargos de equivalencia
    'Ahora, y debeido al orden que van a llevar en los listados
    ' el dato cliprov va a cambiar
    '         --------
    ' Antes era 0 cli 1 pro  2 NO decucibles 4 IVA NO dedu
    'ahora quedara
    '
    '
    '       cliprov     0- Facturas clientes
    '                   1- RECARGO EQUIVALENCIA !!nuevo
    '                   2- Facturas proveedores
    '                   3- libre                !!nuevo
    '                   4- IVAS no deducible
    '                   5- Facturas NO DEDUCIBLES
    '                   6- Facturas PRO con iva = bien de inversion
    For I = 0 To 2
        If Me.txtperiodo(I).Text = "" Then
            MsgBox "Campos periódo no pueden estar vacios", vbExclamation
            Exit Sub
        End If
    Next I
    
    If Val(txtperiodo(0).Text) > Val(txtperiodo(1).Text) Then
        MsgBox "Periódo desde mayor que periódo hasta.", vbExclamation
        Exit Sub
    End If
    
    
    If vParam.periodos = 1 Then
        If Val(txtperiodo(0).Text) > 12 Or Val(txtperiodo(1).Text) > 12 Then
            MsgBox "Periódo no puede ser superior a 12.", vbExclamation
            Exit Sub
        End If
    Else
        'TRIMESTRAL
        If Val(txtperiodo(0).Text) > 4 Or Val(txtperiodo(1).Text) > 4 Then
            MsgBox "Periódo no puede ser superior a 4.", vbExclamation
            Exit Sub
        End If
    End If
    
    
    If List2.ListCount = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Sub
    End If
    
    
    'Si esta marcado liquidacion definitiva comprobaremos k solo hay un periodo
    'y  k no es inferior a la k tenemos
    If chkLiqDefinitiva.Value = 1 Then
        If Val(txtperiodo(1).Text) - Val(txtperiodo(0).Text) > 0 Then
            MsgBox "Para las liquidaciones definitivas solo se abarca un periódo.", vbExclamation
            Exit Sub
        End If
        I = Val(txtperiodo(2).Text)
        Cont = 0
        If I < vParam.anofactu Then
            Cont = 1
        Else
            If I = vParam.anofactu Then
                If Val(txtperiodo(0).Text) <= vParam.perfactu Then Cont = 1
            End If
        End If
        If Cont = 1 Then
            MsgBox "El periodo es igual o inferior al último periódo contabilizado.", vbExclamation
            Exit Sub
        Else
            'Comproamos k esta dentro del año actual o siguiente. Como . Generando una fecha del periodo
            I = Val(txtperiodo(1).Text)
            If vParam.periodos = 0 Then I = I * 3  'Trimestral
            Cad = "15/" & I & "/" & txtperiodo(2).Text
            I = FechaCorrecta2(CDate(Cad))
            If I > 1 Then
                If I = 2 Then
                    Cad = varTxtFec   '¡fecha fuera ambito
                Else
                    Cad = "El periodo no pertenece al ejercicio actual o siguiente."
                End If
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
        End If
    End If

    'AHora generaremos la liquidacion para todos los periodos k abarque la seleecion
    Screen.MousePointer = vbHourglass
    'Guardamos el valor del chk del IVA
    ModeloIva False
    Label13.Caption = "Elimina datos anteriores"
    Label13.visible = True
    Label13.Refresh
    If GeneraLasLiquidaciones Then
        Label13.Caption = ""
        Label13.Refresh
        espera 0.5
        'Periodos
        SQL = ""
        For I = 0 To 2
            SQL = SQL & txtperiodo(I).Text & "|"
        Next I
        
        'Abril 2015
        'SIEMPRE SIEMPRE detallado
        'I = 0
        'If Me.chkIVAdetallado.Value Then
        '    If List2.ListCount = 1 Then I = 1
        'End If
        I = 1
        
        'Solo sera detallado cuando este marcado y ademas, solo sea una empresa
       ' If optModeloLiq(0).Value Then
      '      frmLiqIVA.Periodo = SQL & I & "|"
      '  Else
            frmLiqIVA2.Periodo = SQL & I & "|"
     '   End If
        
        
        'Modelo
        SQL = 0
        For I = 0 To 4
            If optModeloLiq(I).Value Then SQL = I
        Next I
        

        frmLiqIVA2.Modelo = CByte(SQL)
    
        'Empresas para consolidado
        SQL = ""
        If List2.ListCount = 1 Then
            If List2.List(0) <> vEmpresa.nomempre Then SQL = List2.List(0)
        Else
            'Mas de una empresa
            SQL = "'Empresas seleccionadas:' + Chr(13) "
            For I = 0 To List2.ListCount - 1
                SQL = SQL & " + '        " & List2.List(I) & "' + Chr(13)"
            Next I
        End If

        frmLiqIVA2.Consolidado = SQL

        
        

        frmLiqIVA2.FechaIMP = Text3(33).Text

        

        frmLiqIVA2.Show vbModal

        If Me.chkLiqDefinitiva.Value = 1 Then
            'Gurdamos los valores
            I = Val(txtperiodo(1).Text)
            Cont = Val(txtperiodo(2).Text) * 20 + I
            
            'Si son superiores a los anteriores
            NumRegElim = (Val(vParam.anofactu) * 20) + CInt(vParam.perfactu)
            If Cont > NumRegElim Then
                'El periodo solicitado es mayor k el k habia
                'Luego updateamos
                SQL = "UPDATE parametros SET anofactu=" & txtperiodo(2).Text & " , perfactu =" & txtperiodo(1).Text
                SQL = SQL & " WHERE fechaini='" & Format(vParam.fechaini, FormatoFecha) & "'"
                Conn.Execute SQL
                vParam.anofactu = Val(txtperiodo(2).Text)
                vParam.perfactu = Val(txtperiodo(1).Text)
            End If
            NumRegElim = 0
        End If
    End If
    Label13.visible = False
    Me.Refresh
    Screen.MousePointer = vbDefault
End Sub
    

Private Sub ModeloIva(Leer As Boolean)

On Error GoTo EModeloIva
    SQL = App.path & "\modiva.dat"
    If Leer Then
        If Dir(SQL) <> "" Then
            I = FreeFile
            Open SQL For Input As #I
            SQL = 0
            If Not EOF(I) Then Line Input #I, SQL
            Close (I)
            I = Val(SQL)
            Me.optModeloLiq(0).Tag = I
         Else
            I = 0
        End If
        If I < Me.optModeloLiq.Count Then Me.optModeloLiq(I).Value = True
        
    Else
        Cont = 0
        For I = 0 To Me.optModeloLiq.Count - 1
            If Me.optModeloLiq(I).Value Then
                Cont = I
                Exit For
            End If
        Next I
        If Val(Me.optModeloLiq(0).Tag) <> Cont Then
            I = FreeFile
            Open SQL For Output As #I
            Print #I, Cont
            Close (I)
        End If
    End If
    Exit Sub
EModeloIva:
    Err.Clear
End Sub





Private Sub Form_Activate()

    

    If PrimeraVez Then
        PrimeraVez = False
        CommitConexion
        'Ponemos el foco
        Select Case Opcion
        Case 1
            'Listado de EXTRACTOS DE cuentas
            '
        Case 2
            txtCta(3).SetFocus
        Case 3
    
        Case 4
            txtCta(5).SetFocus
        Case 5
            txtCta(6).SetFocus
        Case 6
            txtAsiento(0).SetFocus
        Case 7
            Text3(9).SetFocus
            
            
            
        Case 8
            txtSerie(0).SetFocus
        Case 9
            txtCta(10).SetFocus
        Case 10
            txtCta(12).SetFocus
        Case 11
            Text3(12).SetFocus
        Case 12
            txtperiodo(0).SetFocus
        Case 13
            'txtNumFac(0).SetFocus
        Case 14
            Text3(15).SetFocus
        Case 15
            txtCCost(0).SetFocus
        Case 16
            txtCCost(2).SetFocus
        Case 17
            txtCta(14).SetFocus
        Case 18
            cmbFecha(11).SetFocus
        Case 21
            cmbFecha(13).SetFocus
        Case 22
        
        Case 23
            txtNumFac(2).SetFocus
        Case 30
            Text1(1).SetFocus
            
        'Legalizacion de libros
        Case 32 To 41
            LegalizacionSub
                
        Case 52
            Text3(29).SetFocus
        Case 53
            Text3(10).SetFocus
            
        Case 54
            txtCta(23).SetFocus
        End Select
    End If
        Screen.MousePointer = vbDefault
End Sub

Private Sub LegalizacionSub()
            
            Screen.MousePointer = vbHourglass
            espera 0.1
            Me.Refresh
            Me.MousePointer = vbHourglass
            Select Case Opcion
            Case 32
                cmdLibroDiario_Click
            Case 33
                cmdDiarioRes_Click
            Case 34
                cmdListExtCta_Click
            Case 35
                cmdAceptarHco_Click
            Case 36, 41
                cmdBalance_Click
            Case 37
                cmdFactCli_Click
            Case 38
                cmdFacProv_Click
            Case 39, 40
                cmdBalances_Click
            End Select
            Me.Hide
            espera 0.1
            Me.MousePointer = vbHourglass
            Unload Me


End Sub



Private Sub Form_Load()
Dim H As Single
Dim W As Single
    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    Limpiar Me
    
    'He puesto FALSE a todos los frames en diseño
    
'    FrameCuentas.Visible = False
'    frameListadoCuentas.Visible = False
'    frameAsiento.Visible = False
'    frameCtaConcepto.Visible = False
'    frameDiarioHco.Visible = False
'    frameBalance.Visible = False
'    frameExplotacion.Visible = False
'    frameListFacCli.Visible = False
'    FrameListFactP.Visible = False
'    Me.FramePresu.Visible = False
'    FrameBalPresupues.Visible = False
'    frameIVA.Visible = False
'    Frame4.Visible = False
'    Me.FrameLiq.Visible = False
'    Me.frameLibroDiario.Visible = False
'    frameCCostSaldos.Visible = False
'    frameCtaExpCC.Visible = False
'    frameCCxCta.Visible = False
'    frameResumen.Visible = False
'    frameccporcta.Visible = False
'    Frame347.Visible = False
'    Frame349.Visible = False
'    frameComparativo.Visible = False
'    frameBorrarClientes.Visible = False
'    Me.frameConsolidado.Visible = False
'    FrameBalancesper.Visible = False
'    FramePersa.Visible = False
'    FrameAce.Visible = False
'    frameExploCon.Visible = False
'    FrameBalPersoConso.Visible = False
    
    Select Case Opcion
    Case 1, 34                  '34: Legalizacion
        'Listado de EXTRACTOS DE cuentas
        Me.frameListadoCuentas.visible = True
        If EjerciciosCerrados Then
            I = -1
        Else
            I = 0
        End If
        
        txtPag2(1).Text = ""
        If Opcion = 1 Then
            Text3(0).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
            Text3(1).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
            Text3(2).Text = Format(Now, "dd/mm/yyyy")
        Else
            Text3(2).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            Text3(0).Text = RecuperaValor(Legalizacion, 2)     'Inicio
            Text3(1).Text = RecuperaValor(Legalizacion, 3)     'Fin
        End If
        
        txtPag2(1).visible = Opcion = 1
        ImgAyuda(1).visible = Opcion = 1
        Label4(121).visible = Opcion = 1
        
        W = Me.frameListadoCuentas.Width
        H = Me.frameListadoCuentas.Height
        Combo1.ListIndex = 0
        Combo7.ListIndex = 0
        frameListadoCuentas.Enabled = Opcion = 1
    Case 2
        'Listado de cuentas
        Me.FrameCuentas.visible = True
        W = Me.FrameCuentas.Width
        H = Me.FrameCuentas.Height
        PonerNiveles
    Case 3
        'Listado de asientos
        frameAsiento.visible = True
        W = frameAsiento.Width
        H = frameAsiento.Height
    Case 4
        If EjerciciosCerrados Then
            I = -1
        Else
            I = 0
        End If
        Text3(6).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
        Text3(3).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
        frameCtaConcepto.visible = True
        W = frameCtaConcepto.Width
        H = frameCtaConcepto.Height
    Case 5, 36, 41  '36: Legalizacion Bal sumas.
        'Balance de suma y saldos
        CargarComboFecha
        If EjerciciosCerrados Then
            I = -1
        Else
            I = 0
        End If
        
        Me.chkBalIncioEjercicio.visible = Not EjerciciosCerrados
        frameBalance.visible = True
        W = frameBalance.Width
        H = frameBalance.Height
        chkResetea6y7.visible = False
        pb2.visible = False
        Combo3.ListIndex = 0
        Text1(0).Text = "Balance de sumas y saldos"
        If Opcion = 5 Then
            'Fecha informe
            Text3(7).Text = Format(Now, "dd/mm/yyyy")
            'Fecha inicial
            cmbFecha(0).ListIndex = Month(vParam.fechaini) - 1
            cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1
            txtAno(0).Text = Year(vParam.fechaini) + I
            txtAno(1).Text = Year(vParam.fechafin) + I
            
        Else
            
            'Legalizacion
            Text3(7).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
                
            txtAno(0).Text = Year(CDate(RecuperaValor(Legalizacion, 2)))     'Inicio
            txtAno(1).Text = Year(CDate(RecuperaValor(Legalizacion, 3)))     'Fin
            
            cmbFecha(0).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 2))) - 1
            cmbFecha(1).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 3))) - 1
            
            'Solo check enivado
            Cont = Val(RecuperaValor(Legalizacion, 4))
            For I = 1 To 10
                If I = Cont Then
                    Check2(I).Value = 1
                Else
                    Check2(I).Value = 0
                End If
            Next
            
            
            'Si la opcion es 41  Es el inventario final
            If Opcion = 41 Then
                Text1(0).Text = "Inventario final cierre."
                Cad = "5"
                For I = 2 To vEmpresa.DigitosUltimoNivel
                    Cad = Cad & "9"
                Next
                txtCta(7).Text = Cad
            End If
        End If
        frameBalance.Enabled = Opcion = 5
        
    Case 6, 35     'LEgalizacion Libros. Inventario Incial
        'Reemision de diario
        frameDiarioHco.visible = True
        W = frameDiarioHco.Width
        H = frameDiarioHco.Height
        If EjerciciosCerrados Then
            I = -1
        Else
            I = 0
        End If
        If Opcion = 6 Then
            Text3(4).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
            Text3(5).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
            Text3(8).Text = Format(Now, "dd/mm/yyyy")
            txtReemisionDiario.Text = "REEMISION DEL DIARIO"
        Else
            'Inventario incial legalizacion libros
            Text3(8).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            Text3(4).Text = RecuperaValor(Legalizacion, 2)     'Inicio
            Text3(5).Text = RecuperaValor(Legalizacion, 2)     'Pongo la misma k inicio
            'ASienot apertura
            txtAsiento(0).Text = "1": txtAsiento(1).Text = "1"
            txtReemisionDiario.Text = "INVENTARIO INICIAL"
        End If
        frameDiarioHco.Enabled = Opcion = 6
    Case 7
        'Explotacion
        CargarComboFecha
        If EjerciciosCerrados Then
            I = -1
            txtAno(4).Text = Year(vParam.fechafin) - 1
            cmbFecha(2).ListIndex = Month(vParam.fechafin) - 1
        Else
            I = 0
            txtAno(4).Text = Year(DateAdd("yyyy", I, Now))
            cmbFecha(2).ListIndex = Month(DateAdd("yyyy", I, Now)) - 1
        End If
        
   
        frameExplotacion.visible = True
        W = frameExplotacion.Width
        H = frameExplotacion.Height + 120
        Text3(9).Text = Format(Now, "dd/mm/yyyy")
    Case 8, 37, 53    '37: Presenacion telematica
                      '53 CONSOLIDADO
                      
                      
        If Opcion = 53 Then
            Label2(5).Caption = "List. Fact. Cliente consolidado"
        Else
            Label2(5).Caption = "Listado facturas cliente"
        End If
                      
        txtPag2(2).Text = ""
                      
        'Listado de facturas CLIENTES
        Frame5.visible = Opcion = 13
        FrameClientesCons.visible = Opcion = 53
        
        
        frameListFacCli.visible = True
        W = frameListFacCli.Width
        H = frameListFacCli.Height
                
        If Opcion <> 37 Then
            Text3(10).Text = Format(vParam.fechaini, "dd/mm/yyyy")
            Text3(11).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        Else
            Me.ChkListFac(4).Value = 0
            Text3(31).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            Text3(10).Text = RecuperaValor(Legalizacion, 2)     'Inicio
            Text3(11).Text = RecuperaValor(Legalizacion, 3)
        End If
            
        optListFac(2).visible = False
        ChkListFac(1).Value = 0
        SQL = "Nº Factura"
        RC = "Fecha "
        Label4(20).Caption = "Fecha "
        optListFac(0).Caption = SQL
        optListFac(1).Caption = RC
    
        optListFac(3).visible = vParam.Constructoras
        ChkListFac(1).visible = vParam.Constructoras
    
        If Opcion = 53 Then
            ChkListFac(1).visible = False
            ChkListFac(0).visible = False
            PonerEmpresaSeleccionEmpresa 9
        End If
    Case 9
        Me.FramePresu.visible = True
        W = FramePresu.Width
        H = FramePresu.Height

    Case 10
        pb4.visible = False
        PonerNiveles
        Me.FrameBalPresupues.visible = True
        W = FrameBalPresupues.Width
        H = FrameBalPresupues.Height
        
    Case 11
        frameIVA.visible = True
        W = frameIVA.Width
        H = frameIVA.Height
        Text3(12).Text = Format(Now, "dd/mm/yyyy")
        Text3(13).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Text3(14).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        cargacomboiva Combo4
        PonerEmpresaSeleccionEmpresa 0
        cmdCertIVA.Enabled = vUsu.Nivel <= 2
        
    Case 12
        'Liquidacion IVA
        Me.FrameLiq.visible = True
        W = FrameLiq.Width
        H = FrameLiq.Height
        Text3(33).Text = Format(Now, "dd/mm/yyyy")
        PonValoresLiquidacion
        PonerEmpresaSeleccionEmpresa 1
        Label13.visible = False
        'No dejamos k marque definitiva
        chkLiqDefinitiva.Enabled = vUsu.Nivel < 2
        
        'MODELO 303      Enero 2009
        '--------------------------
        Frame11.visible = False
        optModeloLiq(4).Value = True
        'ModeloIva True
    Case 13, 38, 52
        
        If Opcion = 52 Then
            Label2(22).Caption = "Fact. proveedores consolidado"
            PonerEmpresaSeleccionEmpresa 8
        Else
            Label2(22).Caption = "Listado facturas proveedores"
        End If
        
        FrameFactCons.visible = Opcion = 52
        txtPag2(3).Text = ""
        'FACTURAS PROVEEDORES
        FrameListFactP.visible = True
        W = FrameListFactP.Width
        H = FrameListFactP.Height
        Frame9.visible = vParam.Constructoras
        If Opcion <> 38 Then
            Text3(29).Text = Format(vParam.fechaini, "dd/mm/yyyy")
            Text3(30).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        
        Else
            Me.ChkListFac(5).Value = 0
            Text3(32).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            Text3(29).Text = RecuperaValor(Legalizacion, 2)     'Inicio
            Text3(30).Text = RecuperaValor(Legalizacion, 3)
        
        End If
        
        Combo8.visible = False
        If Opcion = 13 Then
            Combo8.visible = True
            cargacomboiva Combo8
        End If
        
        
        ChkListFac(2).Value = 0
        If vParam.Constructoras Then
            If Opcion = 13 Then ChkListFac(2).Value = 1
        End If
        optMostrarFecha(2).visible = vParam.Constructoras
        ChkListFac(2).visible = Opcion <> 52
        ChkListFac(3).visible = Opcion <> 52
        
        
    Case 14, 32   'El 32 es la impresion para el modelo de legaliza libros
        'Diario oficial
        Me.frameLibroDiario.visible = True
        W = frameLibroDiario.Width
        H = frameLibroDiario.Height
        If EjerciciosCerrados Then
            I = -1
        Else
            I = 0
        End If
        If Opcion = 14 Then
            Text3(17).Text = Format(Now, "dd/mm/yyyy")
            Text3(15).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
            Text3(16).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
            
        Else
            'Legalizacion
            Text3(17).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            Text3(15).Text = RecuperaValor(Legalizacion, 2)     'Inicio
            Text3(16).Text = RecuperaValor(Legalizacion, 3)     'Fin
        End If
        frameLibroDiario.Enabled = Opcion = 14
        I = Opcion
        If Opcion = 14 Then AjustaBotonCancelarAccion
    Case 15
        frameCCostSaldos.visible = True
        W = frameCCostSaldos.Width
        H = frameCCostSaldos.Height
        QueCombosFechaCargar "3|4|"
        cmbFecha(3).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(4).ListIndex = Month(vParam.fechafin) - 1
        txtAno(5).Text = Year(vParam.fechaini)
        txtAno(6).Text = Year(vParam.fechafin)
        
        
    Case 16
        Label15.Caption = ""
        frameCtaExpCC.visible = True
        W = frameCtaExpCC.Width
        H = frameCtaExpCC.Height
        QueCombosFechaCargar "5|6|7|"
        cmbFecha(5).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(6).ListIndex = Month(vParam.fechafin) - 1
        txtAno(7).Text = Year(vParam.fechaini)
        txtAno(8).Text = Year(vParam.fechafin)
        
        
    Case 17
        Label2(26).Caption = ""
        frameCCxCta.visible = True
        W = frameCCxCta.Width
        H = frameCCxCta.Height
        QueCombosFechaCargar "8|9|10|"
        cmbFecha(8).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(9).ListIndex = Month(vParam.fechafin) - 1
        txtAno(9).Text = Year(vParam.fechaini)
        txtAno(10).Text = Year(vParam.fechafin)
        
    Case 18, 33   '33: Legalizacion libros
        'LIBRO; Diario; RESUMEN
        Label2(25).Caption = ""
        frameResumen.visible = True
        W = frameResumen.Width
        H = frameResumen.Height

        QueCombosFechaCargar "11|12|"
        PonerNiveles
        If EjerciciosCerrados Then
            I = -1
        Else
            I = 0
        End If
        
        If Opcion = 18 Then
            Text3(18).Text = Format(Now, "dd/mm/yyyy")
            'Fecha inicial
            txtAno(11).Text = Year(vParam.fechaini) + I
            txtAno(12).Text = Year(vParam.fechafin) + I
        Else
            'legalizacion libros
            Text3(18).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            

            txtAno(11).Text = Year(CDate(RecuperaValor(Legalizacion, 2)))     'Inicio
            txtAno(12).Text = Year(CDate(RecuperaValor(Legalizacion, 3)))     'Fin
            
            cmbFecha(11).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 2))) - 1
            cmbFecha(12).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 3))) - 1
            'Pongo la marca
            Cont = RecuperaValor(Legalizacion, 4) 'Nivel
            For I = 1 To 9
                If I = Cont Then
                    ChkNivelRes(I).Value = 1
                Else
                    ChkNivelRes(I).Value = 0
                End If
            Next I
        End If
        frameResumen.Enabled = Opcion = 18
    Case 19
        'DetalleExplotacion
        frameccporcta.visible = True
        Label2(27).visible = False
        pb7.visible = False
        W = frameccporcta.Width
        H = frameccporcta.Height + 120
        Text3(19).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Text3(20).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    Case 20
        'Modelo IVA 347
        Frame347.visible = True
        OptProv(2).visible = vParam.Constructoras
        chk347(0).visible = vParam.Constructoras
        If vParam.Constructoras Then
            chk347(0).Value = 1
        Else
            chk347(0).Value = 0
        End If
        'Si no es agencia de viajes entonces NO lleva se pone la opcion
        chk347(2).Value = 1
        chk347(3).Value = 0
        chk347(2).visible = vParam.AgenciaViajes
        chk347(3).visible = vParam.AgenciaViajes
        If vParam.AgenciaViajes Then chk347(3).Value = 1
            
        W = Frame347.Width
        H = Frame347.Height
        Text3(21).Text = "01/01/" & Year(vParam.fechaini)
        Text3(22).Text = "31/12/" & Year(vParam.fechaini)
        PonerEmpresaSeleccionEmpresa 2
        Text347(1).Text = Format(vParam.limimpcl, FormatoImporte)
        Label2(30).Caption = ""
        Label2(31).Caption = ""
        Combo5.ListIndex = 1
    Case 21
        'Cta explotacion comparativa
        frameComparativo.visible = True
        W = frameComparativo.Width
        H = frameComparativo.Height
        QueCombosFechaCargar "13|"
        If EjerciciosCerrados Then
            FechaIncioEjercicio = UltimaFechaHcoCabapu
        Else
            FechaIncioEjercicio = vParam.fechaini
        End If
        txtAno(13).Text = Year(FechaIncioEjercicio)
        PonerNiveles
    Case 22, 23
        'Borre facturas cli/proveed
        frameBorrarClientes.visible = True
        W = frameBorrarClientes.Width
        H = frameBorrarClientes.Height
        FrameTapa.visible = Opcion = 23
        Me.FrBorrePorEjercicios.visible = True
        Label2(16).Caption = "Borre registro "
        If Opcion = 22 Then
            Label2(16).Caption = Label2(16).Caption & " clientes"
            Label2(16).ForeColor = &H800000
        Else
            Label2(16).Caption = Label2(16).Caption & " proveedores"
            Label2(16).ForeColor = &H80&
        End If
        PonerFechaBorre
        
        
    Case 24
        PonerNiveles
        'Balance consolidado de empresas
        frameConsolidado.visible = True
        W = frameConsolidado.Width
        H = frameConsolidado.Height
        pb9.visible = False
        Label21.Caption = ""
        Label21.visible = False
        PonerEmpresaSeleccionEmpresa 3
        QueCombosFechaCargar "14|15|"
        cmbFecha(14).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(15).ListIndex = Month(vParam.fechafin) - 1
        txtAno(14).Text = Year(vParam.fechaini)
        txtAno(15).Text = Year(vParam.fechafin)
        
    Case 25, 26, 27, 39, 40
        'Balances personalizados
        chkApaisado.Value = Abs(vParam.NuevoPlanContable)
        FrameBalancesper.visible = True
        H = FrameBalancesper.Height + 120
        W = FrameBalancesper.Width
        QueCombosFechaCargar "16|17|"
        If Opcion < 39 Then
            cmbFecha(16).ListIndex = Month(vParam.fechafin) - 1
            cmbFecha(17).ListIndex = Month(vParam.fechafin) - 1
            txtAno(16).Text = Year(vParam.fechafin)
            txtAno(17).Text = Year(vParam.fechafin) - 1
            Text3(25).Text = Format(vParam.fechafin, "dd/mm/yyyy")
            If Opcion > 25 Then PonerBalancePredeterminado
                
                
                
        Else
            'LEGALIZA legaliza LE-GA-LI-ZACION
                
            PonerBalancePredeterminado
            
            Text3(25).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
                
            'txtAno(0).Text = Year(CDate(RecuperaValor(Legalizacion, 2)))     'Inicio
            txtAno(16).Text = Year(CDate(RecuperaValor(Legalizacion, 3)))     'Fin
            
            'cmbFecha(0).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 2))) - 1
            cmbFecha(16).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 3))) - 1
            
            Cad = RecuperaValor(Legalizacion, 4)
            If Val(Cad) = 0 Then
                chkBalPerCompa.Value = 0
            Else
                txtAno(17).Text = Val(txtAno(16).Text) - 1
                cmbFecha(17).ListIndex = cmbFecha(16).ListIndex
                chkBalPerCompa.Value = 1
            End If
        End If
                
            
    Case 28
        'Modelo IVA 349
        Frame349.visible = True
        W = Frame349.Width
        H = Frame349.Height + 120
        Text3(26).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Text3(27).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        Text3(35).Text = Format(Now, "dd/mm/yyyy")
        PonerEmpresaSeleccionEmpresa 4
        Me.chk349.Value = 1
        Combo6.ListIndex = 0
        LeerSerieFacturas True
        'Cargo el combo6 con los meses o con el trimestre
        CargarComboPeriodo
    Case 29
        'Traspaso PERSA
        Me.FramePersa.visible = True
        W = FramePersa.Width
        H = FramePersa.Height
        lblPersa.Caption = ""
        lblPersa2.Caption = ""
    
    Case 30
        'Traspaso ACE
        Me.FrameAce.visible = True
        W = FrameAce.Width
        H = FrameAce.Height
        PonerEmpresaSeleccionEmpresa 5
        QueCombosFechaCargar "18|"
        CargarComboFecha
        Text1(1).Text = "A12345"
    Case 31
        frameExploCon.visible = True
        W = frameExploCon.Width
        H = frameExploCon.Height
        PonerEmpresaSeleccionEmpresa 6
        QueCombosFechaCargar "19|"
        CargarComboFecha
        txtAno(18).Text = Year(vParam.fechaini)
        Text3(28).Text = Format(Now, "dd/mm/yyyy")
        
    Case 50, 51
        'Balacnces configurables personalizados:  CONSOLIDADO
        FrameBalPersoConso.visible = True
        H = FrameBalPersoConso.Height
        W = FrameBalPersoConso.Width
        QueCombosFechaCargar "20|21|"
        cmbFecha(20).ListIndex = Month(vParam.fechafin) - 1
        cmbFecha(21).ListIndex = Month(vParam.fechafin) - 1
        txtAno(20).Text = Year(vParam.fechafin)
        txtAno(21).Text = Year(vParam.fechafin) - 1
        Text3(34).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        PonerBalancePredeterminado
        PonerEmpresaSeleccionEmpresa 7
        
        
    Case 54
        'Evolucion mensual de saldos
        PonerNiveles
        CargaComboEjercicios 0
        Me.FrameEvolSaldo.visible = True
        H = FrameEvolSaldo.Height
        W = FrameEvolSaldo.Width
        Label2(29).Caption = ""
        
        
    Case 55, 56
        
        Me.FrameRela_x_Cta.visible = True
        H = FrameRela_x_Cta.Height + 120
        W = FrameRela_x_Cta.Width + 60
        If Opcion = 55 Then
            Label4(114).Caption = "Cta clientes"
            Label2(23).Caption = "Relación de clientes por cta de ventas"
            
        Else
            Label4(114).Caption = "Cta proveedores"
            Label2(23).Caption = "Relación de prov. por cta de gastos"
        End If
        FrameProvxGast.visible = Not Opcion = 55
        
        FechaIncioEjercicio = DateAdd("m", -1, Now)
        If FechaIncioEjercicio < vParam.fechaini Then
            FechaIncioEjercicio = vParam.fechafin
        Else
            I = DiasMes(Month(FechaIncioEjercicio), Year(FechaIncioEjercicio))
            FechaIncioEjercicio = CDate(I & "/" & Month(FechaIncioEjercicio) & "/" & Year(FechaIncioEjercicio))
        End If
        
        Text3(36).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Text3(37).Text = Format(FechaIncioEjercicio, "dd/mm/yyyy")
        Label2(24).Caption = ""
        
        
    Case 57
        txtNumBal(2).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        TextDescBalance(2).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        H = FrameCopyBalan.Height
        W = FrameCopyBalan.Width
        FrameCopyBalan.visible = True
    End Select
    HanPulsadoSalir = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    I = Opcion
    If Opcion = 23 Then I = 22
    If Opcion = 26 Or Opcion = 27 Or Opcion = 39 Or Opcion = 40 Then I = 25
    If Opcion = 51 Then I = 50
    If Opcion = 41 Then I = 5
    If Opcion = 52 Then I = 13
    If Opcion = 53 Then I = 8
    If Opcion = 56 Then I = 55
    
    'Legalizacion
    HanPulsadoSalir = True
    
    If Opcion < 32 Or Opcion > 38 Then
    
        Me.cmdCanListExtr(I).Cancel = True
        
        'Ajustaremos el boton para cancelar algunos de los listados k mas puedan costar
        AjustaBotonCancelarAccion
        cmdCancelarAccion.visible = False
        cmdCancelarAccion.ZOrder 0
    
    End If
    Me.Width = W + 240
    Me.Height = H + 400
    
    'Añadimos ejercicios cerrados
    If EjerciciosCerrados Then Caption = Caption & "    EJERC. TRASPASADOS"
End Sub

Private Sub AjustaBotonCancelarAccion()
On Error GoTo EAj
    Me.cmdCancelarAccion.Top = cmdCanListExtr(I).Top
    Me.cmdCancelarAccion.Left = cmdCanListExtr(I).Left + 60
    cmdCancelarAccion.Width = cmdCanListExtr(I).Width
    cmdCancelarAccion.Height = cmdCanListExtr(I).Height + 30
    Exit Sub
EAj:
    MuestraError Err.Number, "Ajuste BOTON cancelar"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not HanPulsadoSalir Then Cancel = 1
    If Opcion = 28 Then LeerSerieFacturas False
    Legalizacion = ""
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Me.txtNumBal(RC).Text = RecuperaValor(CadenaDevuelta, 1)
    TextDescBalance(RC).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    I = Val(Me.imgCCost(0).Tag)
    Me.txtCCost(I).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDCost(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtConcepto = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescConce = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    txtDiario(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescDiario(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmD = New frmTiposDiario
    RC = Index
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub

Private Sub Image10_Click()
    ConsolidadoEmpresas List7
End Sub

Private Sub Image11_Click()
    ConsolidadoEmpresas List8
End Sub

Private Sub Image12_Click()
    ConsolidadoEmpresas List9
End Sub

Private Sub Image13_Click()
    ConsolidadoEmpresas List10
End Sub

Private Sub Image2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    RC = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub


Private Sub Image4_Click()
'    CadenaDesdeOtroForm = ""
'    frmMensajes.Opcion = 4
'    frmMensajes.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        Cont = RecuperaValor(CadenaDesdeOtroForm, 1)
'        List1.Clear
'        If Cont = 0 Then Exit Sub
'        For I = 1 To Cont
'            List1.AddItem RecuperaValor(CadenaDesdeOtroForm, I + 1)
'        Next I
'        For I = 0 To Cont - 1
'            List1.ItemData(I) = RecuperaValor(CadenaDesdeOtroForm, Cont + I + 2)
'        Next I
'    End If
    ConsolidadoEmpresas List1
End Sub


Private Sub Image5_Click()
'    CadenaDesdeOtroForm = ""
'    frmMensajes.Opcion = 4
'    frmMensajes.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        Cont = RecuperaValor(CadenaDesdeOtroForm, 1)
'        List2.Clear
'        If Cont = 0 Then Exit Sub
'        For I = 1 To Cont
'            List2.AddItem RecuperaValor(CadenaDesdeOtroForm, I + 1)
'        Next I
'        For I = 0 To Cont - 1
'            List2.ItemData(I) = RecuperaValor(CadenaDesdeOtroForm, Cont + I + 2)
'        Next I
'    End If
    ConsolidadoEmpresas List2
End Sub

Private Sub Image6_Click()
'   CadenaDesdeOtroForm = ""
'    frmMensajes.Opcion = 4
'    frmMensajes.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        Cont = RecuperaValor(CadenaDesdeOtroForm, 1)
'        List3.Clear
'        If Cont = 0 Then Exit Sub
'        For I = 1 To Cont
'            List3.AddItem RecuperaValor(CadenaDesdeOtroForm, I + 1)
'        Next I
'        For I = 0 To Cont - 1
'            List3.ItemData(I) = RecuperaValor(CadenaDesdeOtroForm, Cont + I + 2)
'        Next I
'    End If
    ConsolidadoEmpresas List3
End Sub

Private Sub Image7_Click()
'    CadenaDesdeOtroForm = ""
'    frmMensajes.Opcion = 4
'    frmMensajes.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        Cont = RecuperaValor(CadenaDesdeOtroForm, 1)
'        List4.Clear
'        If Cont = 0 Then Exit Sub
'        For I = 1 To Cont
'            List4.AddItem RecuperaValor(CadenaDesdeOtroForm, I + 1)
'        Next I
'        For I = 0 To Cont - 1
'            List4.ItemData(I) = RecuperaValor(CadenaDesdeOtroForm, Cont + I + 2)
'        Next I
'    End If
    ConsolidadoEmpresas List4
End Sub

Private Sub Image8_Click()
'    CadenaDesdeOtroForm = ""
'    frmMensajes.Opcion = 4
'    frmMensajes.Show vbModal
'    If CadenaDesdeOtroForm <> "" Then
'        Cont = RecuperaValor(CadenaDesdeOtroForm, 1)
'        List5.Clear
'        If Cont = 0 Then Exit Sub
'        For I = 1 To Cont
'            List5.AddItem RecuperaValor(CadenaDesdeOtroForm, I + 1)
'        Next I
'        For I = 0 To Cont - 1
'            List5.ItemData(I) = RecuperaValor(CadenaDesdeOtroForm, Cont + I + 2)
'        Next I
'    End If
    ConsolidadoEmpresas List5
End Sub



Private Sub ConsolidadoEmpresas(ByRef L As ListBox)
    CadenaDesdeOtroForm = ""
    frmMensajes.Opcion = 4
    frmMensajes.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Cont = RecuperaValor(CadenaDesdeOtroForm, 1)
        L.Clear
        If Cont = 0 Then Exit Sub
        For I = 1 To Cont
            L.AddItem RecuperaValor(CadenaDesdeOtroForm, I + 1)
        Next I
        For I = 0 To Cont - 1
            L.ItemData(I) = RecuperaValor(CadenaDesdeOtroForm, Cont + I + 2)
        Next I
    End If
End Sub


Private Sub Image9_Click()
    ConsolidadoEmpresas List6
    chkAce(10).Enabled = Not (List6.ListCount > 1)
    I = DevuelveDigitosNivelAnterior
    chkAce(I).Value = Val(Abs((List6.ListCount > 1)))
    chkAce(10).Value = Val(Abs((List6.ListCount = 1)))
End Sub

Private Sub ImgAyuda_Click(Index As Integer)
    Cad = String(60, "*")
    Cad = Cad & "    " & vbCrLf
    Select Case Index
    Case 0
        
        SQL = Cad & vbCrLf & vbCrLf & "Se utiliza el campo ""total factura"" tanto en clientes como en proveedores"
        SQL = SQL & vbCrLf & vbCrLf & Cad
    Case 2, 3
        SQL = Cad & vbCrLf & vbCrLf & "Mostrará todas las cuentas en los desde/hasta que sean de este NIF"
        SQL = SQL & vbCrLf & vbCrLf & "Si indica tipo de IVA solo mostrará la de ese tipo de iva"
        SQL = SQL & vbCrLf & vbCrLf & Cad
        
    End Select
    MsgBox SQL, vbInformation
End Sub

Private Sub imgCCost_Click(Index As Integer)
    imgCCost(0).Tag = Index
    Set frmCC = New frmCCoste
    frmCC.DatosADevolverBusqueda = "0|1|"
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub ImgConcepto_Click()
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|1|"
    frmCon.Show vbModal
    Set frmCon = Nothing
End Sub

Private Sub ImgNumBal_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
'            Cad = Cad & "|"
'            Cad = Cad & mTag.Columna & "|"
'            Cad = Cad & mTag.TipoDato & "|"
'            Cad = Cad & AnchoPorcentaje & "·"
    frmB.VCampos = "Codigo|numbalan|N|10·" & "Descripcion|nombalan|T|60·"
    frmB.vTabla = "sbalan"
    frmB.vSQL = ""
    CadenaDesdeOtroForm = ""
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Balances disponibles"
    frmB.vSelElem = 0
    RC = Index
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
End Sub






Private Sub List8_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub mnPrueba_Click()
    MsgBox "prueab"
End Sub

Private Sub opt349_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub optAce_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub optAceAcum_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub optEvolSald_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub optEvolSald_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub optCta_x_gastos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Option1_Click(Index As Integer)
    Me.Frame2.visible = Option1(0).Value
End Sub



Private Sub optListFac_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub optListFacP_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub OptProv_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub optSelFech_Click(Index As Integer)
Dim Cad As String
    If optSelFech(0).Value Then
        Cad = "Fecha recepción"
        optMostrarFecha(2).Caption = "F. liquidación"
    Else
        Cad = "Fecha liquidación"
        optMostrarFecha(2).Caption = "F. recepción"
    End If
    Label4(90).Caption = Cad
    optListFacP(2).Caption = Cad
End Sub

Private Sub optSelFech_KeyPress(Index As Integer, KeyAscii As Integer)
 ListadoKEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    PonFoco Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub



Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = 112 Then HacerF1
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub







Private Sub Text347_GotFocus(Index As Integer)
    PonFoco Text347(Index)
End Sub

Private Sub Text347_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text347_LostFocus(Index As Integer)
Dim Mal As Boolean
    If Index = 1 Then
        Text347(Index).Text = Trim(Text347(Index).Text)
        If Text347(Index).Text = "" Then Exit Sub
        Mal = False
        If Not EsNumerico(Text347(Index).Text) Then Mal = True
        
        If Not Mal Then Mal = Not CadenaCurrency(Text347(Index).Text, Importe)
        
        If Mal Then
            Text347(Index).Text = ""
            PonerFoco Text347(Index)
        Else
            Text347(Index).Text = Format(Importe, FormatoImporte)
        End If
    End If
    
End Sub

Private Sub Text4_GotFocus(Index As Integer)
    PonFoco Text4(Index)
End Sub

Private Sub Text4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub Text4_LostFocus(Index As Integer)
    Text4(Index).Text = Trim(Text4(Index).Text)
End Sub

Private Sub txtAno_GotFocus(Index As Integer)
PonFoco txtAno(Index)
End Sub

Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtAno_KeyPress(Index As Integer, KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub txtAno_LostFocus(Index As Integer)
txtAno(Index).Text = Trim(txtAno(Index).Text)
If txtAno(Index).Text = "" Then Exit Sub
If Not IsNumeric(txtAno(Index).Text) Then
    MsgBox "Campo año debe ser numérico", vbExclamation
    txtAno(Index).SetFocus
Else
    If Index = 0 Then ComprobarFechasBalanceQuitar6y7
End If
End Sub

Private Sub txtAS_GotFocus(Index As Integer)
    PonFoco txtAS(Index)
End Sub

Private Sub txtAS_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtAS_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtAS_LostFocus(Index As Integer)
    txtAS(Index).Text = Trim(txtAS(Index).Text)
    If txtAS(Index).Text = "" Then Exit Sub
    SQL = DevuelveDesdeBD("nomaspre", "cabasipre", "numaspre", txtAS(Index).Text, "N")
    If SQL = "" Then
        MsgBox "No se ha encontrado el asiento predefinido.", vbExclamation
        txtDesAs(Index).Text = ""
        txtAS(Index).Text = ""
        txtAS(Index).SetFocus
    Else
        txtDesAs(Index).Text = SQL
    End If
End Sub



Private Sub txtAsiento_GotFocus(Index As Integer)
PonFoco txtAsiento(Index)
End Sub

Private Sub txtAsiento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtAsiento_KeyPress(Index As Integer, KeyAscii As Integer)
ListadoKEYpress KeyAscii
End Sub

Private Sub txtAsiento_LostFocus(Index As Integer)
txtAsiento(Index).Text = Trim(txtAsiento(Index).Text)
If txtAsiento(Index).Text <> "" Then
    If Not IsNumeric(txtAsiento(Index).Text) Then
        MsgBox "El asiento debe de ser numérico: " & txtAsiento(Index).Text, vbExclamation
        txtAsiento(Index).Text = ""
    End If
End If
End Sub



Private Sub txtCCost_GotFocus(Index As Integer)
    PonFoco txtCCost(Index)
End Sub

Private Sub txtCCost_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtCCost_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtCCost_LostFocus(Index As Integer)
    txtCCost(Index).Text = Trim(txtCCost(Index).Text)
    If txtCCost(Index).Text = "" Then
        Me.txtDCost(Index).Text = ""
        Exit Sub
    End If
    
    SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtCCost(Index).Text, "T")
    If SQL = "" Then
        If Index > 7 Then
            MsgBox "Centro de coste NO encontrado: " & txtCCost(Index).Text, vbExclamation
            txtCCost(Index).Text = ""
            txtCCost(Index).SetFocus
        End If
    Else
        txtCCost(Index).Text = UCase(txtCCost(Index).Text)
    End If
    Me.txtDCost(Index).Text = SQL
End Sub

Private Sub txtConcepto_GotFocus()
    PonFoco txtConcepto
End Sub

Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtConcepto_LostFocus()
    txtConcepto.Text = Trim(txtConcepto.Text)
    Me.txtDescConce.Text = ""
    If txtConcepto.Text = "" Then Exit Sub

    If Not IsNumeric(txtConcepto.Text) Then
        MsgBox "El concepto debe ser numérico: " & txtConcepto.Text, vbExclamation
        txtConcepto.Text = ""
        txtConcepto.SetFocus
    End If
    
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto.Text, "N")
    If SQL = "" Then
        MsgBox "No se ha encontrado el concepto: " & txtConcepto.Text, vbExclamation
        txtConcepto.Text = ""
        txtConcepto.SetFocus
    Else
        txtDescConce.Text = SQL
    End If
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 112 Then
        HacerF1
    Else
        If KeyCode = 107 Or KeyCode = 187 Then
            KeyCode = 0
            txtCta(Index).Text = ""
            Image3_Click Index
        End If
    End If
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
    
    Select Case Index
    Case 0 To 7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 23 To 30
        'NO hace falta que sean de ultimo nivel
        Cta = (txtCta(Index).Text)
                                '********
        B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
        If B = 0 Then
            MsgBox "NO existe la cuenta: " & txtCta(Index).Text, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
        Else
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            If B = 1 Then
                DtxtCta(Index).Tag = ""
            Else
                DtxtCta(Index).Tag = SQL
            End If
            Hasta = -1
            If Index = 6 Then
                Hasta = 7
            Else
                If Index = 0 Then
                    Hasta = 1
                Else
                    If Index = 5 Then
                        Hasta = 4
                    Else
                        If Index = 23 Then Hasta = 24
                    End If
                End If
                
            End If
                
                'If txtCta(1).Text = "" Then 'ANTES solo lo hacia si el texto estaba vacio
            If Hasta >= 0 Then
                txtCta(Hasta).Text = txtCta(Index).Text
                DtxtCta(Hasta).Text = DtxtCta(Index).Text
            End If
        End If
    Case Else
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, SQL) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If
    End Select
End Sub


Private Sub txtDiario_GotFocus(Index As Integer)
    PonFoco txtDiario(Index)
End Sub

Private Sub txtDiario_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
With txtDiario(Index)
    .Text = Trim(.Text)
    Me.txtDescDiario(Index).Text = ""
    If .Text = "" Then Exit Sub

    If Not IsNumeric(.Text) Then
        MsgBox "El diario debe ser numérico: " & .Text, vbExclamation
        .Text = ""
        .SetFocus
    End If
    
    SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", .Text, "N")
    If SQL = "" Then
        MsgBox "No se ha encontrado el diario: " & .Text, vbExclamation
        .Text = ""
        .SetFocus
    Else
        Me.txtDescDiario(Index).Text = SQL
    End If
End With
End Sub

Private Sub txtExplo_GotFocus(Index As Integer)
    PonFoco txtExplo(Index)
End Sub

Private Sub txtExplo_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtExplo_LostFocus(Index As Integer)
    txtExplo(Index).Text = Trim(txtExplo(Index).Text)
    If txtExplo(Index).Text = "" Then Exit Sub
    If Not IsNumeric(txtExplo(Index)) Then
        MsgBox "Los importes deben de ser numéricos: " & txtExplo(Index).Text, vbExclamation
        txtExplo(Index).Text = ""
        txtExplo(Index).SetFocus
        Exit Sub
    End If
    
        If InStr(1, txtExplo(Index).Text, ",") > 0 Then
            Cad = ImporteFormateado(txtExplo(Index).Text)
        Else
            Cad = CCur(TransformaPuntosComas(txtExplo(Index).Text))
        End If
        txtExplo(Index).Text = Cad
    
End Sub




Private Sub txtLibroOf_GotFocus(Index As Integer)
    PonFoco txtLibroOf(Index)
End Sub

Private Sub txtLibroOf_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtLibroOf_LostFocus(Index As Integer)
    txtLibroOf(Index).Text = Trim(txtLibroOf(Index).Text)
    If txtLibroOf(Index).Text = "" Then Exit Sub
    If Not IsNumeric(txtLibroOf(Index)) Then
        MsgBox "El campo debe ser numérico: " & txtLibroOf(Index).Text, vbExclamation
        txtLibroOf(Index).Text = ""
        txtLibroOf(Index).SetFocus
    End If
End Sub

Private Sub txtMes_GotFocus(Index As Integer)
    PonFoco txtMes(Index)
End Sub


Private Sub txtMes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtMes_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtMes_LostFocus(Index As Integer)
'Comprobar valores
    
    txtMes(Index).Text = Trim(txtMes(Index).Text)
    If txtMes(Index).Text <> "" Then
        If Not IsNumeric(txtMes(Index).Text) Then
            MsgBox "El campo no es válido: " & txtMes(Index).Text, vbExclamation
            txtMes(Index).Text = ""
            txtMes(Index).SetFocus
        End If
    End If
    
End Sub



Private Sub txtNpag_GotFocus()
    PonFoco txtNpag
End Sub

Private Sub txtNpag_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtNpag_LostFocus()
txtNpag.Text = Trim(txtNpag.Text)
If txtNpag.Text <> "" Then
    If Not IsNumeric(txtNpag.Text) Then
        MsgBox "Número de pagina no es un campo válido: " & txtNpag.Text, vbExclamation
        txtNpag.Text = ""
        txtNpag.SetFocus
    End If
End If
End Sub



Private Sub txtNpag2_GotFocus(Index As Integer)
    PonFoco txtNpag2(Index)
End Sub

Private Sub txtNpag2_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtNpag2_LostFocus(Index As Integer)
    With txtNpag2(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "Número de pagina no es un campo válido: " & .Text, vbExclamation
                .Text = ""
                .SetFocus
            End If
        End If
    End With
End Sub

Private Sub txtNumBal_GotFocus(Index As Integer)
    PonFoco txtNumBal(Index)
End Sub



Private Sub txtNumBal_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtNumBal_LostFocus(Index As Integer)
    SQL = ""
    With txtNumBal(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "Numero de balance debe de ser numérico: " & .Text, vbExclamation
                .Text = ""
            Else
                SQL = DevuelveDesdeBD("nombalan", "sbalan", "numbalan", .Text)
                If SQL = "" Then
                    MsgBox "El balance " & .Text & " NO existe", vbExclamation
                    .Text = ""
                End If
            End If
        End If
    End With
    TextDescBalance(Index).Text = SQL
End Sub

Private Sub txtNumFac_GotFocus(Index As Integer)
    PonFoco txtNumFac(Index)
End Sub

Private Sub txtNumFac_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtNumFac_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtNumFac_LostFocus(Index As Integer)
    With txtNumFac(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        If Not IsNumeric(.Text) Then
            MsgBox "Numero de factura debe de ser numérico: " & .Text, vbExclamation
            .Text = ""
            Exit Sub
        End If
    End With
End Sub



Private Sub txtNumRes_GotFocus(Index As Integer)
    PonFoco txtNumRes(Index)
End Sub

Private Sub txtNumRes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtNumRes_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtNumRes_LostFocus(Index As Integer)

txtNumRes(Index).Text = Trim(txtNumRes(Index).Text)
If txtNumRes(Index).Text = "" Then Exit Sub
If Not IsNumeric(txtNumRes(Index).Text) Then
    MsgBox "El campo tiene que ser numérico: " & txtNumRes(Index).Text, vbExclamation
    txtNumRes(Index).Text = ""
    txtNumRes(Index).SetFocus
    Exit Sub
End If
End Sub




Private Sub PonerNiveles()
Dim I As Integer
Dim J As Integer


    Frame2.visible = True
    Combo2.Clear
    Check1(10).visible = True
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        Cad = "Digitos: " & J
        Check1(I).visible = True
        Me.Check1(I).Caption = Cad
        
        'Para los de balance presupuestario
        Me.ChkCtaPre(I).visible = True
        Me.ChkCtaPre(I).Caption = Cad
        'para los de resumen dairio
        Me.ChkNivelRes(I).visible = True
        Me.ChkNivelRes(I).Caption = Cad
        
        'Evolucion de saldos
        ChkEvolSaldo(I).visible = True
        ChkEvolSaldo(I).Caption = Cad
        
        'Consolidado
        Me.ChkConso(I).visible = True
        Me.ChkConso(I).Caption = Cad
        
        chkcmp(I).Caption = Cad
        chkcmp(I).visible = True
        
        Combo2.AddItem "Nivel :   " & I
        Combo2.ItemData(Combo2.NewIndex) = J
    Next I
    For I = vEmpresa.numnivel To 9
        Check1(I).visible = False
        Me.ChkCtaPre(I).visible = False
        Me.ChkNivelRes(I).visible = False
        chkcmp(I).visible = False
        ChkConso(I).visible = False
        Me.ChkEvolSaldo(I).visible = False
    Next I
    
End Sub


Private Sub ImprimirListadoCuentas()
Dim MostrarAnterior As Byte
    'Resto parametros
    SQL = ""
    
    'Fechas intervalor
    If txtTitulo.Text = "" Then
        SQL = "EXTRACTOS DE CUENTAS"
    Else
        SQL = txtTitulo.Text
    End If
    SQL = "Titulo= """ & SQL & """|"
    
    SQL = SQL & "Fechas= ""Fechas:  desde " & Text3(0).Text & " hasta " & Text3(1).Text
    If Combo1.ListIndex > 1 Then
        SQL = SQL & "         "
        If Combo1.ListIndex = 2 Then
            'Solo kiere punteadas
            SQL = SQL & "PUNTEADOS"
        Else
            'Solo kiere PENDIENTES DE PUNTEAR
            SQL = SQL & "PENDIENTES"
        End If
    End If
    
    If txtPag2(1).Text <> "" Then SQL = SQL & "  NIF: " & txtPag2(1).Text
    
    SQL = SQL & """|"
    
    'Cuentas
    RC = ""
    If txtCta(0).Text <> "" Then RC = " desde " & txtCta(0).Text & " -" & DtxtCta(0).Text
    If txtCta(1).Text <> "" Then RC = RC & " hasta " & txtCta(1).Text & " -" & DtxtCta(1).Text
    If RC <> "" Then RC = "Cuentas: " & RC
    SQL = SQL & "Cuenta= """ & RC & """|"
    
    'Fecha impresion
    SQL = SQL & "FechaIMP= """ & Text3(2).Text & """|"
    
    'Numero página
    If txtPag2(0).Text <> "" Then
        RC = Val(txtPag2(0).Text) - 1
    Else
        RC = 0
    End If
    SQL = SQL & "NumPag= " & RC & "|"
    
    'Salto por cuenta
    If Check1(0).Value = 1 Then
        RC = "1"
        Else
        RC = "2"
    End If
    SQL = SQL & "Salto= " & RC & "|"
    
    
    'Veremos si la fecha coincide con la fecha de incio
    'MostrarAnterior
    MostrarAnterior = FechaInicioIGUALinicioEjerecicio(CDate(Text3(0).Text), EjerciciosCerrados)
    SQL = SQL & "MostrarAnterior= " & MostrarAnterior & "|"
    
    If Opcion <> 34 Then
        'Impresion normal
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 7
            .FormulaSeleccion = "{ado_lineas.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            If Combo1.ListIndex < 4 Then
                .Opcion = 8 + Combo1.ListIndex
            Else
                .Opcion = 81
            End If
            .Show vbModal
        End With
        
    Else
        GeneraLegalizaPRF SQL, 7
        CadenaDesdeOtroForm = "OK"
 
    End If
End Sub



Private Sub CargarComboFecha()
Dim J As Integer


QueCombosFechaCargar "0|1|2|"


'Y ademas deshabilitamos los niveles no utilizados por la aplicacion
For I = vEmpresa.numnivel To 9
    Check2(I).visible = False
    Me.chkCtaExplo(I).visible = False
    chkCtaExploC(I).visible = False
    chkAce(I).visible = False
Next I

For I = 1 To vEmpresa.numnivel - 1
    J = DigitosNivel(I)
    Check2(I).visible = True
    Check2(I).Caption = "Digitos: " & J
    chkCtaExplo(I).visible = True
    chkCtaExplo(I).Caption = "Digitos: " & J
    chkAce(I).visible = True
    chkAce(I).Caption = "Digitos: " & J
    chkCtaExploC(I).visible = True
    chkCtaExploC(I).Caption = "Digitos: " & J
Next I




'Cargamos le combo de resalte de fechas
Combo3.AddItem "Sin remarcar"
Combo3.ItemData(Combo3.NewIndex) = 1000
For I = 1 To vEmpresa.numnivel - 1
    Combo3.AddItem "Nivel " & I
    Combo3.ItemData(Combo3.NewIndex) = I
Next I
End Sub



Private Sub GeneraSQL(Busqueda As String, vOP As Integer)
Dim SQL As String
Dim nexo As String
Dim J As Integer
Dim wildcar As String
Dim DigiTNivel As Integer
Dim IndiceAnyo As Integer
Dim IndiceMes As Integer

    SQL = ""
    nexo = ""
    If vOP = -1 Then
        If Check2(10).Value Then
            SQL = "( apudirec = 'S')"
            nexo = " OR "
        End If
    End If
    For I = 1 To vEmpresa.numnivel - 1
        wildcar = ""
        
        If vOP = -1 Then
            'Balance normal
            If Check2(I).Value = 1 Then
                DigiTNivel = DigitosNivel(I)
                For J = 1 To DigiTNivel
                    wildcar = wildcar & "_"
                Next J
            End If
        Else
            'Balance consolidado
            If Me.ChkConso(I).Value = 1 Then
                DigiTNivel = DigitosNivel(I)
                For J = 1 To DigiTNivel
                    wildcar = wildcar & "_"
                Next J
            End If
        End If
        If wildcar <> "" Then
            SQL = SQL & nexo & " (cuentas.codmacta like '" & wildcar & "')"
            nexo = " OR "
            If SQL <> "" Then SQL = "(" & SQL & ")"
        End If
    Next I


'Nexo
    Cad = "SELECT cuentas.codmacta,nommacta From "
    If vOP >= 0 Then Cad = Cad & "Conta" & vOP & "."
    Cad = Cad & "cuentas as cuentas"
    
    
        
    
    
    'MODIFICACION DE 20 OCTUBRE 2005
    Cad = Cad & ","
    If vOP >= 0 Then Cad = Cad & "Conta" & vOP & "."
    Cad = Cad & "hsaldos"
    If EjerciciosCerrados Then Cad = Cad & "1"
    Cad = Cad & " as hs WHERE "
    Cad = Cad & "cuentas.codmacta = hs.codmacta"


    Cad = Cad & " AND "
    Cad = Cad & SQL
    If Busqueda <> "" Then Cad = Cad & Busqueda
    
    
    
    'modificacion 21 Nov 2008 . MAAAAAl para años partidos
    'Rehacemos en Marzo 2009
    If Opcion = 24 Then
        IndiceAnyo = 14
        IndiceMes = 14
    Else
        IndiceAnyo = 0  'Val(txtAno(0).Text)
        IndiceMes = 0 'Val(Me.cmbFecha(0).ListIndex + 1)
    End If
    
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'AÑOS NATURALES. Normal. No toco nada
        If Val(txtAno(IndiceAnyo).Text) > Year(vParam.fechaini) Then   'Pide en siguiente
            Cad = Cad & " AND (anopsald >= " & Year(vParam.fechaini) & " AND anopsald <= " & Year(vParam.fechafin) + 1 & ")"
        Else
            Cad = Cad & " AND (anopsald >= " & txtAno(IndiceAnyo).Text & " AND anopsald <= " & txtAno(IndiceAnyo + 1).Text & ")"
        End If
    
    Else
        'AÑOS PARTIDOS.
        'Si pide en ejercicio siguiente entonces hay que contemplar desde fechaini
        J = 0
        If Val(txtAno(IndiceAnyo).Text) > Year(vParam.fechafin) Then
            J = 1
        Else
            'MAYO 2009.
            
            
            'Año de fecha fin y mes mayor que fecha fin
            If Val(txtAno(IndiceAnyo).Text) = Year(vParam.fechafin) And (Me.cmbFecha(IndiceMes).ListIndex + 1) > Month(vParam.fechafin) Then J = 1
        End If
        
        
        
        'Siempre año partido
        If J = 0 Then
            
            
            
            'Buscamos mes/anyo para la fecha de inicio del balance
            If Me.cmbFecha(IndiceAnyo).ListIndex + 1 < Month(vParam.fechaini) Then
                'EL año es el anterior
                Cad = Cad & " AND ((anopsald = " & Val(txtAno(IndiceAnyo).Text) - 1 & " AND mespsald >= " & Month(vParam.fechaini) & ")"
                Cad = Cad & " OR (anopsald = " & txtAno(IndiceAnyo + 1).Text & " AND mespsald <= " & Month(vParam.fechafin) & "))"
            
            
            Else
                'Años partidos
                Cad = Cad & " AND ((anopsald = " & txtAno(IndiceAnyo).Text & " AND mespsald >= " & Month(vParam.fechaini) & ")"
                Cad = Cad & " OR (anopsald = " & txtAno(IndiceAnyo + 1).Text & " AND mespsald <= " & Month(vParam.fechafin) & "))"
            End If
        Else
            'Ha pedido de siguiente. Las cuentas las contemplo desde INICIO de ejercicio
            Cad = Cad & " AND ((anopsald = " & Year(vParam.fechaini) & " AND mespsald >= " & Month(vParam.fechaini) & ")"
            Cad = Cad & " OR (anopsald = " & txtAno(IndiceAnyo + 1).Text & " AND mespsald <= " & Me.cmbFecha(IndiceAnyo + 1).ListIndex + 1 & ")"
            'Diferencia de DOS años
            If Val(txtAno(IndiceAnyo + 1).Text) - Year(vParam.fechaini) > 1 Then Cad = Cad & " OR (anopsald = " & Year(vParam.fechaini) + 1 & ")"
            Cad = Cad & ")"
            
        End If
    End If
    
    
    
    
    'Esto es lo que estaba en 21 / NOv / 08
'    If Val(txtAno(0).Text) > Year(vParam.fechaini) Then
'        'Si es en siguiente. Han pedido el balance en ejercicio siguiente
'        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
'            Cad = Cad & " AND (anopsald >= " & Year(vParam.fechaini) & " AND anopsald <= " & Year(vParam.fechafin) + 1 & ")"
'        Else
'            'Años partidos
'            'Desde inicio de ejercicio hasta fin de siguiente
'            Cad = Cad & " AND ((anopsald = " & Year(vParam.fechaini) & " AND mespsald >= " & Me.cmbFecha(0).ListIndex + 1 & ")"
'            Cad = Cad & " OR (anopsald > " & Year(vParam.fechaini) & "))"
'        End If
'
'    Else
'        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
'            'años naturales
'            Cad = Cad & " AND (anopsald >= " & txtAno(0).Text & " AND anopsald <= " & txtAno(1).Text & ")"
'        Else
'            'años partidos
'            Cad = Cad & " AND ((anopsald = " & Year(vParam.fechaini) & " AND mespsald >= " & Me.cmbFecha(0).ListIndex + 1 & ")"
'            Cad = Cad & " OR (anopsald = " & Year(vParam.fechaini) + 1 & " AND mespsald <= " & Me.cmbFecha(1).ListIndex + 1 & "))"
'        End If
'    End If
'
    
    
    Busqueda = Cad

End Sub


Private Sub EmpezarBalance(vConta As Integer, ByRef PB As ProgressBar)
Dim Cade As String
Dim Apertura As Boolean
Dim MesInicioContieneFechaInicioEjercicio As Boolean
Dim QuitarSaldos2 As Byte
Dim Agrupa As Boolean
Dim IndiceCombo As Integer
Dim vOpcion As Byte
Dim Resetea6y7 As Boolean
Dim C1 As Long
Dim UltimoNivel As Byte

Dim PreCargarCierre As Boolean

    Screen.MousePointer = vbHourglass
    
    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    
    
    If vConta = -1 Then
        IndiceCombo = 0
        I = 6
    Else
        IndiceCombo = 14
        I = 18
    End If
    Cade = ""
    
     
'
'    'Es para comprobar tb en los desde hasta fechas
'    Stop
'    If txtAno(0).Text = txtAno(1).Text Then
'        Cade = " (anopsald = " & txtAno(0).Text & " AND mespsald >= " & txtAno(0).Text & " AND mespsald <=" & txtAno(1).Text & ") "
'    Else
'        Cade = " (anopsald = " & txtAno(0).Text & " AND mespsald >= " & txtAno(0).Text & " AND mespsald <=" & txtAno(1).Text & ")"
'        Cade = " (anopsald = " & txtAno(0).Text & " AND mespsald >= " & txtAno(0).Text & " AND mespsald <=" & txtAno(1).Text & ")"
'    End If
'
'
    
    
    If txtCta(I).Text <> "" Then Cade = Cade & " AND ((cuentas.codmacta)>='" & txtCta(I).Text & "')"
    If txtCta(I + 1).Text <> "" Then Cade = Cade & " AND ((cuentas.codmacta)<='" & txtCta(I + 1).Text & "')"
    
    
    
    
    'Genramos el sql
    GeneraSQL Cade, vConta
        
  
    Cad = Cade
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Anterior a 28 ENERO 2004
    Cad = Cad & " GROUP BY codmacta " ' , linapu.fechaent, linapu.fechaent, linapu.codmacta, linapu.codmacta
    Cad = Cad & " ORDER By codmacta"

    Apertura = (Me.chkApertura.Value = 1) And vConta < 0
    
    
    'Para luego veremos que opciones ponemos
    If Me.chkApertura = 1 Then
        vOpcion = 3                   'Con apertura 3 y 4
    Else
        vOpcion = 1                   'Sin apertura 1,2
    End If
    If chkMovimientos = 1 Then
        vOpcion = vOpcion + 1        'Con movimientos
    Else
        vOpcion = vOpcion + 0        ' Sin movimientos
    End If
            
    '1.- Sin apertura y sin movimientos
    '2.- Sin apertura y con movimientos
    '3.- Con apertura y sin movimientos
    '4.- Con apertura y con movimientos
    
    'Vemos si la fecha de inicio de jercicio esta incluida en la fecha inicio seleccionada
    MesInicioContieneFechaInicioEjercicio = False
    If Month(vParam.fechaini) = Me.cmbFecha(IndiceCombo).ListIndex + 1 Then
        'Antes del 20 Febrero estaba este IF tb con el _
        'If Year(vParam.fechaini) = Me.txtAno(IndiceCombo).Text Then
            MesInicioContieneFechaInicioEjercicio = True
    End If
    
    Set RS = New ADODB.Recordset
    
    'Comprobamos si hay que quitar el pyg y el cierre
    QuitarSaldos2 = 0   'No hay k kitar
    Cont = 0
    If Me.chkQuitaCierre(0).Value And Me.chkQuitaCierre(1).Value Then
        Cont = 1  'Ambos
    Else
        If Me.chkQuitaCierre(0).Value Then
            Cont = 2
        Else
            If Me.chkQuitaCierre(1).Value Then Cont = 3
        End If
    End If
    If Me.chkQuitaCierre(0).Value Or Me.chkQuitaCierre(1).Value Then
        'Si el mes contiene el cierre, entonces adelante
        If Month(vParam.fechafin) = Me.cmbFecha(IndiceCombo + 1).ListIndex + 1 Then
            'Si estamos en ejerccicios cerrados seguro que hay asiento de cierre y p y g
            If EjerciciosCerrados Then
                QuitarSaldos2 = Cont
            Else
                'Si no lo comprobamos. Concepto=960 y 980
                Agrupa = HayAsientoCierre((Me.cmbFecha(IndiceCombo + 1).ListIndex + 1), CInt(txtAno(IndiceCombo + 1).Text))
                If Agrupa Then QuitarSaldos2 = Cont
            End If
        End If
    End If
    
    'Agruparemos si esta seleccionado el chekc de agrupar y esta seleccionado
    'ultimo nivel y hay moivmientos para agrupar
    Agrupa = False
    If vConta < 0 And Me.chkAgrupacionCtasBalance.Value = 1 Then 'chekc de agrupar
        If Check2(10).Value = 1 Then                 'sheck de ultimo nivel
            RS.Open "Select * from ctaagrupadas", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then Agrupa = True
            RS.Close
        End If
    End If
    
    'Para los balances de ejercicios siguientes existe la opcion
    ' de que si la cuenta esta en el grupo gto o grupo venta, resetear el importe a 0
    Resetea6y7 = False
    If Me.chkResetea6y7.visible Then
        If Me.chkResetea6y7.Value = 1 Then Resetea6y7 = True
    End If
    
    
    
    
    
    RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If RS.EOF Then
        'NO hay registros a mostrar
        If vConta < 0 Then
            MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
            Screen.MousePointer = vbDefault
        End If
        Agrupa = False
        Cont = -1
    Else
    
        'Voy a ver si precargamos el RScon los datos para el cierr/pyg apertura
        'Veamos si precargamos los
        SQL = ""
        If Check2(10).Value Then
            'Esta chequeado ultimo nivel
            'Veamos si tiene seleccionado alguno mas
            SQL = "1"
            For Cont = 1 To 9
                If Check2(CInt(Cont)).Value = 1 Then SQL = SQL & "1"
            Next Cont
        End If
        PreCargarCierre = Len(SQL) = 1
    
        'Mostramos el frame de resultados
        Cont = 0
        While Not RS.EOF
            Cont = Cont + 1
            RS.MoveNext
        Wend
        PB.visible = True
        PB.Value = 0
        Me.Refresh
        
        
        
            
        
        
        
        
        
        
        
        
        
        
        
        'Borramos los temporales
        SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        
        
        'Nuevo  13 Enero 2005
        ' Pondremos el frame a disabled, veremos el boton de cancelar
        ' y dejaremos k lo pulse
        ' Si lo pulsa cancelaremos y no saldremos
        PulsadoCancelar = False
        Me.frameBalance.Enabled = False
        'Me.cmdCancelarAccion.Visible = True
        Me.cmdCancelarAccion.visible = Legalizacion = ""
        HanPulsadoSalir = False
        Me.Refresh
        
        
        If PreCargarCierre Then PrecargaPerdidasyGanancias EjerciciosCerrados, FechaIncioEjercicio, FechaFinEjercicio, QuitarSaldos2
        
        
        'Dim t1 As Single
        C1 = 0
        RS.MoveFirst
        't1 = Timer
        While Not RS.EOF
            
            CargaBalanceNuevo RS.Fields(0), RS.Fields(1), Apertura, cmbFecha(IndiceCombo).ListIndex + 1, cmbFecha(IndiceCombo + 1).ListIndex + 1, CInt(txtAno(IndiceCombo).Text), CInt(txtAno(IndiceCombo + 1).Text), MesInicioContieneFechaInicioEjercicio, FechaIncioEjercicio, FechaFinEjercicio, EjerciciosCerrados, QuitarSaldos2, vConta, False, Resetea6y7, PreCargarCierre
            
            
            PB.Value = Round((C1 / Cont), 3) * 1000
            PB.Refresh
            DoEvents
            If PulsadoCancelar Then RS.MoveLast
            'Siguiente cta
            C1 = C1 + 1
            RS.MoveNext
        Wend
        'Debug.Print "PreCargarCierre: " & Abs(PreCargarCierre) & "  Reg: " & C1 & "   Tiempo: " & Round(Timer - t1, 4)
        
        
        
        If PreCargarCierre Then CerrarPrecargaPerdidasyGanancias
        
        'Reestablecemos
        PonerFoco cmdCanListExtr(5)
        Me.frameBalance.Enabled = True
        Me.cmdCancelarAccion.visible = False
        HanPulsadoSalir = True
        If PulsadoCancelar Then
            RS.Close
            Screen.MousePointer = vbDefault
            PB.visible = False
            Exit Sub
        End If
        
    End If
    RS.Close
    
    'Ninguna entrada
    If Cont <= 0 Then Exit Sub
    
    'Realizar agrupacion
    If Agrupa Then
        PB.Value = 0
        RS.Open "Select count(*) from ctaagrupadas", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cont = DBLet(RS.Fields(0), "N")
        RS.Close
        If Cont > 0 Then
            SQL = "Select ctaagrupadas.codmacta,nommacta from ctaagrupadas,cuentas where ctaagrupadas.codmacta =cuentas.codmacta "
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 0
            While Not RS.EOF
                If AgrupacionCtasBalance(RS.Fields(0), RS.Fields(1)) Then
                    I = I + 1
                    PB.Value = Round((I / Cont), 3) * 1000
                    RS.MoveNext
                Else
                    RS.Close
                    Exit Sub
                End If
            Wend
        End If
    End If
    
    
    
    
    'Quitamos progress
    PB.Value = 0
    PB.visible = False
    Me.Refresh
    
    
    '--------------------
    'Balance consolidado
    If vConta >= 0 Then
        
        SQL = "Select nomempre from Usuarios.empresas where codempre =" & vConta
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Cad = ""
        If Not RS.EOF Then Cad = DBLet(RS.Fields(0))
        RS.Close
        If Cad = "" Then
            MsgBox "Error leyendo datos empresa: Codempre=" & vConta
            Exit Sub
        End If
        
        SQL = "Select count(*) from Usuarios.ztmpbalancesumas where codusu = " & vUsu.Codigo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Cont = DBLet(RS.Fields(0), "N")
        RS.Close
    
        
        SQL = "Select * from Usuarios.ztmpbalancesumas where codusu = " & vUsu.Codigo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        I = 0
        PB.Value = 0
        Me.Refresh
        SQL = "INSERT INTO Usuarios.ztmpbalanceconsolidado (codempre, nomempre, codusu, cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES ("
        SQL = SQL & vConta & ",'" & Cad & "',"
        While Not RS.EOF
            PB.Value = Round((I / Cont), 3) * 1000
            BACKUP_Tabla RS, Cade
            Cade = Mid(Cade, 2)
            Cade = SQL & Cade
            Conn.Execute Cade
            'Sig
            RS.MoveNext
            I = I + 1
        Wend
        RS.Close
        'Ponemos CONT=0 para k no entre en el if de abajo
        Cont = 0
    End If
    
    Set RS = Nothing
    
    'Si hay datos los mostramos
    If Cont > 0 Then
        'Las fechas
        SQL = "Fechas= ""Desde " & cmbFecha(0).ListIndex + 1 & "/" & txtAno(0).Text & "   hasta "
        SQL = SQL & cmbFecha(1).ListIndex + 1 & "/" & txtAno(1).Text
        
        'Mayo 2015
        'Que en los desde hasta ponga si es antes cierre o antes py
        Cad = ""
        If Me.chkQuitaCierre(0).Value = 1 Then Cad = " pérdidas y ganancias "
        If Me.chkQuitaCierre(1).Value = 1 Then
            If Cad <> "" Then Cad = Cad & " - "
            Cad = Cad & " Cierre "
        End If
        If Cad <> "" Then Cad = "    Antes " & Cad
        SQL = SQL & Cad & """|"
        
        'Si tiene desde hasta codcuenta
        Cad = ""
        If txtCta(6).Text <> "" Then Cad = Cad & "Desde " & txtCta(6).Text & " - " & DtxtCta(6).Tag
        If txtCta(7).Text <> "" Then
            If Cad <> "" Then
                Cad = Cad & "    h"
            Else
                Cad = "H"
            End If
            Cad = Cad & "asta " & txtCta(7).Text & " - " & DtxtCta(7).Tag
        End If
        If Cad = "" Then Cad = " "
        SQL = SQL & "Cuenta= """ & Cad & """|"
        
        'Fecha de impresion
        SQL = SQL & "FechaImp= """ & Text3(7).Text & """|"
        
        
        'Salto
        If Combo3.ListIndex >= 0 Then
            SQL = SQL & "Salto= " & Combo3.ItemData(Combo3.ListIndex) & "|"
            Else
            SQL = SQL & "Salto= 11|"
        End If
        
        'Titulo
        Text1(0).Text = Trim(Text1(0).Text)
        If Text1(0).Text = "" Then
            Cad = "Balance de sumas y saldos"
        Else
            Cad = Text1(0).Text
        End If
        SQL = SQL & "Titulo= """ & Cad & """|"
        
        'Numero de página
        If txtNpag.Text = "" Then
            I = 1
        Else
            I = Val(txtNpag.Text)
        End If
        If I > 0 Then I = I - 1
        
        Cad = "NumPag= " & I & "|"
        SQL = SQL & Cad
        
        
        '------------------------------
        'Numero de niveles
        'Para cada nivel marcado veremos si tiene cuentas en la tmp
        Cont = 0
        UltimoNivel = 0
        For I = 1 To 10
            If Check2(I).visible Then
'                If Check2(I).Value = 1 Then Cont = Cont + 1
                If Check2(I).Value = 1 Then
                    If I = 10 Then
                        Cad = vEmpresa.DigitosUltimoNivel
                    Else
                        Cad = CStr(DigitosNivel(I))
                    End If
                    If TieneCuentasEnTmpBalance(Cad) Then
                        Cont = Cont + 1
                        UltimoNivel = CByte(Cad)
                    End If
                End If
            End If
        Next I
        Cad = "numeroniveles= " & Cont & "|"
        SQL = SQL & Cad
        'Otro parametro mas
        Cad = "vUltimoNivel= " & UltimoNivel & "|"
        SQL = SQL & Cad
        If Opcion = 5 Then
            With frmImprimir
                .OtrosParametros = SQL
                .NumeroParametros = 8
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                'Opcion dependera del combo
                .Opcion = 14 + vOpcion
                .Show vbModal
            End With
        Else
            
            GeneraLegalizaPRF SQL, 7
            CadenaDesdeOtroForm = "OK"
        End If
    End If
    
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function TieneCuentasEnTmpBalance(DigitosNivel As String) As Boolean
Dim RS As ADODB.Recordset
Dim C As String

    Set RS = New ADODB.Recordset
    TieneCuentasEnTmpBalance = False
    C = Mid("__________", 1, CInt(DigitosNivel))
    C = "Select count(*) from Usuarios.ztmpbalancesumas  where cta like '" & C & "'"
    C = C & " AND codusu = " & vUsu.Codigo
    RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then TieneCuentasEnTmpBalance = True
        End If
    End If
    RS.Close
End Function




Private Function ListadoExplotacion(Digitos As Integer, Optional Contabilidad As Integer) As Boolean
Dim Anyo As Integer
Dim Mes As Integer
Dim QuitarSaldos As Boolean
Dim Tot As Long
Dim Cad As String
Dim C1 As String
Dim F As Date
Dim SQL2 As String

    'Borramos las tmp
    ListadoExplotacion = False
    Conn.Execute "Delete from Usuarios.ztmpctaexplotacion where codusu =" & vUsu.Codigo
    'Conn.Execute "Delete from tmpctaexplotacioncierre where codusu =" & vUsu.Codigo
    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    
    If Contabilidad > 0 Then
        Cad = Trim("conta" & Contabilidad) & "."
    Else
        Cad = ""
    End If
    
    'Segun el nivel y segun las ctas de analitica
    SQL2 = "Select " & vUsu.Codigo & ",hsaldos" & Tablas & ".codmacta,nommacta from "
    SQL2 = SQL2 & Cad & "hsaldos" & Tablas & "," & Cad & "cuentas where hsaldos" & Tablas & ".codmacta=cuentas.codmacta AND "
    
    
    '---------------------------------------------------
    'Ubicamos las fechas fin e inicio de ejercicio
    If Contabilidad = 0 Then
        Mes = 2
        Anyo = 4
    Else
        Mes = 19
        Anyo = 18
    End If

    
    Set RS = New ADODB.Recordset
    
    'El probelama es que luego dentro de otras fucniones se utiliza
    'vparam fechaini y fechafin en lugar de enviarles las fechas desde este modulo
    If Contabilidad = 0 Then
        'Solo la empres actual
        F = vParam.fechaini
    Else
       F = ObtenerFechasEjercicioContabilidad(True, Contabilidad)
    End If
    
    
        I = cmbFecha(Mes).ListIndex + 1
        If I >= Month(F) Then
            Cont = Val(txtAno(Anyo).Text)
        Else
            Cont = Val(txtAno(Anyo).Text) - 1
        End If
        Cad = Day(F) & "/" & Month(F) & "/" & Cont
        FechaIncioEjercicio = CDate(Cad)
        
        
    If Contabilidad = 0 Then
        'Solo la empres actual
        F = vParam.fechafin
    Else
       ' If ObtenerFechasEjercicioContabilidad Then
        F = ObtenerFechasEjercicioContabilidad(False, Contabilidad)
    End If
        
        I = cmbFecha(Mes).ListIndex + 1
        If I <= Month(F) Then
            Cont = Val(txtAno(Anyo).Text)
        Else
            Cont = Val(txtAno(Anyo).Text) + 1
        End If
        Cad = Day(F) & "/" & Month(F) & "/" & Cont
        FechaFinEjercicio = CDate(Cad)
    
    
    If Contabilidad = 0 Then
        Anyo = txtAno(4).Text
    Else
        Anyo = txtAno(18).Text
    End If
    
    
    'Si el mes contiene el cierre, entonces adelante cargamos la tabla con los importes a descontar
    'Si estamos en ejerccicios cerrados seguro que hay asiento de cierre y p y g
    QuitarSaldos = False
    If EjerciciosCerrados Then
        If Month(FechaFinEjercicio) = (Me.cmbFecha(Mes).ListIndex + 1) Then QuitarSaldos = True
    Else
        'Si no lo comprobamos. Concepto=960 y 980
        If Month(FechaFinEjercicio) = (Me.cmbFecha(Mes).ListIndex + 1) Then
            QuitarSaldos = HayAsientoCierre((Me.cmbFecha(Mes).ListIndex + 1), Anyo)
        End If
    End If

    If Contabilidad = 0 Then
        Mes = cmbFecha(2).ListIndex + 1
    Else
        Mes = cmbFecha(19).ListIndex + 1
    End If
    
    'Borramos tmp
    Cad = "DELETE FROM tmpcierre1 WHERE codusu = " & vUsu.Codigo
    Conn.Execute Cad
    Cad = " GROUP BY hsaldos" & Tablas & ".codmacta"
    Cad = Cad & " ORDER BY hsaldos" & Tablas & ".codmacta"
    
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'AÑO NATURAL
        SQL = " (hsaldos" & Tablas & ".codmacta like '"
        RC = Mid(vParam.grupogto & "_________", 1, Digitos)
        SQL = SQL & RC
        SQL = SQL & "' OR hsaldos" & Tablas & ".codmacta like '"
        RC = Mid(vParam.grupovta & "_________", 1, Digitos)
        SQL = SQL & RC & "')  And anopsald = " & Anyo
        SQL = SQL2 & SQL
        C1 = "INSERT INTO tmpcierre1 (codusu, cta, nomcta) " & SQL & Cad
        Conn.Execute C1
        
        
        'Si tiene otro grupo.  Ejemplo AUTOMOCION (alcodar, perez pascual...)
        If vParam.grupoord <> "" Then
            SQL = " (hsaldos" & Tablas & ".codmacta like '"
            RC = Mid(vParam.grupoord & "_________", 1, Digitos) & "'"
            SQL = SQL & RC
            If vParam.Automocion <> "" Then
                'Grupo excepcion en el grupo "otro"
                SQL = SQL & " AND not (hsaldos" & Tablas & ".codmacta like '" & vParam.Automocion & "%')"

            End If
            SQL = SQL & " AND anopsald = " & Anyo & ")"
            SQL = SQL2 & SQL
            C1 = "INSERT INTO tmpcierre1 (codusu, cta, nomcta) " & SQL & Cad
            Conn.Execute C1
        End If
    Else
        'AÑO pARTIDO
        'GASTo
        RC = "'" & Mid(vParam.grupogto & "_________", 1, Digitos) & "'"
        C1 = " hsaldos" & Tablas & ".codmacta like " & RC
        'EJEMPLO
        'AND (( anopsald = 2003 AND mespsald >=9) OR ( anopsald = 2004 AND mespsald <=8))
        C1 = C1 & " AND (( anopsald = " & Year(FechaIncioEjercicio) & " AND mespsald >=" & Month(FechaIncioEjercicio) & ")"
        C1 = C1 & " OR ( anopsald = " & Year(FechaFinEjercicio) & " AND mespsald <=" & Month(FechaFinEjercicio) & "))"
    
        C1 = "INSERT INTO tmpcierre1 (codusu, cta, nomcta)  " & SQL2 & C1 & Cad
        Conn.Execute C1
        'VEntas
        RC = "'" & Mid(vParam.grupovta & "_________", 1, Digitos) & "'"
        C1 = " hsaldos" & Tablas & ".codmacta like " & RC
        'EJEMPLO
        'AND (( anopsald = 2003 AND mespsald >=9) OR ( anopsald = 2004 AND mespsald <=8))
        C1 = C1 & " AND (( anopsald = " & Year(FechaIncioEjercicio) & " AND mespsald >=" & Month(FechaIncioEjercicio) & ")"
        C1 = C1 & " OR ( anopsald = " & Year(FechaFinEjercicio) & " AND mespsald <=" & Month(FechaFinEjercicio) & "))"
        C1 = "INSERT INTO tmpcierre1 (codusu, cta, nomcta)  " & SQL2 & C1 & Cad
        Conn.Execute C1
    End If
    
    
    
    'Continuamos
    SQL = "Select * from tmpcierre1 where codusu = " & vUsu.Codigo & " ORDER BY cta"
    
    RS.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then
        RS.Close
        If Contabilidad = 0 Then MsgBox "Ningun dato con estos valores ", vbExclamation
        Exit Function
    End If
    
    Dim S
    
    SQL = "INSERT INTO Usuarios.ztmpctaexplotacion (codusu, cta, Contador, nomcta,"
    SQL = SQL & "acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES (" & vUsu.Codigo & ","
    
    'Veo cuantos regitros hay para el progress
    RS.MoveFirst
    Cont = 0
    While Not RS.EOF
        RS.MoveNext
        Cont = Cont + 1
    Wend
    RS.MoveFirst
    Tot = Cont + 3
    If Contabilidad = 0 Then
        pb3.Value = 0
        pb3.visible = True
    Else
        pb10.Value = 0
        pb10.visible = True
    End If
    Cont = 1
    
    'Insertamos , si ha lugar, las existencias inciales
    If Contabilidad = 0 Then
        If Me.txtExplo(0).Text <> "" Or txtExplo(2).Text <> "" Then
            Importe = 0
            'Hay que insertar existencias iniciales
            If txtExplo(0).Text = "" Then
                RC = "NULL"
            Else
                RC = TransformaComasPuntos(txtExplo(0).Text)
                Importe = CCur(txtExplo(0).Text)
            End If
            '---    cta, conta    , titulo
            Cad = "''," & Cont & ",'EXISTENCIAS INICIALES'," & RC & ",NULL,"
            If txtExplo(2).Text = "" Then
                RC = "NULL"
            Else
                RC = TransformaComasPuntos(txtExplo(2).Text)
                Importe = Importe + CCur(txtExplo(2).Text)
            End If
            Cad = Cad & RC & ",NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL)"
            Conn.Execute SQL & Cad
            Cont = Cont + 1
        End If
    End If
    'Movimientos
    While Not RS.EOF
        CuentaExplotacion RS.Fields(1), RS.Fields(2), Mes, Anyo, Cont, 1, SQL, EjerciciosCerrados, FechaIncioEjercicio, FechaFinEjercicio, QuitarSaldos, Contabilidad
        I = CInt((Cont / Tot) * 1000)
        If Contabilidad = 0 Then
            pb3.Value = I
        Else
            pb10.Value = I
        End If
        'Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'Auqi ira el de fin de existencias
    'Insertamos , si ha lugar, las existencias inciales
    If Contabilidad = 0 Then
        If Me.txtExplo(1).Text <> "" Or txtExplo(3).Text <> "" Then
            Importe = 0
            'Hay que insertar existencias iniciales
            If txtExplo(1).Text = "" Then
                RC = "NULL"
            Else
                RC = TransformaComasPuntos(txtExplo(1).Text)
                Importe = CCur(txtExplo(1).Text)
            End If
            '---    cta, conta    , titulo
            Cad = "''," & Cont & ",'EXISTENCIAS FINALES',NULL," & RC & ","
            If txtExplo(3).Text = "" Then
                RC = "NULL"
            Else
                RC = TransformaComasPuntos(txtExplo(3).Text)
                Importe = Importe + CCur(txtExplo(3).Text)
            End If
            Cad = Cad & "NULL," & RC & ",NULL," & TransformaComasPuntos(CStr(Importe)) & ")"
            Conn.Execute SQL & Cad
            pb3.Value = pb3.Max
        End If
    End If
    
    ListadoExplotacion = True
End Function




Private Sub txtPag2_GotFocus(Index As Integer)
    PonFoco txtPag2(Index)
End Sub

Private Sub txtPag2_KeyPress(Index As Integer, KeyAscii As Integer)
     ListadoKEYpress KeyAscii
End Sub

Private Sub txtPag2_LostFocus(Index As Integer)
    
    txtPag2(Index).Text = Trim(txtPag2(Index).Text)
    If txtPag2(Index).Text <> "" Then
        If Index = 0 Then
            If Not IsNumeric(txtPag2(Index).Text) Then
                MsgBox "Numero de página incorrecto: " & txtPag2(Index).Text, vbExclamation
                txtPag2(Index).Text = ""
            End If
         End If
    End If
End Sub

Private Sub txtperiodo_GotFocus(Index As Integer)
 PonFoco txtperiodo(Index)
End Sub

Private Sub txtperiodo_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtperiodo_LostFocus(Index As Integer)
Dim Bien As Boolean
   txtperiodo(Index).Text = Trim(txtperiodo(Index).Text)
   If txtperiodo(Index).Text = "" Then Exit Sub
   
   Bien = True
   If Not IsNumeric(txtperiodo(Index).Text) Then
        MsgBox "El campo debe ser numérico", vbExclamation
        Bien = False
    Else
        If Val(txtperiodo(Index).Text) <= 0 Then
            MsgBox "El campo debe ser mayor que 0.", vbExclamation
            Bien = False
        End If
    End If
    If Not Bien Then txtperiodo(Index).SetFocus
End Sub


Private Sub txtReemisionDiario_GotFocus()
    PonFoco txtReemisionDiario
End Sub

Private Sub txtReemisionDiario_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

Private Sub txtserie_GotFocus(Index As Integer)
    PonFoco txtSerie(Index)
End Sub

Private Sub txtSerie_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtserie_KeyPress(Index As Integer, KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub txtserie_LostFocus(Index As Integer)
    txtSerie(Index).Text = Trim(txtSerie(Index).Text)
    If txtSerie(Index) <> "" Then txtSerie(Index) = UCase(txtSerie(Index).Text)
End Sub


Public Function GeneraBalancePresupuestario() As Boolean
Dim Aux As String
Dim Importe As Currency
Dim AUX2 As String
Dim vMes  As Integer
Dim Cta As String

On Error GoTo EGeneraBalancePresupuestario
    GeneraBalancePresupuestario = False
    If Me.chkPreMensual.Value = 0 Then
        Aux = "select codmacta,sum(imppresu)  from presupuestos "
        If SQL <> "" Then Aux = Aux & " where " & SQL
        Aux = Aux & " group by codmacta"
        
        'Para el otro
        Cad = "Select SUM(impmesde),SUM(impmesha) from hsaldos where anopsald=" & I
        Cad = Cad & " and codmacta = '"
    Else
        Aux = "select codmacta,imppresu,mespresu from presupuestos where " & SQL
        If txtMes(2).Text <> "" Then Aux = Aux & " and mespresu = " & txtMes(2).Text
        Aux = Aux & " ORDER BY codmacta,mespresu"
        'para luego
        Cad = "Select impmesde,impmesha from hsaldos where anopsald=" & I
        If txtMes(2).Text <> "" Then Cad = Cad & " and mespsald = " & txtMes(2).Text
        Cad = Cad & " and codmacta = '"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ningún registro a mostrar.", vbExclamation
        RS.Close
        Exit Function
    End If
    
    'Borramos tmp de presu 2
    Aux = "DELETE FROM Usuarios.ztmppresu2 where codusu =" & vUsu.Codigo
    Conn.Execute Aux
    
    SQL = "INSERT INTO Usuarios.ztmppresu2 (codusu, codigo, cta, titulo,  mes, Presupuesto, realizado) VALUES ("
    SQL = SQL & vUsu.Codigo & ","
    
    Cont = 0
    Do
        Cont = Cont + 1
        RS.MoveNext
    Loop Until RS.EOF
    RS.MoveFirst
    
    'Ponemos el PB4
    pb4.Max = Cont + 1
    pb4.Value = 0
    If Cont > 3 Then pb4.visible = True
    Cta = ""
    Cont = 1   'Contador
    While Not RS.EOF
        If Me.chkPreMensual.Value = 1 Then
            If Cta <> RS!codmacta Then
                vMes = 1
                Cta = RS!codmacta
            End If
            
            If RS!mespresu > vMes Then
                For I = vMes To RS!mespresu - 1
                
                    Aux = RS!codmacta  'Aqui pondremos el nombre
                    Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
                    Aux = Cont & ",'" & RS!codmacta & "','" & DevNombreSQL(Aux) & "',"
                    Aux = Aux & I
             
                    Aux = Aux & ",0,"
                    
                    AUX2 = Cad & RS!codmacta & "'"
                    AUX2 = AUX2 & " AND mespsald =" & I
                    
                
                
                    Importe = ImporteBalancePresupuestario(AUX2)
                    
                    Aux = Aux & TransformaComasPuntos(CStr(Importe)) & ")"
                    If Importe <> 0 Then
                        Conn.Execute SQL & Aux
                        Cont = Cont + 1
                    End If
                Next I
            End If
            
        End If
                
        
    
    
        Aux = RS!codmacta  'Aqui pondremos el nombre
        Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
        Aux = Cont & ",'" & RS!codmacta & "','" & DevNombreSQL(Aux) & "',"
        If Me.chkPreMensual.Value = 0 Then
            Aux = Aux & "0"
        Else
            Aux = Aux & RS!mespresu
        End If
        Aux = Aux & "," & TransformaComasPuntos(CStr(RS.Fields(1))) & ","
        
        'SQL
        AUX2 = Cad & RS!codmacta & "'"
        If Me.chkPreMensual.Value = 1 Then
            AUX2 = AUX2 & " AND mespsald =" & RS!mespresu
            'AUmento el mes
            vMes = RS!mespresu + 1
        End If
        
        
        Importe = ImporteBalancePresupuestario(AUX2)
        'Debug.Print Importe
        Aux = Aux & TransformaComasPuntos(CStr(Importe)) & ")"
        Conn.Execute SQL & Aux
        
        'Sig
        pb4.Value = pb4.Value + 1
        Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.Close
    
    
        '2013  Junio
    ' QUitaremos si asi lo pide, el saldo de la apertura
    ' Curiosamente, las 6 y 7  NO tienen apertura(perdi y ganacias)
    RC = "" 'Por si quitamos el apunte de apertura. Guardare las cuentas para buscarlas despues en la apertura
    If chkQuitarApertura.Value = 1 Then
        Aux = "SELECT cta from Usuarios.ztmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
        RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            RC = RC & ", '" & RS!Cta & "'"
            RS.MoveNext
        Wend
        RS.Close
        
        
        
        'Subo qui lo de quitar apertura
        If RC <> "" Then
            RC = Mid(RC, 2)
            Aux = " AND codmacta IN (" & RC & ")"
            
            Cad = "SELECT codmacta cta,sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) as importe"
            Cad = Cad & " from hlinapu where codconce=970 and fechaent='" & Format(vParam.fechaini, FormatoFecha) & "'"
            Cad = Cad & Aux
            Cad = Cad & " GROUP BY 1"
            RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                Cad = "UPDATE Usuarios.ztmppresu2 SET realizado=realizado-" & TransformaComasPuntos(CStr(RS!Importe))
                
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & RS!Cta & "' AND mes = "
                If Me.chkPreMensual.Value = 1 Then
                    Cad = Cad & " 1"
                Else
                    Cad = Cad & " 0"
                End If
                Conn.Execute Cad
                RS.MoveNext
            Wend
            RS.Close
                
            
            
        End If
        
        
        
    End If
    
    
    'Si pide a 3 DIGITOS este es el momemto
    'Sera facil.
    'Hacemos un insert into con substring
 
        'SUBNIVEL
        Aux = ""
        For I = 1 To 9
            If ChkCtaPre(I).Value = 1 Then
                
                Aux = DevuelveDesdeBD("count(*)", "Usuarios.ztmppresu2", "codusu", CStr(vUsu.Codigo))
                Cont = Val(Aux)
                
                '@rownum:=@rownum+1 AS rownum      (SELECT @rownum:=0) r
                Aux = "Select " & vUsu.Codigo & " us,@rownum:=@rownum+1 AS rownum,substring(cta,1," & I & ") as cta2,mes,sum(presupuesto),sum(realizado)"
                Aux = Aux & " FROM Usuarios.ztmppresu2,(SELECT @rownum:=" & Cont & ") r WHERE codusu = " & vUsu.Codigo
                
                Aux = Aux & " AND length(cta)=" & vEmpresa.DigitosUltimoNivel
                
                Aux = Aux & " group by cta2,us,mes"
                Aux = "insert into Usuarios.ztmppresu2 (codusu, codigo, cta,   mes, Presupuesto, realizado) " & Aux
                'Insertamos
                Conn.Execute Aux
                
                'Quito los de ultimo nivel

                
                Aux = "SELECT cta from Usuarios.ztmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
                RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    'Actualizo el nommacta
                    Aux = RS!Cta  'Aqui pondremos el nombre
                    Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
                    Aux = "UPDATE Usuarios.ztmppresu2  SET titulo = '" & DevNombreSQL(Aux) & "' WHERE codusu = " & vUsu.Codigo & " AND Cta = '" & RS!Cta & "'"
                    Conn.Execute Aux
                    RS.MoveNext
                Wend
                RS.Close
                
                
                
            End If
        Next
        
        
        If ChkCtaPre(10).Value = 0 Then
            Aux = "DELETE FROM Usuarios.ztmppresu2 WHERE codusu = " & vUsu.Codigo & " AND cta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
            Conn.Execute Aux
        End If
        
    
    
    
  
            
        
  
    
    
    Set RS = Nothing
    GeneraBalancePresupuestario = True
    Exit Function
EGeneraBalancePresupuestario:
    MuestraError Err.Number, "Gen. balance presupuestario"
    Set RS = Nothing
End Function


Private Function GneraListadoPresupuesto() As Boolean

    On Error GoTo EGneraListadoPresupuesto
    GneraListadoPresupuesto = False
    If SQL <> "" Then SQL = " AND " & SQL
    SQL = "select presupuestos.* ,nommacta from presupuestos,cuentas where presupuestos.codmacta=cuentas.codmacta " & SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        RS.Close
        MsgBox "Ningun registro a listar.", vbExclamation
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztmppresu1 where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "INSERT INTO Usuarios.ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (" & vUsu.Codigo & ","
    I = I
    While Not RS.EOF
        Cad = I & ",'" & RS!codmacta & "','" & RS!nommacta & "'," & RS!anopresu
        Cad = Cad & "," & RS!mespresu & "," & TransformaComasPuntos(CStr(RS!imppresu)) & ")"
        Conn.Execute SQL & Cad
        'Sig
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close
    GneraListadoPresupuesto = True
EGneraListadoPresupuesto:
    If Err.Number <> 0 Then MuestraError Err.Number, "Listado Presupuesto"
    Set RS = Nothing
    
End Function



'Certificado de IVA
Private Sub ProcesaRegstrosParaBD(CONTA As String, TipoIva As Integer)
Dim Aux As String

    pb5.Value = 2
    SQL = "Select *,nommacta,nifdatos,pais FROM " & CONTA & ".cabfact as cabf," & CONTA & ".cuentas as cta"
    SQL = SQL & " WHERE cabf.codmacta = cta.codmacta AND "
    'Los tipos de iva
    SQL = SQL & "( tp1faccl= " & TipoIva & " OR tp2faccl= " & TipoIva & " OR tp2faccl= " & TipoIva & ")"
    
    'Las fechas
    If Text3(13).Text <> "" Then SQL = SQL & " AND fecfaccl >=' " & Format(Text3(13).Text, FormatoFecha) & "'"
    If Text3(14).Text <> "" Then SQL = SQL & " AND fecfaccl <=' " & Format(Text3(14).Text, FormatoFecha) & "'"
    
    SQL = SQL & " order by numserie,codfaccl"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    'Refrescamos
    Me.Refresh
    If Not RS.EOF Then
        'Fiajmos el pb
        I = 0
        Do
            RS.MoveNext
            I = I + 1
        Loop Until RS.EOF
        If I > 32000 Then
            I = 32000
        End If
        RS.MoveFirst
        pb5.Max = I + 4
        pb5.Value = 3
        'Insertamos
        SQL = "INSERT INTO Usuarios.zcertifiva (codusu,  factura, fecha, destino,pais"
        'ANTES
        'SQL = SQL & "codigo,Importe, tipoiva, iva)"
        SQL = SQL & ",nif,codigo,Importe, tipoiva, iva)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & ","
        While Not RS.EOF
            
            Aux = SerieNumeroFactura(9, RS!NUmSerie, RS!codfaccl)
            Aux = "'" & Aux & "','" & Format(RS!fecfaccl, "dd/mm/yyyy") & "','" & DevNombreSQL(CStr(RS!nommacta)) & "','"
            Aux = Aux & DBLet(RS!Pais) & "','" & DBLet(RS!nifdatos) & "',"
                        
            'Tipo de IVA 1
            If RS!tp1faccl = TipoIva Then
                RC = TransformaComasPuntos(CStr(RS!ba1faccl))
                RC = RC & "," & RS!tp1faccl & ","
                RC = RC & TransformaComasPuntos(CStr(RS!pi1faccl)) & ")"
                RC = Cont & "," & RC
                Conn.Execute SQL & Aux & RC
                Cont = Cont + 1
            End If
            
            
            'Tipo de iva 2
            If Not IsNull(RS!tp2faccl) Then
                If RS!tp2faccl = TipoIva Then
                    RC = TransformaComasPuntos(CStr(RS!ba2faccl))
                    RC = RC & "," & RS!tp2faccl & ","
                    RC = RC & TransformaComasPuntos(CStr(RS!pi2faccl)) & ")"
                    RC = Cont & "," & RC
                    Conn.Execute SQL & Aux & RC
                    Cont = Cont + 1
                End If
            End If
            
            
            'Tipo de iva 3
            If Not IsNull(RS!tp3faccl) Then
                If RS!tp3faccl = TipoIva Then
                    RC = TransformaComasPuntos(CStr(RS!ba3faccl))
                    RC = RC & "," & RS!tp3faccl & ","
                    RC = RC & TransformaComasPuntos(CStr(RS!pi3faccl)) & ")"
                    RC = Cont & "," & RC
                    Conn.Execute SQL & Aux & RC
                    Cont = Cont + 1
                End If
            End If
            
            
            'Siguiente
            RS.MoveNext
            If pb5.Value < pb5.Max Then pb5.Value = pb5.Value + 1
        Wend
    End If
    RS.Close
    
End Sub



Private Sub cargacomboiva(ByRef QueComboIva As ComboBox)
    List1.Clear
    'Combo4.Clear
    QueComboIva.Clear
    If QueComboIva.Name = "Combo8" Then
        QueComboIva.AddItem "  "
        QueComboIva.ItemData(QueComboIva.NewIndex) = -1
    End If
    
    SQL = "Select * from tiposiva order by codigiva"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        'Combo4.AddItem RS!nombriva & "    (Cod: " & RS!codigiva & " )"
        'Combo4.ItemData(Combo4.NewIndex) = RS!codigiva
        QueComboIva.AddItem RS!nombriva & "    (Cod: " & RS!codigiva & " )"
        QueComboIva.ItemData(QueComboIva.NewIndex) = RS!codigiva
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub PonValoresLiquidacion()
On Error GoTo EPonValoresLiquidacion

    txtperiodo(2).Text = vParam.anofactu
    I = vParam.perfactu + 1
    If vParam.periodos = 0 Then
        Cont = 4
    Else
        Cont = 12
    End If
        
    If I > Cont Then
            I = 1
            txtperiodo(2).Text = vParam.anofactu + 1
    End If
    txtperiodo(0).Text = I
    txtperiodo(1).Text = I
    Exit Sub
EPonValoresLiquidacion:
    MuestraError Err.Number
End Sub




Private Sub PonerEmpresaSeleccionEmpresa(vOpcion As Byte)

    SQL = "Select * from Usuarios.empresas "
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    
    Cont = 0
    I = 0
    While Not RS.EOF
        Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Si solo hay una empresa deshabilitamos el boton, dependiendo de vOpcion
    Select Case vOpcion
    Case 0
        Image4.visible = (Cont > 1)
        List1.AddItem vEmpresa.nomempre
        List1.ItemData(0) = vEmpresa.codempre
    Case 1
        Image5.visible = (Cont > 1)
        List2.AddItem vEmpresa.nomempre
        List2.ItemData(0) = vEmpresa.codempre
    Case 2
        Image6.visible = (Cont > 1)
        List3.AddItem vEmpresa.nomempre
        List3.ItemData(0) = vEmpresa.codempre
    Case 3
        Image7.visible = (Cont > 1)
        List4.AddItem vEmpresa.nomempre
        List4.ItemData(0) = vEmpresa.codempre
    Case 4
        '349
        Image8.visible = (Cont > 1)
        List5.AddItem vEmpresa.nomempre
        List5.ItemData(0) = vEmpresa.codempre
    Case 5
        'Traspaso ACE
        Image9.visible = (Cont > 1)
        List6.AddItem vEmpresa.nomempre
        List6.ItemData(0) = vEmpresa.codempre
    Case 6
        'Cta explotacion consolidada
        Image10.visible = (Cont > 1)
        List7.AddItem vEmpresa.nomempre
        List7.ItemData(0) = vEmpresa.codempre
    Case 7
        'Cta pyg y balance situacion
        Image11.visible = (Cont > 1)
        List8.AddItem vEmpresa.nomempre
        List8.ItemData(0) = vEmpresa.codempre
    
    
    Case 8
        'Facturas proveedores
        Image12.visible = (Cont > 1)
        List9.AddItem vEmpresa.nomempre
        List9.ItemData(0) = vEmpresa.codempre
    
    Case 9
        'Facturas proveedores
        Image13.visible = (Cont > 1)
        List10.AddItem vEmpresa.nomempre
        List10.ItemData(0) = vEmpresa.codempre

    
    End Select
    
        
End Sub



Private Function GeneraLasLiquidaciones() As Boolean
    
    '       cliprov     0- Facturas clientes
    '                   1- RECARGO EQUIVALENCIA
    '                   2- Facturas proveedores
    '                   3- libre
    '                   4- IVAS no deducible
    '                   5- Facturas NO DEDUCIBLES
    '                   6- IVA BIEN INVERSION
    '                   7- Compras extranjero
    '                   8- Inversion sujeto pasivo (Abril 2015)
    
    'Borramos los datos temporales
    SQL = "DELETE FROM Usuarios.zliquidaiva WHERE codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    'Modificacion para desglosaar los IVAS que sean:
    '   Intracom
    '   Regimen especial agrario
    '    inversion sujeto pasivo
    '...
    '  Para ello en tmpcierre1 pondremos para el usuario
    '  en nommacta: adqintra   ,  ventintra, campo
    '  para cada empresa
    SQL = "DELETE FROM tmpctaexplotacioncierre where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    'Si quiere ver el IVA detallado
    If Me.chkIVAdetallado.Value = 1 Then
        SQL = "DELETE FROM Usuarios.ztmpimpbalan WHERE codusu =" & vUsu.Codigo
        Conn.Execute SQL
        SQL = "DELETE FROM tmpimpbalance WHERE codusu =" & vUsu.Codigo
        Conn.Execute SQL
    End If
    
    NumRegElim = 0
    'Para cada empresa
    'Para cada periodo
    For I = 0 To List2.ListCount - 1
        For Cont = CInt(txtperiodo(0).Text) To CInt(txtperiodo(1).Text)
            Label13.Caption = Mid(List2.List(I), 1, 20) & ".  " & Cont
            Label13.Refresh
            LiquidacionIVA CByte(Cont), CInt(txtperiodo(2).Text), List2.ItemData(I), (chkIVAdetallado.Value = 1)
        Next Cont
    Next I
    'Borraremos todos aquellos IVAS de Base imponible=0
    SQL = "DELETE FRom Usuarios.zliquidaiva WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " AND bases = 0"
    Conn.Execute SQL
    
    If Me.chkIVAdetallado.Value = 1 Then
        'Insertamos en Usuarios para el posible informe
        SQL = "INSERT INTO Usuarios.ztmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita) "
        SQL = SQL & " SELECT codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita FROM tmpimpbalance "
        SQL = SQL & " WHERE codusu=" & vUsu.Codigo
        Conn.Execute SQL
        
    End If
    
    GeneraLasLiquidaciones = True
End Function


Private Function PonerDatosEmpresa() As String
Dim F As Date
    Set RS = New ADODB.Recordset
    RS.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If RS.EOF Then
        MsgBox "Error leyendo datos de la empresa", vbExclamation
        SQL = ""
    Else
        'La direccion
        Cad = DBLet(RS!siglasvia, "T")
        If Cad <> "" Then Cad = Cad & "  "
        Cad = Cad & Trim(DBLet(RS!Direccion)) & " "
        SQL = Trim(DBLet(RS!Numero, "T") & " " & DBLet(RS!escalera, "T") & " " & DBLet(RS!piso, "T") & " " & DBLet(RS!puerta, "T"))
        Cad = Cad & SQL
        
        SQL = "Apoderado= """ & DBLet(RS!apoderado) & """|"
        SQL = SQL & "Domicilio= """ & Cad & """|"
        
        'El año lo cojo de la fecha hasta
        F = CDate(Text3(14).Text)
        Cad = Year(F)
        SQL = SQL & "Ejercicio= """ & Cad & """|"
        SQL = SQL & "Empresa= """ & vEmpresa.nomempre & """|"
        
        F = CDate(Text3(12).Text)
        Cad = DBLet(RS!poblacion) & ", " & Format(F, "dd") & " de " & Format(F, "mmmm") & " de " & Format(F, "yyyy")
        SQL = SQL & "Fecha= """ & Cad & """|"
        SQL = SQL & "NIF= """ & DBLet(RS!nifempre) & """|"
        
        'Periodo. Vere si es un mes u tres
        F = CDate(Text3(14).Text)
        I = DateDiff("m", CDate(CDate(Text3(13).Text)), F)
        
        If I > 1 Then
            'TRIMESTRAL
            I = Month(CDate(Text3(13).Text))
            I = (I Mod 4)
        Else
            I = Month(F)
        End If
        Cad = CStr(I)
        SQL = SQL & "Periodo= """ & Cad & """|"
    End If
    RS.Close
    PonerDatosEmpresa = SQL
End Function


Private Sub txtTitulo_GotFocus()
    PonFoco txtTitulo
End Sub

Private Sub txtTitulo_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub

'Siempre k la fecha no este en fecha siguiente
Private Function HayAsientoCierre(Mes As Byte, Anyo As Integer, Optional Contabilidad As String) As Boolean
Dim C As String
    HayAsientoCierre = False
    'C = "01/" & CStr(Me.cmbFecha(1).ListIndex + 1) & "/" & txtAno(1).Text
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierre = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            C = "Select count(*) from " & Contabilidad
            C = C & " hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then
                    If RS.Fields(0) > 0 Then HayAsientoCierre = True
                End If
            End If
            RS.Close
        End If
    End If
End Function


Private Function GeneraDiarioOficial() As Boolean
Dim Total As Long
Dim Pos As Long
Dim miCo As Long

    On Error GoTo EGeneraDiarioOficial

    GeneraDiarioOficial = False
    Tablas = "hlinapu"
    If EjerciciosCerrados Then Tablas = Tablas & "1"
    'Parte comun
    Cad = " from " & Tablas & ",cuentas"
    Cad = Cad & " WHERE " & Tablas & ".codmacta = cuentas.codmacta"
    Cad = Cad & " AND fechaent >='" & Format(Text3(15).Text, FormatoFecha) & "'"
    Cad = Cad & " AND fechaent <='" & Format(Text3(16).Text, FormatoFecha) & "'"
    
    'Para el contador
    SQL = "Select count(*) " & Cad
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Total = 0
    If Not RS.EOF Then
        Total = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    
    If Total = 0 Then
        MsgBox "Ningun asiento entre esas fechas.", vbExclamation
        Exit Function
    End If
    
    Me.cmdCancelarAccion.visible = True
    PulsadoCancelar = False
    
    'Borramos la temporal
    Conn.Execute "Delete from usuarios.ztmplibrodiario where codusu = " & vUsu.Codigo
    
    'Ya tenemos el total
    SQL = "select fechaent,numasien,linliapu,cuentas.codmacta, cuentas.nommacta,numdocum,"
    SQL = SQL & "ampconce,timported,timporteh " & Cad
    SQL = SQL & " ORDER BY fechaent,numasien"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    DoEvents
    If PulsadoCancelar Then
        RS.Close
        Exit Function
    End If
    pb6.Max = 1000
    pb6.Value = 0
    pb6.visible = True
    Pos = 1
    miCo = RS!Numasien   'le damos el primer valor
    Cont = Val(txtLibroOf(0).Text)
    
    'Construimos la mitad de cadena de insercion
    SQL = "INSERT INTO usuarios.ztmplibrodiario (codusu, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, nommacta, numdocum, ampconce, debe, haber) VALUES (" & vUsu.Codigo & ","
    
    While Not RS.EOF
       
        If chkRenumerar.Value = 1 Then
            'Si k estamos renumerando
           If miCo <> RS!Numasien Then
                Cont = Cont + 1
                miCo = RS!Numasien
            End If
        Else
            Cont = RS!Numasien
        End If
        pb6.Value = CInt(((Pos / Total) * 1000))
        
        Cad = RS!nommacta
        NombreSQL Cad
        Cad = "'" & Format(RS!fechaent, FormatoFecha) & "'," & Cont & "," & RS!Linliapu & ",'" & RS!codmacta & "','" & Cad & "','"
        Cad = Cad & DevNombreSQL(DBLet(RS!numdocum)) & "','" & DevNombreSQL(DBLet(RS!ampconce)) & "',"
        If Not IsNull(RS!timported) Then
            Tablas = TransformaComasPuntos(CStr(RS!timported))
            Cad = Cad & Tablas & ",NULL)"
        Else
            Tablas = TransformaComasPuntos(CStr(RS!timporteH))
            Cad = Cad & "NULL," & Tablas & ")"
        End If
        Cad = SQL & Cad
        Conn.Execute Cad
        'Siguiente
        Pos = Pos + 1
        DoEvents
        If PulsadoCancelar Then
            RS.Close
            Exit Function
        End If
        RS.MoveNext
    Wend
    RS.Close
     
    GeneraDiarioOficial = True
    Exit Function
EGeneraDiarioOficial:
    MuestraError Err.Number
End Function



Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    Cad = RecuperaValor(Lista, L)
    If Cad <> "" Then
        I = Val(Cad)
        With cmbFecha(I)
            .Clear
            For Cont = 1 To 12
                RC = "25/" & Cont & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next Cont
        End With
    End If
    L = L + 1
Loop Until Cad = ""
End Sub





Private Sub UltimoMesAnyoAnal1(ByRef Mes As Integer, ByRef Anyo As Integer)
    Anyo = 1900
    Mes = 13
    Set miRsAux = New ADODB.Recordset
    SQL = "select max(anoccost) from hsaldosanal1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Anyo = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    If Anyo > 1900 Then
        SQL = "select max(mesccost) from hsaldosanal1 where anoccost =" & Anyo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Mes = 1
        If Not miRsAux.EOF Then
            Mes = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
    End If
    Set miRsAux = Nothing
End Sub

Private Function GeneraCtaExplotacionCC() As Boolean
Dim RC As Byte

    GeneraCtaExplotacionCC = False
    
    
    'Borramos datos
    SQL = "Delete from Usuarios.zctaexpcc where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    

    If chkCtaExpCC(1).Value = 1 Then
        'Hacemos primero el periodo anterior
        RC = HacerCtaExploxCC(CInt(txtAno(7).Text) - 1, CInt(txtAno(8).Text) - 1)
        If RC = 0 Then Exit Function  'ha habido algun error
        
        'Si ha Generado datos los paso un momento para que el seguiente proceso no los borre
        If RC = 2 Then
            If Not ProcesoCtaExplotacionCC(0) Then Exit Function
        End If
        
    End If
    
    'Este bloque lo hace siempre
    RC = HacerCtaExploxCC(CInt(txtAno(7).Text), CInt(txtAno(8).Text))
    If RC > 0 Then
        RC = 0
        'Si era compartativo
        'Mezclamos los datos de ahora con los de antes
        If chkCtaExpCC(1).Value = 1 Then
            If Not ProcesoCtaExplotacionCC(1) Then RC = 1
        End If
        If RC = 0 Then GeneraCtaExplotacionCC = True
    End If
    

    
    'Eliminamos datos temporales
    If chkCtaExpCC(1).Value = 1 Then
        ProcesoCtaExplotacionCC 2
    End If
    

End Function

'0: Error    1: No hya datods       2: OK
Private Function HacerCtaExploxCC(Anyo1 As Integer, Anyo2 As Integer) As Byte
Dim A1 As Integer, M1 As Integer
Dim Post As Boolean

    On Error GoTo EGeneraCtaExplotacionCC
    HacerCtaExploxCC = 0
    
    UltimoMesAnyoAnal1 M1, A1
    
    'Si años consulta iguales
    If txtAno(7).Text = txtAno(8).Text Then
         Cad = " anoccost=" & Anyo1 & " AND mesccost>=" & Me.cmbFecha(5).ListIndex + 1
         Cad = Cad & " AND mesccost<=" & Me.cmbFecha(6).ListIndex + 1
         
    Else
        'Años disitintos
        'Inicio
        Cad = "( anoccost=" & Anyo1 & " AND mesccost>=" & Me.cmbFecha(5).ListIndex + 1 & ")"
        Cad = Cad & " OR ( anoccost=" & Anyo2 & " AND mesccost<=" & Me.cmbFecha(6).ListIndex + 1 & ")"
        'Por si la diferencia es mas de un año
        If Val(txtAno(8).Text) - Val(txtAno(7).Text) > 1 Then
            Cad = Cad & " OR (anoccost >" & Anyo1 & " AND anoccost < " & Anyo2 & ")"
        End If
    End If
    Cad = " (" & Cad & ")"
    
    RC = ""
    If txtCCost(2).Text <> "" Then RC = " codccost >='" & txtCCost(2).Text & "'"
    If txtCCost(3).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codccost <='" & txtCCost(3).Text & "'"
    End If
    
    
    'Si han puesto desde hasta cuenta
    If txtCta(29).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codmacta >='" & txtCta(29).Text & "'"
    End If
    
    If txtCta(30).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codmacta <='" & txtCta(30).Text & "'"
    End If
    
    
    'Cogemos prestada la tabla tmpCierre cargando las cuentas k
    'tengan en hpsaldanal y hpsaldana1 si asi lo recuieren las fechas
    SQL = "Delete  from tmpctaexpCC"
    Conn.Execute SQL
    
    
    
    'Insertamos las cuentas desde hpsald1 si hicera o hiciese falta
    Tablas = ""
    If Anyo1 < A1 Then
        Tablas = "SI"
    Else
        If Anyo1 = A1 Then
            'Dependera del mes
            If M1 > (Me.cmbFecha(5).ListIndex + 1) Then Tablas = "OK"
        End If
    End If

    If RC <> "" Then Cad = RC & " AND " & Cad
    SQL = "INSERT INTO tmpctaexpCC (codusu,cta,codccost) SELECT "
    SQL = SQL & vUsu.Codigo & ",codmacta,codccost from hsaldosanal"
    'Si es de hco
    If Tablas <> "" Then SQL = SQL & "1"
    SQL = SQL & " Where "
    SQL = SQL & Cad
    SQL = SQL & " GROUP BY codccost,codmacta"
    Conn.Execute SQL
    
    
    
    If Tablas <> "" Then CuentasDesdeHco
    
    
    
    
        
    'Diciembre 2012
    'Si ha marcado "solo" centros de reparto elimino aquellos eque n
    'Borro todos aquellos cc que no sean de reparto
    If chkCtaExpCC(2).Value = 1 Then
        SQL = "DELETE FROm tmpctaexpCC WHERE codusu = " & vUsu.Codigo & " AND "
        SQL = SQL & " NOT codccost IN (select distinct(subccost) from linccost) "
        Conn.Execute SQL
    End If
    
    
    'AHora en  tenemos todas las cuentas a tratar
    'Para ello cogeremos
    SQL = "Select count(*) from tmpctaexpCC where codusu = " & vUsu.Codigo
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not RS.EOF Then
        Cont = DBLet(RS.Fields(0), "N")
    End If
    RS.Close

    If Cont = 0 Then
        HacerCtaExploxCC = 1
        'MsgBox "Ningun registro a mostrar", vbExclamation
    Else
        'mes ini, ano ini, mes pedido, ano pedido, mes fin, ano fin
        'Cad = cmbFecha(5).ListIndex + 1 & "|" & txtAno(7).Text & "|"
        'Cad = Cad & cmbFecha(7).ListIndex + 1 & "|"
        Cad = cmbFecha(5).ListIndex + 1 & "|" & Anyo1 & "|"
        Cad = Cad & cmbFecha(7).ListIndex + 1 & "|"
        
        
        
        'El año del mes de calculo tiene k estar entre los años pedidos
        If cmbFecha(7).ListIndex >= cmbFecha(5).ListIndex Then
            Cad = Cad & Anyo1
        Else
            Cad = Cad & Anyo2
        End If
        Cad = Cad & "|"
        Cad = Cad & cmbFecha(6).ListIndex + 1 & "|" & Anyo2 & "|"
        
        'Ajusta los valores en modulo
        AjustaValoresCtaExpCC Cad
        
        'Si ha pediod los movimientos posteriores
        Post = (chkCtaExpCC(0).Value = 1)
        
        SQL = "Select cta,tmpctaexpCC.codccost,nommacta,nomccost from tmpctaexpCC,cuentas,cabccost where cuentas.codmacta=tmpctaexpCC.cta and cabccost.codccost=tmpctaexpCC.codccost and codusu = " & vUsu.Codigo
        'Vemos hasta donde hay de fechas en hco
        FechaFinEjercicio = CDate("01/" & M1 & "/" & A1)
        RS.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
        While Not RS.EOF
            
            Tablas = RS.Fields(0) & "|" & RS.Fields(1) & "|" & DevNombreSQL(RS.Fields(2))
            Tablas = Tablas & "|" & DevNombreSQL(RS.Fields(3)) & "|"
        
            'Tb ponemos la pb
            Label15.Caption = RS.Fields(0)
            Label15.Refresh
        
            CtaExploCentroCoste Tablas, Post, FechaFinEjercicio
    
            'Siguiente
            RS.MoveNext
        Wend
        RS.Close
        
        
        
        
        
        A1 = 0
        SQL = "Select count(*) from Usuarios.zctaexpcc where codusu =" & vUsu.Codigo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            A1 = DBLet(RS.Fields(0), "N")
        End If
        RS.Close
        If A1 = 0 Then
            'MsgBox "Ningun registro a mostrar", vbExclamation
            'GeneraCtaExplotacionCC = False
            HacerCtaExploxCC = 1
        Else
            HacerCtaExploxCC = 2
            'GeneraCtaExplotacionCC = True
        End If
        
    End If
    
    Exit Function
EGeneraCtaExplotacionCC:
    MuestraError Err.Number, "Genera Cta Explotacion CC"
End Function

'Proceso:    1.- Updatear codusu para que no borre los datos
'            2.- Mezclar los datos de otros actual /siguiente
'            3.- Borramos
Private Function ProcesoCtaExplotacionCC(Proceso As Byte) As Boolean

    ProcesoCtaExplotacionCC = False
    Select Case Proceso
    Case 0
        Cont = 0
        I = 0
        SQL = "Select min(codusu) from Usuarios.zctaexpcc"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then
                I = RS.Fields(0)
                Cont = 1 'para indicar que no es null
            End If
        End If
        RS.Close
        If I >= 10 Then
            I = 9
        Else
            If I = 0 Then
                'Si NO ERA NULL es que estan ocupados(cosa rara) desde el 0 al 9
                If Cont = 1 Then
                    MsgBox "Error inesperado. Descripcion: codusu entre 0..9", vbExclamation
                    Exit Function
                End If
                I = 9
            Else
                I = I - 1
            End If
        End If
        NumRegElim = I
        SQL = "UPDATE  Usuarios.zctaexpcc set codusu = " & NumRegElim & " WHERE codusu = " & vUsu.Codigo
        Conn.Execute SQL
    
    Case 1
        'UPDATEO los valores de codusu=vusu
        SQL = "UPDATE Usuarios.zctaexpcc  SET acumd=0,acumh=0,postd=0,posth=0 where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
        'Para los valores comarativos
        SQL = "UPDATE Usuarios.zctaexpcc  SET acumd=perid,acumh=perih,postd=saldod,posth=saldoh where codusu = " & NumRegElim
        Conn.Execute SQL
        SQL = "UPDATE Usuarios.zctaexpcc  SET perid=0,perih=0,saldod=0,saldoh=0 where codusu = " & NumRegElim
        Conn.Execute SQL
        
        'En RS cargo todas las referencias de codusu= vusu
        SQL = "Select * from Usuarios.zctaexpcc WHERE codusu = "
        RS.Open SQL & vUsu.Codigo, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        'Cojere Todas las referencias de la tabla zctaexpcc para numregelim
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL & NumRegElim, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            'Busco si tiene referncia en ACTUAL
            SQL = "codmacta = '" & miRsAux!codmacta & "' AND codccost ='" & DevNombreSQL(miRsAux!codccost) & "'"
            
            
            If Not EncontrarEn_zctaexpcc(miRsAux!codmacta, UCase(miRsAux!codccost)) Then
                'Updateo solo el codusu
                Cad = "UPDATE Usuarios.zctaexpcc  SET codusu = " & vUsu.Codigo & " WHERE codusu = " & NumRegElim & " AND " & SQL
            Else
                'UPDATEO
                Cad = "UPDATE Usuarios.zctaexpcc  SET acumd=" & TransformaComasPuntos(CStr(miRsAux!acumd))
                Cad = Cad & ", acumH=" & TransformaComasPuntos(CStr(miRsAux!acumh))
                Cad = Cad & ", postd=" & TransformaComasPuntos(CStr(miRsAux!postd))
                Cad = Cad & ", posth=" & TransformaComasPuntos(CStr(miRsAux!posth))
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND " & SQL
            End If
            Conn.Execute Cad
            'Siguiente
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        RS.Close
        
    Case 2
            'Finalmente borramos los datos de codusu=numregeleim
            SQL = "DELETE FROM Usuarios.zctaexpcc WHERE codusu = " & NumRegElim
            Conn.Execute SQL
        
            'Ahora, para que """SOLO""" aparezcan  los que tienen valor
            ' Los importes seran
            '                   MES                 SALDO
            '   actual      perid   perdih     saldod  saldoh
            '   anterior    acumd   acumh      postd   posth
        
            If optCCComparativo(1).Value Then
                'QUiero ver movimientos del periodo con lo cual me cargare aquellos
                ' que mov periodo en actual y anterior sea 0
                Cad = "acumd =0 and acumh=0 and perid=0 and perih=0"
                SQL = "DELETE FROM Usuarios.zctaexpcc WHERE codusu = " & vUsu.Codigo & " AND " & Cad
                Conn.Execute SQL
            End If
    End Select
    ProcesoCtaExplotacionCC = True
End Function

Private Function EncontrarEn_zctaexpcc(ByRef Cta As String, ByRef CC As String) As Boolean
    EncontrarEn_zctaexpcc = False
    RS.MoveFirst
    While Not RS.EOF
        If RS!codmacta = Cta And RS!codccost = CC Then
            EncontrarEn_zctaexpcc = True
            Exit Function
        Else
            RS.MoveNext
        End If
    Wend
    
        
End Function


Private Sub CuentasDesdeHco()


    'Haremos las inserciones desde hsaldosanal 1, es decir, ejercicios traspasados
    'si la fecha de incio de los calculos es  menor k la ultima fecha k haya en hco 1
    ' EN i tneemos el año y en mesfin1 el ultimo mes grabado en saldosanal1
    
    
    'SQL = "SELECT   codmacta,codccost from hsaldosanal1 where "
    SQL = "SELECT   codmacta,codccost from hsaldosanal1 as hsaldosanal where "
    If RC <> "" Then SQL = SQL & RC & " AND "
    SQL = SQL & Cad
    SQL = SQL & " group by codmacta"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "INSERT INTO tmpctaexpCC (codusu,cta,codccost) VALUES ("
    While Not miRsAux.EOF
        InseretaDesdeHCO SQL & vUsu.Codigo & ",'" & miRsAux.Fields(0) & "','" & miRsAux.Fields(1) & "');"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Sub InseretaDesdeHCO(ByRef Cuenta As String)
On Error Resume Next
    Conn.Execute Cuenta
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub





Private Function GeneraCCxCtaExplotacion() As Boolean
Dim A1 As Integer, M1 As Integer
Dim Post As Boolean

    On Error GoTo EGeneraCCxCtaExplotacion
    GeneraCCxCtaExplotacion = False
    
    UltimoMesAnyoAnal1 M1, A1
    
    'Si años consulta iguales
    If txtAno(9).Text = txtAno(10).Text Then
         Cad = " anoccost=" & txtAno(9).Text & " AND mesccost>=" & Me.cmbFecha(8).ListIndex + 1
         Cad = Cad & " AND mesccost<=" & Me.cmbFecha(9).ListIndex + 1
         
    Else
        'Años disitintos
        'Inicio
        Cad = "( anoccost=" & txtAno(9).Text & " AND mesccost>=" & Me.cmbFecha(8).ListIndex + 1 & ")"
        Cad = Cad & " OR ( anoccost=" & txtAno(10).Text & " AND mesccost<=" & Me.cmbFecha(9).ListIndex + 1 & ")"
        'Por si la diferencia es mas de un año
        If Val(txtAno(10).Text) - Val(txtAno(9).Text) > 1 Then
            Cad = Cad & " OR (anoccost >" & txtAno(9).Text & " AND anoccost < " & txtAno(10).Text & ")"
        End If
    End If
    Cad = " (" & Cad & ")"

    Tablas = ""
    If txtCta(14).Text <> "" Then Tablas = "codmacta >= '" & txtCta(14).Text & "'"
    If txtCta(15).Text <> "" Then
    If Tablas <> "" Then Tablas = Tablas & " AND "
     Tablas = Tablas & "codmacta <= '" & txtCta(15).Text & "'"
    End If
    
    RC = ""
    If txtCCost(4).Text <> "" Then RC = " hsaldosanal.codccost >='" & txtCCost(4).Text & "'"
    If txtCCost(5).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " hsaldosanal.codccost <='" & txtCCost(5).Text & "'"
    End If
    
    
    'Cogemos presta la tabla tmpCierre cargando las cuentas k
    'tengan en hpsaldanal y hpsaldana1 si asi lo recuieren las fechas
    SQL = "Delete  from tmpctaexpCC"
    Conn.Execute SQL
    
    If RC <> "" Then Cad = RC & " AND " & Cad
    If Tablas <> "" Then Cad = Cad & " AND " & Tablas
    
    SQL = "INSERT INTO tmpctaexpCC (codusu,cta,codccost) SELECT "
    SQL = SQL & vUsu.Codigo & ",codmacta,hsaldosanal.codccost from hsaldosanal"
    If txtAno(9).Text <= A1 Then
        If M1 <= Me.cmbFecha(9).ListIndex + 1 Then
            SQL = SQL & "1" 'ANALITICA EN CERRADOS
        End If
    End If
        
    SQL = SQL & " as hsaldosanal,cabccost where "
    SQL = SQL & " hsaldosanal.codccost = cabccost.codccost AND "
    SQL = SQL & Cad

    'Esta marcado solo los de reparto
    If Me.optCCxCta(1).Value Then SQL = SQL & " AND idsubcos <> 1"
        
    SQL = SQL & " group by codccost,codmacta"
    Conn.Execute SQL
    
    'Si estaba marcado el 2 entonces tendre k eliminar de la tabla tmpctaexpCC los datos
    'de   codccost que esten en linccost
    If Me.optCCxCta(2).Value Then
        Label2(26).Caption = "CC de reparto /"
        Label2(26).Refresh
        espera 0.2
        
        SQL = "Select distinct(subccost) from linccost"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            SQL = "DELETE FROM tmpctaexpCC where codusu = " & vUsu.Codigo & " AND codccost = '" & miRsAux.Fields(0) & "'"
            Conn.Execute SQL
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    'Insertamos las cuentas desde hpsald1 si hicerao hiciese falta
    Tablas = ""
    If Val(txtAno(9).Text) < A1 Then
        Tablas = "SI"
    Else
        If Val(txtAno(9).Text) = A1 Then
            'Dependera del mes
            If M1 > (Me.cmbFecha(8).ListIndex + 1) Then Tablas = "OK"
        End If
    End If
    
    If Tablas <> "" Then CuentasDesdeHco
    
    
    'AHora en  tenemos todas las cuentas a tratar
    'Para ello cogeremos
    SQL = "Select count(*) from tmpctaexpCC where codusu = " & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not miRsAux.EOF Then
        Cont = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Cont = 0 Then
        MsgBox "Ningun registro a mostrar", vbExclamation
    Else
        'mes ini, ano ini, mes pedido, ano pedido, mes fin, ano fin
        Cad = cmbFecha(8).ListIndex + 1 & "|" & txtAno(9).Text & "|"
        Cad = Cad & cmbFecha(10).ListIndex + 1 & "|"
        'El año del mes de calculo tiene k estar entre los años pedidos
        If cmbFecha(8).ListIndex >= cmbFecha(9).ListIndex Then
            Cad = Cad & txtAno(9).Text
        Else
            Cad = Cad & txtAno(10).Text
        End If
        Cad = Cad & "|"
        Cad = Cad & cmbFecha(9).ListIndex + 1 & "|" & txtAno(10).Text & "|"
        
        'Ajusta los valores en modulo
        AjustaValoresCtaExpCC Cad
        
        'Si ha pediod los movimientos posteriores
        Post = (chkCC_Cta.Value = 1)
        
        'Borramos datos
        SQL = "Delete from Usuarios.zctaexpcc where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
                
        
        SQL = "Select cta,tmpctaexpCC.codccost,nommacta,nomccost from tmpctaexpCC,cuentas,cabccost where cuentas.codmacta=tmpctaexpCC.cta and cabccost.codccost=tmpctaexpCC.codccost and codusu = " & vUsu.Codigo
        'Vemos hasta donde hay de fechas en hco
        FechaFinEjercicio = CDate("01/" & M1 & "/" & A1)
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
        While Not RS.EOF
            Tablas = RS.Fields(0) & "|" & RS.Fields(1) & "|" & DevNombreSQL(RS.Fields(2)) & "|" & DevNombreSQL(RS.Fields(3)) & "|"
            
            'Tb ponemos la pb
            Label2(26).Caption = RS.Fields(0)
            Label2(26).Refresh
    
            CtaExploCentroCoste Tablas, Post, FechaFinEjercicio
    
            'Siguiente
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        GeneraCCxCtaExplotacion = True
    End If
    
    'Contamos para ver si tiene datos
    If GeneraCCxCtaExplotacion Then
        A1 = 0
        Set miRsAux = New ADODB.Recordset
        SQL = "Select count(*) from Usuarios.zctaexpcc where codusu =" & vUsu.Codigo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            A1 = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        If A1 = 0 Then
            MsgBox "Ningun registro a mostrar", vbExclamation
            GeneraCCxCtaExplotacion = False
        End If
    End If
    Set miRsAux = Nothing
    Exit Function
EGeneraCCxCtaExplotacion:
    MuestraError Err.Number, "Genera C. coste por Cta. Explotacion" & vbCrLf & Err.Description
End Function




Private Function GenerarLibroResumen() As Boolean
Dim I2 As Currency

    On Error GoTo EGenerarLibroResumen
    GenerarLibroResumen = False
    
    'Eliminamos registros tmp
    SQL = "Delete FROM Usuarios.zdirioresum where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
        
    'Comprobamos k nivel
    For I = 1 To ChkNivelRes.Count
        If ChkNivelRes(I).visible Then
            If ChkNivelRes(I).Value Then
                Cont = I
                Exit For
            End If
        End If
    Next I
    
    
    
    
    I = Cont
    FijaValoresLibroResumen FechaIncioEjercicio, FechaFinEjercicio, I, EjerciciosCerrados, txtNumRes(0).Text
    
    Importe = 0
    I2 = 0
    If txtAno(11).Text = txtAno(12).Text Then
        I = CInt(Val(txtAno(11).Text))
        For Cont = cmbFecha(11).ListIndex + 1 To cmbFecha(12).ListIndex + 1
           Label2(25).Caption = "Fecha: " & Cont & " / " & I
           Label2(25).Refresh
           
           'Si ha puesto ACUMULADOS ANTERIORES
           If Cont = cmbFecha(11).ListIndex + 1 Then
                If txtNumRes(3).Text <> "" Then Importe = CCur(TransformaPuntosComas(txtNumRes(3).Text))
                If txtNumRes(4).Text <> "" Then I2 = CCur(TransformaPuntosComas(txtNumRes(4).Text))
           End If
           ProcesaLibroResumen Cont, I, Importe, I2
           Importe = 0
           I2 = 0
        Next Cont
    Else
        'Años partidos
        'El primer tramo de hasta fin de años
        I = CInt(Val(txtAno(11).Text))
        For Cont = cmbFecha(11).ListIndex + 1 To 12
           Label2(25).Caption = "Fecha: " & Cont & " / " & I
           Label2(25).Refresh
           If Cont = cmbFecha(11).ListIndex + 1 Then
                If txtNumRes(3).Text <> "" Then Importe = CCur(txtNumRes(3).Text)
                If txtNumRes(4).Text <> "" Then I2 = CCur(txtNumRes(4).Text)
           End If
           ProcesaLibroResumen Cont, I, Importe, I2
           Importe = 0: I2 = 0
        Next Cont
        'Años siguiente
        I = CInt(Val(txtAno(12).Text))
        For Cont = 1 To cmbFecha(12).ListIndex + 1
           Label2(25).Caption = "Fecha: " & Cont & " / " & I
           Label2(25).Refresh
           ProcesaLibroResumen Cont, I, Importe, I2
        Next Cont
    End If
    
    'Vemos si ha generado datos
    Set miRsAux = New ADODB.Recordset
    SQL = "Select count(*) from Usuarios.zdirioresum where codusu =" & vUsu.Codigo
    Cont = 0
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then Cont = miRsAux.Fields(0)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Cont = 0 Then
        MsgBox "Ningun dato generado para estos valores.", vbExclamation
        Exit Function
    End If
    
    Label2(25).Caption = ""
    Label2(25).Refresh
    GenerarLibroResumen = True
    Exit Function
EGenerarLibroResumen:
    MuestraError Err.Number, "Generar libro resumen"
End Function




Private Function ObtenerDatosCCCtaExp() As Boolean

On Error GoTo EObtenerDatosCCCtaExp
    ObtenerDatosCCCtaExp = False
    
    Label2(27).Caption = "Obteniendo conjunto registros"
    Label2(27).visible = True
    Me.Refresh

    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    Tablas = "hlinapu" & Tablas
    SQL = "Select cuentas.codmacta,cabccost.codccost,nommacta,nomccost FROM "
    SQL = SQL & Tablas
    SQL = SQL & ",cuentas,cabccost"
    SQL = SQL & " WHERE "
    SQL = SQL & Tablas & ".codmacta=cuentas.codmacta AND "
    SQL = SQL & Tablas & ".codccost=cabccost.codccost AND "
    'Fechas
    SQL = SQL & " fechaent >='" & Format(CDate(Text3(19).Text), FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <='" & Format(CDate(Text3(20).Text), FormatoFecha) & "'"
    'Si ha puesto ctas
    If txtCta(16).Text <> "" Then SQL = SQL & " AND cuentas.codmacta >='" & txtCta(16).Text & "'"
    If txtCta(17).Text <> "" Then SQL = SQL & " AND cuentas.codmacta <='" & txtCta(17).Text & "'"
    'Si ha puesto CC
    If txtCCost(6).Text <> "" Then SQL = SQL & " AND " & Tablas & ".codccost >='" & txtCCost(6).Text & "'"
    If txtCCost(7).Text <> "" Then SQL = SQL & " AND " & Tablas & ".codccost <='" & txtCCost(7).Text & "'"
    
    'K codccost no sea nulo
    SQL = SQL & " AND not (" & Tablas & ".codccost is null)"
    'Agrupado
    SQL = SQL & " group by cuentas.codmacta,codccost"
    
    'Ya tenemos el SQL
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then
        RS.Close
        MsgBox "Ningun dato entre estos parametros.", vbExclamation
        Exit Function
    End If

    'Preparar
    pb7.Value = 0
    pb7.visible = True
    Label2(27).Caption = "Preparando datos"
    Me.Refresh
    
    Cont = 0
    While Not RS.EOF
        Cont = Cont + 1
        RS.MoveNext
    Wend

    RS.MoveFirst
    DoEvents
    If PulsadoCancelar Then
        RS.Close
        Exit Function
    End If
        
    'Eliminamos datos
    SQL = "DELETE FROM Usuarios.zlinccexplo Where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    SQL = "DELETE FROM Usuarios.zcabccexplo WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    FijaValoresCtapoCC FechaIncioEjercicio, CDate(Text3(19).Text), CDate(Text3(20).Text), EjerciciosCerrados
    
    
    
    DoEvents
    If PulsadoCancelar Then
        RS.Close
        Exit Function
    End If
    
    RS.MoveFirst
    I = 1
    While Not RS.EOF
        DoEvents
        If PulsadoCancelar Then
                RS.Close
            Set RS = Nothing
            Exit Function
        End If
        'Los labels, progress y demas
        Label2(27).Caption = RS!nommacta
        Label2(27).Refresh
        pb7.Value = CInt((I / Cont) * pb7.Max)
        'Hacer accion
        SQL = RS!nommacta & "|" & RS!nomccost & "|"
        Cta_por_CC RS!codmacta, RS!codccost, SQL
        'Siguiente
        RS.MoveNext
        I = I + 1
    Wend
    RS.Close
    
    
    ObtenerDatosCCCtaExp = True
    Exit Function
EObtenerDatosCCCtaExp:
    MuestraError Err.Number
End Function



Private Function UltimaFechaHcoCabapu() As Date


UltimaFechaHcoCabapu = CDate("01/12/1900")
SQL = "Select max(fechaent) from hcabapu1"
Set RS = New ADODB.Recordset
RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then
        UltimaFechaHcoCabapu = Format(RS.Fields(0), "dd/mm/yyyy")
    End If
End If
RS.Close
Set RS = Nothing
End Function



Private Function ComprobarCuentas347_() As Boolean
Dim I As Integer
Dim I1 As Currency
Dim I2 As Currency
Dim I3 As Currency
Dim I4 As Currency
Dim I5 As Currency
Dim Pais As String
    ComprobarCuentas347_ = False
    
    'Esto sera para las inserciones de despues
    Tablas = "INSERT INTO Usuarios.z347 (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla,Provincia,pais) "
    Tablas = Tablas & " VALUES (" & vUsu.Codigo
         

    For I = 0 To List3.ListCount - 1
        Label2(30).Caption = List3.List(I)
        Label2(31).Caption = "Comprobar Cuentas"
        Me.Refresh
        If Not ComprobarCuentas347_DOS("Conta" & List3.ItemData(I), List3.List(I)) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        
    
       'Iremos NIF POR NIF
       
          Label2(31).Caption = "Insertando datos tmp(I)"
          Label2(31).Refresh
          SQL = "SELECT  cliprov,nif, sum(importe) as suma, razosoci,dirdatos,codposta,"
          SQL = SQL & "despobla,desprovi,pais from tmp347,Conta" & List3.ItemData(I) & ".cuentas where codusu=" & vUsu.Codigo
          'Primero ponemos todas, luego quitamos la linea de abajo
          'Importe = ImporteFormateado(txtImporte.Text)
          'SQL = SQL & " and cta=codmacta group by cliprov,nif having suma >=" & TransformaComasPuntos(CStr(Importe))
          SQL = SQL & " and cta=codmacta group by cliprov,nif"
          
          Set RS = New ADODB.Recordset
          RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
          
          While Not RS.EOF
               Label2(31).Caption = RS!NIF
               Label2(31).Refresh
               If ExisteEntrada Then
                    Importe = Importe + RS!Suma
                    SQL = "UPDATE Usuarios.z347 SET importe=" & TransformaComasPuntos(CStr(Importe))
                    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & RS!cliprov
                    SQL = SQL & " AND nif = '" & RS!NIF & "';"
               Else
                    'Nuevo para lo de las agencias de viajes
                    'SQL = "," & RS!cliprov & ",'" & RS!NIF & "'," & TransformaComasPuntos(CStr(RS!Suma))
                    SQL = "," & RS!cliprov & ",'" & RS!NIF & "'," & TransformaComasPuntos(CStr(RS!Suma))
                    SQL = SQL & ",'" & DevNombreSQL(DBLet(RS!razosoci, "T")) & "','" & DevNombreSQL(DBLet(RS!dirdatos)) & "','" & DBLet(RS!codposta, "T") & "','"
                    SQL = SQL & DevNombreSQL(DBLet(RS!despobla, "T")) & "','" & DevNombreSQL(DBLet(RS!desprovi, "T"))
                    If DBLet(RS!Pais, "T") = "" Then
                        Pais = "ESPAÑA"
                    Else
                        Pais = RS!Pais
                    End If
                    SQL = SQL & "','" & DevNombreSQL(DBLet(Pais, "T")) & "')"
                    SQL = Tablas & SQL
               End If
               Conn.Execute SQL
               RS.MoveNext
          Wend
          RS.Close
          
          
          'trimestral
          Label2(31).Caption = "Insertando datos tmp(II)"
          Label2(31).Refresh
          SQL = "SELECT  tmp347trimestre.cliprov,nif, sum(trim1) as t1, sum(trim2) as t2,"
          SQL = SQL & " sum(trim3) as t3, sum(trim4) as t4,sum(metalico) as metalico"
          SQL = SQL & " from tmp347,tmp347trimestre where tmp347.codusu=" & vUsu.Codigo
          SQL = SQL & " and tmp347.codusu=tmp347trimestre.codusu"
          SQL = SQL & " and tmp347.cliprov=tmp347trimestre.cliprov"
          SQL = SQL & " and tmp347.cta=tmp347trimestre.cta group by tmp347.cliprov,nif"
          
          Set RS = New ADODB.Recordset
          RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
          
          While Not RS.EOF
               Label2(31).Caption = RS!NIF
               Label2(31).Refresh
               If ExisteEntradaTrimestral(I1, I2, I3, I4, I5) Then
                    I1 = I1 + RS!T1
                    I2 = I2 + RS!t2
                    I3 = I3 + RS!T3
                    I4 = I4 + RS!T4
                    I5 = I5 + RS!metalico
                    SQL = "UPDATE Usuarios.z347trimestral SET "
                    SQL = SQL & " trim1=" & TransformaComasPuntos(CStr(I1))
                    SQL = SQL & ", trim2=" & TransformaComasPuntos(CStr(I2))
                    SQL = SQL & ", trim3=" & TransformaComasPuntos(CStr(I3))
                    SQL = SQL & ", trim4=" & TransformaComasPuntos(CStr(I4))
                    SQL = SQL & ", metalico=" & TransformaComasPuntos(CStr(I5))
                    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & RS!cliprov
                    SQL = SQL & " AND nif = '" & RS!NIF & "';"
               Else
                    
                    SQL = "insert into Usuarios.z347trimestral (`codusu`,`cliprov`,`nif`,`trim1`,`trim2`"
                    SQL = SQL & ",`trim3`,`trim4`,metalico) values ( " & vUsu.Codigo
                    SQL = SQL & "," & RS!cliprov & ",'" & RS!NIF & "',"
                    SQL = SQL & TransformaComasPuntos(CStr(RS!T1)) & "," & TransformaComasPuntos(CStr(RS!t2)) & ","
                    SQL = SQL & TransformaComasPuntos(CStr(RS!T3)) & "," & TransformaComasPuntos(CStr(RS!T4)) & ","
                    SQL = SQL & TransformaComasPuntos(CStr(RS!metalico)) & ")"

               End If
               Conn.Execute SQL
               RS.MoveNext
          Wend
          RS.Close
          
          
          
          
          
          espera 0.5
    Next I
    ComprobarCuentas347_ = True
    
End Function




Private Function ComprobarCuentas347DatosExternos() As Boolean
Dim I As Integer


    ComprobarCuentas347DatosExternos = False

    
    'Esto sera para las inserciones de despues
    Tablas = "INSERT INTO Usuarios.z347 (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla,Provincia) "
    Tablas = Tablas & " VALUES (" & vUsu.Codigo
         
    Set miRsAux = New ADODB.Recordset
    For I = 0 To List3.ListCount - 1
        Label2(30).Caption = "EXT: " & List3.List(I)
        Label2(31).Caption = "Comprobar Cuentas"
        Me.Refresh
    
       'Iremos NIF POR NIF
       
           
          SQL = "SELECT ascii(letra) as cliprov, nif, nombre as razosoci, direc as dirdatos, codposta,"
          SQL = SQL & " poblacion as despobla, provincia as desprovi, importe as suma"
          SQL = SQL & " from datosext347 where año =" & Year(CDate(Text3(21).Text))

          
          Set RS = New ADODB.Recordset
          RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
          
          While Not RS.EOF
               Label2(31).Caption = RS!NIF
               Label2(31).Refresh
               If ExisteEntrada Then
                    Importe = Importe + RS!Suma
                    SQL = "UPDATE Usuarios.z347 SET importe=" & TransformaComasPuntos(CStr(Importe))
                    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & RS!cliprov
                    SQL = SQL & " AND nif = '" & RS!NIF & "';"
               Else
                    SQL = "," & RS!cliprov & ",'" & RS!NIF & "'," & TransformaComasPuntos(CStr(RS!Suma))
                    SQL = SQL & ",'" & DevNombreSQL(DBLet(RS!razosoci, "T")) & "','" & DevNombreSQL(DBLet(RS!dirdatos)) & "','" & DBLet(RS!codposta, "T") & "','"
                    SQL = SQL & DevNombreSQL(DBLet(RS!despobla, "T")) & "','" & DevNombreSQL(DBLet(RS!desprovi, "T")) & "')"
                    SQL = Tablas & SQL
               End If
               Conn.Execute SQL
               RS.MoveNext
          Wend
          DoEvents
          RS.Close
          Set RS = Nothing
          espera 0.5
    Next I
    
    
    ComprobarCuentas347DatosExternos = True
    
End Function








Private Function ComprobarCuentas347_DOS(Contabilidad As String, Empresa As String) As Boolean
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim I1 As Currency
Dim I2 As Currency
Dim I3 As Currency
Dim Trimestre(3) As Currency
Dim Impor As Currency
Dim Tri As Byte

On Error GoTo EComprobarCuentas347
    ComprobarCuentas347_DOS = False
    'Utilizaremos la tabla tmpcierre1, prestada
    SQL = "DELETE FROM tmp347 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    'Enero 2012
    'Calcular por trimiestre
    SQL = "DELETE FROM tmp347trimestre where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    
    
'    'Cargamos la tabla con los valores
'    SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe)   SELECT " & vUsu.Codigo
'    SQL = SQL & " ,0,cabfact.codmacta,nifdatos,sum(totfaccl)as suma from " & Contabilidad & ".cabfact," & Contabilidad & ".cuentas  where "
'    SQL = SQL & " cuentas.codmacta=cabfact.codmacta and model347=1 "
'    SQL = SQL & " AND fecfaccl >='" & Format(Text3(21).Text, FormatoFecha) & "'"
'    SQL = SQL & " AND fecfaccl <='" & Format(Text3(22).Text, FormatoFecha) & "'"
'    SQL = SQL & " group by codmacta "
'    Conn.Execute SQL
    
    Set RS = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    
    'Para lo nuevo. Iremos codmacta a codmacta


   
    SQL = " Select cabfact.codmacta,nifdatos from " & Contabilidad & ".cabfact," & Contabilidad & ".cuentas  where "
    SQL = SQL & " cuentas.codmacta=cabfact.codmacta and model347=1 "
    SQL = SQL & " AND fecfaccl >='" & Format(Text3(21).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecfaccl <='" & Format(Text3(22).Text, FormatoFecha) & "'"
    SQL = SQL & " group by codmacta "
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
    
        Label2(31).Caption = RS!codmacta
        Label2(31).Refresh
        
        Trimestre(0) = 0: Trimestre(1) = 0: Trimestre(2) = 0: Trimestre(3) = 0
        
        
        SQL = "Select * from " & Contabilidad & ".cabfact where codmacta = '" & RS.Fields(0) & "'"
        'Modificacion 15 Febrero 2005  Coframa. Si ha marcado fecliq aqui cojo fecha liq
        If chk347(0).Value = 1 Then
            SQL = SQL & " AND fecliqcl >='" & Format(Text3(21).Text, FormatoFecha) & "'"
            SQL = SQL & " AND fecliqcl <='" & Format(Text3(22).Text, FormatoFecha) & "'"
        Else
            SQL = SQL & " AND fecfaccl >='" & Format(Text3(21).Text, FormatoFecha) & "'"
            SQL = SQL & " AND fecfaccl <='" & Format(Text3(22).Text, FormatoFecha) & "'"
        End If
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I1 = 0
        I2 = 0
        While Not RT.EOF
            'falta sumar
        
        
            'La primera base siempre la
            I1 = I1 + RT!ba1faccl
            I2 = I2 + RT!ti1faccl
            
            Impor = RT!ba1faccl + RT!ti1faccl
            
            'Si tiene tipo iva 2
            If Not IsNull(RT!tp2faccl) Then
                I1 = I1 + RT!ba2faccl
                I2 = I2 + RT!ti2faccl
                Impor = Impor + RT!ba2faccl + RT!ti2faccl
            End If
            
            If Not IsNull(RT!tp3faccl) Then
                I1 = I1 + RT!ba3faccl
                I2 = I2 + RT!ti3faccl
                Impor = Impor + RT!ba3faccl + RT!ti3faccl
            End If
            
            'Nuevo 25 Febrero 2008.
            'El recargo de equivalencia entra dentro tb.
            If Not IsNull(RT!tr1faccl) Then
                I2 = I2 + RT!tr1faccl
                Impor = Impor + RT!tr1faccl
            End If
            If Not IsNull(RT!tr2faccl) Then
                I2 = I2 + RT!tr2faccl
                Impor = Impor + RT!tr2faccl
            End If
            If Not IsNull(RT!tr3faccl) Then
                I2 = I2 + RT!tr3faccl
                Impor = Impor + RT!tr3faccl
            End If
            
            
            'El trimestre
            Tri = QueTrimestre(RT!fecliqcl)
            Tri = Tri - 1
            Trimestre(Tri) = Trimestre(Tri) + Impor
            
            RT.MoveNext
        Wend
        RT.Close
        
        'El importe final es la suma de las bases mas los ivas
        I1 = I1 + I2
        SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe)  "
        'SQL = SQL & " VALUES (" & vUsu.Codigo & ",0,'" & RS!Codmacta & "','"
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("0") & ",'" & RS!codmacta & "','"
        SQL = SQL & DBLet(RS!nifdatos) & "'," & TransformaComasPuntos(CStr(I1)) & ")"
        Conn.Execute SQL
        
        
        'El del trimestre
        SQL = "insert into `tmp347trimestre` (`codusu`,`cliprov`,`cta`,`trim1`,`trim2`,`trim3`,`trim4`)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("0") & ",'" & RS!codmacta & "'"
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(0))) & "," & TransformaComasPuntos(CStr(Trimestre(1)))
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(2))) & "," & TransformaComasPuntos(CStr(Trimestre(3))) & ")"
        Conn.Execute SQL

        
        
        RS.MoveNext
    Wend
    RS.Close
    If OptProv(0).Value Then
        Cad = "fecrecpr"
    Else
        If OptProv(1).Value Then
            Cad = "fecfacpr"
        Else
            Cad = "fecliqpr"
        End If
    End If
'    SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe)   SELECT " & vUsu.Codigo
'    SQL = SQL & " ,1,cabfactprov.codmacta,nifdatos,sum(totfacpr)as suma from " & Contabilidad & ".cabfactprov," & Contabilidad & ".cuentas  where "
'    SQL = SQL & " cuentas.codmacta=cabfactprov.codmacta and model347=1 "
'    SQL = SQL & " AND " & Cad & " >='" & Format(Text3(21).Text, FormatoFecha) & "'"
'    SQL = SQL & " AND " & Cad & " <='" & Format(Text3(22).Text, FormatoFecha) & "'"
'    SQL = SQL & " group by codmacta "
'    Conn.Execute SQL
'
    
    
    Label2(31).Caption = "Comprobando datos facturas proveedor"
    DoEvents
    espera 0.2
    
    
    SQL = "SELECT cabfactprov.codmacta,nifdatos from " & Contabilidad & ".cabfactprov," & Contabilidad & ".cuentas  where "
    SQL = SQL & " cuentas.codmacta=cabfactprov.codmacta and model347=1 "
    SQL = SQL & " AND " & Cad & " >='" & Format(Text3(21).Text, FormatoFecha) & "'"
    SQL = SQL & " AND " & Cad & " <='" & Format(Text3(22).Text, FormatoFecha) & "'"
    SQL = SQL & " group by codmacta "
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Label2(31).Caption = RS.Fields(0)
        Label2(31).Refresh
        DoEvents
        SQL = "Select cabfactprov.*," & Cad & " fecha from " & Contabilidad & ".cabfactprov cabfactprov where codmacta = '" & RS.Fields(0) & "'"
        SQL = SQL & " AND " & Cad & " >='" & Format(Text3(21).Text, FormatoFecha) & "'"
        SQL = SQL & " AND " & Cad & " <='" & Format(Text3(22).Text, FormatoFecha) & "'"
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I1 = 0
        I2 = 0
        Trimestre(0) = 0: Trimestre(1) = 0: Trimestre(2) = 0: Trimestre(3) = 0
        While Not RT.EOF
            'La primera base siempre la
            I1 = I1 + RT!ba1facpr
            I2 = I2 + RT!ti1facpr
            Impor = RT!ba1facpr + RT!ti1facpr
            'Si tiene tipo iva 2
            If Not IsNull(RT!tp2facpr) Then
                I1 = I1 + RT!ba2facpr
                I2 = I2 + RT!ti2facpr
                Impor = Impor + RT!ba2facpr + RT!ti2facpr
            End If

            If Not IsNull(RT!tp3facpr) Then
                I1 = I1 + RT!ba3facpr
                I2 = I2 + RT!ti3facpr
                Impor = Impor + RT!ba3facpr + RT!ti3facpr
            End If


            'Nuevo 25 Febrero 2008.
            'El recargo de equivalencia entra dentro tb.
            If Not IsNull(RT!tr1facpr) Then
                I2 = I2 + RT!tr1facpr
                Impor = Impor + RT!tr1facpr
            End If
            If Not IsNull(RT!tr2facpr) Then
                I2 = I2 + RT!tr2facpr
                Impor = Impor + RT!tr2facpr
            End If
            If Not IsNull(RT!tr3facpr) Then
                I2 = I2 + RT!tr3facpr
                Impor = Impor + RT!tr3facpr
            End If
                
            
            
                
            'El trimestre
            Tri = QueTrimestre(RT!Fecha)
            Tri = Tri - 1
            Trimestre(Tri) = Trimestre(Tri) + Impor
            
            
            
            RT.MoveNext
        Wend
        RT.Close
        
        'El importe final es la suma de las bases mas los ivas
        I1 = I1 + I2
        SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe)  "
        'SQL = SQL & " VALUES (" & vUsu.Codigo & ",1,'" & RS!Codmacta & "','"
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("1") & ",'" & RS!codmacta & "','"
        SQL = SQL & DBLet(RS!nifdatos) & "'," & TransformaComasPuntos(CStr(I1)) & ")"
        Conn.Execute SQL
        
        
        
        
        'El del trimestre
        SQL = "insert into `tmp347trimestre` (`codusu`,`cliprov`,`cta`,`trim1`,`trim2`,`trim3`,`trim4`)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("1") & ",'" & RS!codmacta & "'"
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(0))) & "," & TransformaComasPuntos(CStr(Trimestre(1)))
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(2))) & "," & TransformaComasPuntos(CStr(Trimestre(3))) & ")"
        Conn.Execute SQL
        
        
        RS.MoveNext
        
    Wend
    RS.Close
    
    
    
    
    'DICIEMBRE 2012
    ' CObros en metalico superiores a 6000
    Label2(31).Caption = "Cobros metalico"
    Label2(31).Refresh
    DoEvents
    SQL = "Select ImporteMaxEfec340 from " & Contabilidad & ".parametros "
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO pues ser eof
    I1 = DBLet(RS!ImporteMaxEfec340, "N")
    RS.Close
    If I1 > 0 Then
        'SI que lleva control de cobros en efectivo
        'Veremos si hay conceptos de efectivo
        SQL = "Select codconce from " & Contabilidad & ".conceptos where EsEfectivo340 = 1"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not RS.EOF
            SQL = SQL & ", " & RS!codconce
            RS.MoveNext
        Wend
        RS.Close
        SQL2 = "" 'Errores en Datos en efectivo sin ventas
        If SQL <> "" Then
            SQL = Mid(SQL, 2) 'quit la coma
            
            Cad = "Select * from tmp347trimestre WHERE codusu = " & vUsu.Codigo & " ORDER BY cta"
            RT.Open Cad, Conn, adOpenKeyset, adCmdText
            
            'HABER -DEBE"
            Cad = "Select hlinapu.codmacta,sum(if(timporteh is null,0,timporteh))-sum(if(timported is null,0,timported)) importe"
            Cad = Cad & " from " & Contabilidad & ".hlinapu,cuentas WHERE hlinapu.codmacta =cuentas.codmacta "
            Cad = Cad & " AND model347=1 AND fechaent >='" & Format(Text3(21).Text, FormatoFecha) & "'"
            Cad = Cad & " AND fechaent <='" & Format(Text3(22).Text, FormatoFecha) & "'"
            Cad = Cad & " AND codconce IN (" & SQL & ")"
            Cad = Cad & " group by 1 order by 1"

            RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS.EOF
                Label2(31).Caption = RS!codmacta
                Label2(31).Refresh
        
                If RS!Importe <> 0 Then
                    SQL = "cta  = '" & RS!codmacta & "'"
                    RT.Find SQL, , adSearchForward, 1
                    
                    If RT.EOF Then
                        SQL2 = SQL2 & RS!codmacta & " (" & RS!Importe & ") " & vbCrLf
                    Else
                        SQL = "UPDATE tmp347trimestre SET metalico = " & TransformaComasPuntos(CStr(RS!Importe))
                        SQL = SQL & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & RT!Cta & "'"
                        Conn.Execute SQL
                    End If
                End If
                RS.MoveNext
            Wend
            RS.Close
            RT.Close
            
            If SQL2 <> "" Then
                SQL2 = "Cobros en efectivo sin asociar a ventas" & vbCrLf & SQL2
                MsgBox SQL2, vbExclamation
            End If
        End If
    End If
    
    Set RT = Nothing
    RC = ""
    Cad = ""
    SQL2 = ""
    'Comprobaremos k el nif no es nulo, ni el codppos de las cuentas a tratar
    SQL = "Select cta from tmp347 where (nif is null or nif = '') and codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        Cad = Cad & RS.Fields(0) & "       "
        If I = 3 Then
            Cad = Cad & vbCrLf
            I = 0
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    If Cad <> "" Then
        RC = "Cuentas con NIF sin valor: " & vbCrLf & vbCrLf & Cad
        Cad = ""
    End If
    
    'Comprobamos el codpos
    SQL = "Select cta,nommacta,codposta from tmp347," & Contabilidad & ".cuentas where codusu = " & vUsu.Codigo
    SQL = SQL & " AND tmp347.cta=cuentas.codmacta and (codposta is null or codposta='')"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        Cad = Cad & RS.Fields(0) & "       "
        If I = 3 Then
            Cad = Cad & vbCrLf
            I = 0
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    If Cad <> "" Then
        If RC <> "" Then RC = RC & vbCrLf & vbCrLf & vbCrLf
        RC = RC & "Cuentas con codigo postal sin valor: " & vbCrLf & vbCrLf & Cad
    End If
    
    If RC <> "" Then
        RC = "Empresa: " & Empresa & vbCrLf & vbCrLf & RC & vbCrLf & " Desea continuar igualmente?"
        If MsgBox(RC, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    
    Set RS = Nothing
    
    ComprobarCuentas347_DOS = True
    Exit Function
EComprobarCuentas347:
    MuestraError Err.Number, "Comprobar Cuentas 347" & vbCrLf & vbCrLf & SQL & vbCrLf
End Function

'Dada una fecha me da el trimestre
Private Function QueTrimestre(Fecha As Date) As Byte
Dim C As Byte
    
        C = Month(Fecha)
        If C < 4 Then
            QueTrimestre = 1
        ElseIf C < 7 Then
            QueTrimestre = 2
        ElseIf C < 10 Then
            QueTrimestre = 3
        Else
            QueTrimestre = 4
        End If
    
End Function
Private Function ExisteEntrada() As Boolean
    SQL = "Select importe from Usuarios.z347  where codusu = " & vUsu.Codigo & " and cliprov =" & RS!cliprov & " AND nif ='" & RS!NIF & "';"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ExisteEntrada = True
        Importe = miRsAux!Importe
    Else
        ExisteEntrada = False
    End If
    miRsAux.Close
End Function

Private Function ExisteEntradaTrimestral(ByRef I1 As Currency, ByRef I2 As Currency, ByRef I3 As Currency, ByRef I4 As Currency, ByRef I5 As Currency) As Boolean
    SQL = "Select trim1,trim2,trim3,trim4,metalico from Usuarios.z347trimestral  where codusu = " & vUsu.Codigo & " and cliprov =" & RS!cliprov & " AND nif ='" & RS!NIF & "';"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ExisteEntradaTrimestral = True
        I1 = miRsAux!trim1
        I2 = miRsAux!trim2
        I3 = miRsAux!trim3
        I4 = miRsAux!trim4
        I5 = DBLet(miRsAux!metalico, "N")
    Else
        ExisteEntradaTrimestral = False
        I1 = 0: I2 = 0: I3 = 0: I4 = 0: I5 = 0
    End If
    miRsAux.Close
End Function

Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function


Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function




Private Sub HacerBorreFacturas()
Dim Ejecutar As String

On Error GoTo EHacerBorreFacturas
    
    SQL = ""
    
            '----------SERIE
            If txtSerie(2).Text <> "" Then SQL = SQL & " AND numserie >= '" & txtSerie(2).Text & "'"
            If txtSerie(3).Text <> "" Then SQL = SQL & " AND numserie <= '" & txtSerie(3).Text & "'"
            
            
            'Numero factura
            If Opcion = 22 Then
                RC = "codfaccl"
            Else
                RC = "numregis"
            End If
            If txtNumFac(2).Text <> "" Then SQL = SQL & " AND " & RC & " >=" & txtNumFac(2).Text
            If txtNumFac(3).Text <> "" Then SQL = SQL & " AND " & RC & " <=" & txtNumFac(3).Text
        
            
            'Fecha factura
            If Opcion = 22 Then
                RC = "fecfaccl"
            Else
                RC = "fecrecpr"
            End If
            If Text3(23).Text <> "" Then SQL = SQL & " AND " & RC & " >='" & Format(Text3(23).Text, FormatoFecha) & "'"
            If Text3(24).Text <> "" Then SQL = SQL & " AND " & RC & " <='" & Format(Text3(24).Text, FormatoFecha) & "'"
            
    
    
    'Tablas de trabajo
    Tablas = ""
    If Opcion = 23 Then Tablas = "prov"
    
    'Comprobamos k no existen datos con las fechas para los intervalos solicitados
    SQL = Mid(SQL, 5) 'Quitamos el primer and
    Set RS = New ADODB.Recordset
    
    'Comprobamos que existen datos para traspasar
    Cad = "Select count(*) from cabfact" & Tablas & " WHERE " & SQL
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not RS.EOF Then Cont = DBLet(RS.Fields(0), "N")
    RS.Close
    
    If Cont = 0 Then
        MsgBox "Ningun dato a traspasar con esos parametros", vbExclamation
        Exit Sub
    End If
    
    

    
    Cad = "Select * from cabfact" & Tablas & "1 WHERE " & SQL
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then I = 1
    RS.Close
    
    
    If I > 0 Then
        If MsgBox("Ya existen datos en facturas traspasadas para estos intervalos. ¿Desea continuar ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    
    'Comprobamos que todas tienen numasien
    Cad = "SELECT count(*) FROM cabfact" & Tablas & " WHERE " & SQL
    Cad = Cad & " and numasien is null "
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    If Not RS.EOF Then I = DBLet(RS.Fields(0), "N")
    RS.Close
    If I > 0 Then
        Cad = "Algunas de las facturas no han sido contabilizadas." & vbCrLf & Space(20) & "¿Continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    Cad = "SELECT count(*) FROM cabfact" & Tablas & " WHERE " & SQL
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not RS.EOF Then Cont = DBLet(RS.Fields(0), "N")
    RS.Close
    If Cont = 0 Then
        MsgBox "Ningun dato a traspasar.", vbExclamation
    Else
        Cad = "Las facturas serán traspasadas y borradas. " & vbCrLf & "   ¿Seguro que desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Cont = 0
    End If
    If Cont = 0 Then Exit Sub
    
    Cad = "SELECT * FROM cabfact" & Tablas & " WHERE " & SQL
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Para las lineas
    Ejecutar = "INSERT INTO linfact" & Tablas & "1 SELECT * FROM linfact" & Tablas & " WHERE "
    'Para las cabeceras
    BACKUP_TablaIzquierda RS, SQL
    
    I = 0
    pb8.Value = 0
    pb8.visible = True
    While Not RS.EOF
          If Opcion = 22 Then
            Cad = " codfaccl =" & RS!codfaccl & " AND anofaccl =" & RS!anofaccl
          Else
            Cad = " numregis =" & RS!NumRegis & " AND anofacpr =" & RS!anofacpr
          End If
          RC = Ejecutar & Cad
          Conn.Execute RC
    
          BACKUP_Tabla RS, RC
          RC = "INSERT INTO cabfact" & Tablas & "1 " & SQL & " VALUES " & RC
          Conn.Execute RC
          
          
          'Borrar
          RC = "DELETE FROM linfact" & Tablas & " WHERE " & Cad
          Conn.Execute RC
          
          RC = "DELETE FROM cabfact" & Tablas & " WHERE " & Cad
          Conn.Execute RC
          
          RS.MoveNext
          I = I + 1
          pb8.Value = Round((I / Cont) * 1000, 0)
          If (I Mod 100) = 0 Then
                Me.Refresh
                DoEvents
          End If
    Wend
    RS.Close
    Exit Sub
EHacerBorreFacturas:
    MuestraError Err.Number, "HacerBorreFacturas"
End Sub


Private Function ComparaFechasCombos(Indice1 As Integer, Indice2 As Integer, InCombo1 As Integer, InCombo2 As Integer) As Boolean
    ComparaFechasCombos = False
    If txtAno(Indice1).Text <> "" And txtAno(Indice2).Text <> "" Then
        If Val(txtAno(Indice1).Text) > Val(txtAno(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        Else
            If Val(txtAno(Indice1).Text) = Val(txtAno(Indice2).Text) Then
                If Me.cmbFecha(InCombo1).ListIndex > Me.cmbFecha(InCombo2).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    ComparaFechasCombos = True
End Function




Private Sub PonerBalancePredeterminado()

    'El balance de P y G tiene el campo Perdidas=1
    Select Case Opcion
    Case 27, 39
        I = 1
    Case Else
        I = 0
    End Select
    If Opcion >= 50 Then
        Cont = 1
    Else
        Cont = 0
    End If
    SQL = "Select * from sbalan where predeterminado = 1 AND perdidas =" & I
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Me.txtNumBal(Cont).Text = RS.Fields(0)
        TextDescBalance(Cont).Text = RS.Fields(1)
    End If
    RS.Close
    Set RS = Nothing
    Cont = 0
End Sub






Private Function ComprobarCuentas349(ByRef C1 As Integer, ByRef C2 As Integer) As Boolean
Dim I As Integer
Dim Trim(3) As Currency


'Contadores para facturas de abono

    ComprobarCuentas349 = False
    SQL = "DELETE FROM Usuarios.z347 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    'Para el listado de facturas utilizaremos los datos
    SQL = "DELETE FROM Usuarios.ztmpfaclin WHERE codusu =" & vUsu.Codigo
    Conn.Execute SQL
    SQL = "DELETE FROM Usuarios.ztmpfaclinprov WHERE codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    'AGOSTO 2015
    SQL = "DELETE FROM Usuarios.ztesoreriacomun WHERE codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    C1 = 0
    C2 = 0

    
    'Esto sera para las inserciones de despues
    Tablas = "INSERT INTO Usuarios.z347 (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla) "
    Tablas = Tablas & " VALUES (" & vUsu.Codigo
         
    Set miRsAux = New ADODB.Recordset
    For I = 0 To List5.ListCount - 1
        If Not ComprobarCuentas349_DOS("Conta" & List5.ItemData(I), C1, C2) Then
            Set miRsAux = Nothing
            Exit Function
        End If
    
    
    
    
    
       'Iremos NIF POR NIF
       
           
          SQL = "SELECT  cliprov,nif, sum(importe) as suma, razosoci,dirdatos,codposta,"
          'Modificacion MARZO 2009 en lugar de despobla pondre pais
          SQL = SQL & "pais despobla from tmp347,Conta" & List5.ItemData(I) & ".cuentas where codusu=" & vUsu.Codigo
          SQL = SQL & " and cta=codmacta group by cliprov,nif "
          
          Set RS = New ADODB.Recordset
          RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
          While Not RS.EOF
               If ExisteEntrada Then
                    Importe = Importe + RS!Suma
                    SQL = "UPDATE Usuarios.z347 SET importe=" & TransformaComasPuntos(CStr(Importe))
                    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & RS!cliprov
                    SQL = SQL & " AND nif = '" & RS!NIF & "';"
               Else
                    
                    SQL = "," & RS!cliprov & ",'" & RS!NIF & "'," & TransformaComasPuntos(CStr(RS!Suma))
                    SQL = SQL & ",'" & DevNombreSQL(DBLet(RS!razosoci)) & "','" & DevNombreSQL(DBLet(RS!dirdatos)) & "','" & RS!codposta & "','" & DevNombreSQL(DBLet(RS!despobla)) & "')"
                    SQL = Tablas & SQL
               End If
               Conn.Execute SQL
               RS.MoveNext
          Wend
          RS.Close
    Next I
    
    'Agosto 2015
    'Para el Listview que selecciona tipo factura.
    'Para saber si son ventas o compras
    SQL = "UPDATE usuarios.ztesoreriacomun set importe2=opcion where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    'Comprobamos si hay datos
    SQL = "Select count(*) FROM Usuarios.z347 where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            Cont = miRsAux.Fields(0)
        End If
    End If
    miRsAux.Close
    
    If Cont = 0 Then
        If Me.chk349.Value = 1 Then
            'Listado
            MsgBox "Ningún dato se ha generado con esos valores", vbExclamation
        Else
            'DEjo continuar
            ComprobarCuentas349 = True
        End If
    Else
            
        'Por si quiere cambiar las claves de las operaciones de las intracomunitarias
        frmIntracom349.Show vbModal
        SQL = DevuelveDesdeBD("count(*)", "Usuarios.z347", "codusu", CStr(vUsu.Codigo))
        If Val(SQL) > 0 Then ComprobarCuentas349 = True
    End If
    Set miRsAux = Nothing
    
End Function



Private Function ComprobarCuentas349_DOS(Contabilidad As String, ByRef ContadorCli As Integer, ByRef ContadorPro As Integer) As Boolean
Dim ContadorInserciones As Integer

On Error GoTo EComprobarCuentas349

    ComprobarCuentas349_DOS = False
    'Utilizaremos la tabla tmpcierre1, prestada
    SQL = "DELETE FROM tmp347 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    
    SQL = DevuelveDesdeBD("max(codigo)", "usuarios.ztesoreriacomun ", "codusu", CStr(vUsu.Codigo))
    If SQL = "" Then SQL = "0"
    ContadorInserciones = CInt(SQL)
    
    'Agosto 2015
    'Cargamos las facturas para que puedan asignarle una clave a mano en un frm
    SQL = "SELECT  " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,'" & Contabilidad & "',0,cabfact.codmacta,nifdatos,concat(numserie,codfaccl),"
    SQL = SQL & " nommacta,ba1faccl + coalesce(ba2faccl,0) + coalesce(ba3faccl,0) ,fecfaccl,substring(coalesce(pais,''),1,2)"
    SQL = SQL & "  from " & Contabilidad & ".cabfact," & Contabilidad & ".cuentas , (SELECT @rownum:=" & ContadorInserciones & ") r"
    SQL = SQL & " where cuentas.codmacta=cabfact.codmacta "
    SQL = SQL & " AND fecfaccl >='" & Format(Text3(26).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecfaccl <='" & Format(Text3(27).Text, FormatoFecha) & "'"
    'Factura extranjero
    SQL = SQL & " AND intracom=1"
    
    'Pero si tiene serie de AUTOFACTURAS, la quitamos
    If txtSerie(4).Text <> "" Then SQL = SQL & " AND numserie <> '" & txtSerie(4).Text & "'"
    
    SQL = "INSERT INTO usuarios.ztesoreriacomun(codusu,codigo,texto1,opcion,texto2,texto3,texto4,texto5,importe1,fecha1,texto6) " & SQL
    Conn.Execute SQL
    
    
    'Cargamos la tabla con los valores
    SQL = "SELECT "
'    SQL = SQL & " ,0,cabfact.codmacta,nifdatos,sum(totfaccl)as suma from " & Contabilidad & ".cabfact," & Contabilidad & ".cuentas  where "
    SQL = SQL & " cabfact.codmacta,nifdatos,sum(ba1faccl)as s1,sum(ba2faccl) as s2,sum(ba3faccl) as s3  from " & Contabilidad & ".cabfact," & Contabilidad & ".cuentas  where "
    SQL = SQL & " cuentas.codmacta=cabfact.codmacta "
    SQL = SQL & " AND fecfaccl >='" & Format(Text3(26).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecfaccl <='" & Format(Text3(27).Text, FormatoFecha) & "'"
    'Factura extranjero
    SQL = SQL & " AND intracom=1"
    
    'Pero si tiene serie de AUTOFACTURAS, la quitamos
    If txtSerie(4).Text <> "" Then SQL = SQL & " AND numserie <> '" & txtSerie(4).Text & "'"
    SQL = SQL & " group by codmacta "

    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe)  VALUES (" & vUsu.Codigo & ",0,'"
    
    While Not RS.EOF
        'Antes 20 Abril
        'Importe = RS!s1 + RS!s2 + RS!s3
        Importe = RS!s1
        If Not IsNull(RS!s2) Then Importe = Importe + RS!s2
        If Not IsNull(RS!s3) Then Importe = Importe + RS!s3
        Cad = RS!codmacta & "','" & RS!nifdatos & "'," & TransformaComasPuntos(CStr(Importe))
        Cad = SQL & Cad & ")"
        Conn.Execute Cad

        'sig
        RS.MoveNext
    Wend
    RS.Close
    
    
    'Proveedores
    SQL = DevuelveDesdeBD("max(codigo)", "usuarios.ztesoreriacomun ", "codusu", CStr(vUsu.Codigo))
    If SQL = "" Then SQL = "0"
    ContadorInserciones = CInt(SQL)
    
    Cad = "fecrecpr"
    
    'Agosto 2015
    'Cargamos las facturas para que puedan asignarle una clave a mano en un frm  ,
    SQL = "SELECT  " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,'" & Contabilidad & "',1,cabfactprov.codmacta,nifdatos,"
    SQL = SQL & " numfacpr,nommacta,"
    SQL = SQL & " ba1facpr + coalesce(ba2facpr,0) + coalesce(ba3facpr,0),fecrecpr,substring(coalesce(pais,''),1,2) "
    SQL = SQL & " from " & Contabilidad & ".cabfactprov," & Contabilidad & ".cuentas ,(SELECT @rownum:=" & ContadorInserciones & ") r "
    SQL = SQL & " WHERE cuentas.codmacta=cabfactprov.codmacta "
    SQL = SQL & " AND " & Cad & " >='" & Format(Text3(26).Text, FormatoFecha) & "'"
    SQL = SQL & " AND " & Cad & " <='" & Format(Text3(27).Text, FormatoFecha) & "'"
    SQL = SQL & " AND extranje = 1"
    SQL = "INSERT INTO usuarios.ztesoreriacomun(codusu,codigo,texto1,opcion,texto2,texto3,texto4,texto5,importe1,fecha1,texto6) " & SQL
    Conn.Execute SQL
    
    
    
    SQL = "SELECT cabfactprov.codmacta,nifdatos,sum(ba1facpr)as s1,sum(ba2facpr)as s2,sum(ba3facpr)as s3 from " & Contabilidad & ".cabfactprov," & Contabilidad & ".cuentas  where "
    SQL = SQL & " cuentas.codmacta=cabfactprov.codmacta "
    SQL = SQL & " AND " & Cad & " >='" & Format(Text3(26).Text, FormatoFecha) & "'"
    SQL = SQL & " AND " & Cad & " <='" & Format(Text3(27).Text, FormatoFecha) & "'"
    'Extranjero
    SQL = SQL & " AND extranje = 1"
    SQL = SQL & " group by codmacta "
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe)  VALUES (" & vUsu.Codigo & ",1,'"
    While Not RS.EOF
        'Antes 20 Abril
        'Importe = RS!s1 + RS!s2 + RS!s3
        Importe = RS!s1
        If Not IsNull(RS!s2) Then Importe = Importe + RS!s2
        If Not IsNull(RS!s3) Then Importe = Importe + RS!s3
        Cad = RS!codmacta & "','" & RS!nifdatos & "'," & TransformaComasPuntos(CStr(Importe))
        Cad = SQL & Cad & ")"
        Conn.Execute Cad
        
        'sig
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    
    
    
    
    RC = ""
    Cad = ""
    'Comprobaremos k el nif no es nulo, ni el codppos de las cuentas a tratar
    SQL = "Select cta from tmp347 where (nif is null or nif = '') and codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        Cad = Cad & RS.Fields(0) & "       "
        If I = 3 Then
            Cad = Cad & vbCrLf
            I = 0
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    If Cad <> "" Then
        RC = "Cuentas con NIF sin valor: " & vbCrLf & vbCrLf & Cad
    End If
    
    
    If RC <> "" Then
       MsgBox RC, vbExclamation
       Exit Function
    End If
    
   
    '----------------------------------------------------------
    'Listado detallado de las facturas en negativo
    '----------------------------------------------
    'CLIENTES
    
    'Para insertar
    RC = "INSERT INTO Usuarios.ztmpfaclin (codusu, codigo, Numfac, Fecha, cta,  NIF, "
    RC = RC & " IVA,  Total,cliente) VALUES (" & vUsu.Codigo & ","
    SQL = "SELECT  numserie,codfaccl,fecfaccl,totfaccl,nif,cabfact.codmacta,nommacta,ba1faccl,ba2faccl,ba3faccl  from " & Contabilidad & ".cabfact,"
    SQL = SQL & Contabilidad & ".cuentas ,"
    'La tmp es la de la empresa local
    'SQL = SQL & Contabilidad & ".tmp347  where "
    SQL = SQL & "tmp347  where "
    SQL = SQL & " tmp347.cta=cuentas.codmacta "
    SQL = SQL & " AND tmp347.cta= cabfact.codmacta"
    SQL = SQL & " AND fecfaccl >='" & Format(Text3(26).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecfaccl <='" & Format(Text3(27).Text, FormatoFecha) & "'"
    'Factura extranjero
    SQL = SQL & " AND intracom=1"
    'De compras / vetnas cojemos compras
    SQL = SQL & " AND cliprov = 0"
    
    
    'Importes negativos
    SQL = SQL & " AND totfaccl <0"
    
    'Pero si tiene serie de AUTOFACTURAS, la quitamos
    If txtSerie(4).Text <> "" Then SQL = SQL & " AND numserie <> '" & txtSerie(4).Text & "'"
    
    
    'Modificacion del 27 Febrero 2006
    SQL = SQL & " AND tmp347.codusu = " & vUsu.Codigo
    
    
    'Nº Empresa
    I = Val(Mid(Contabilidad, 6))

    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        ContadorCli = ContadorCli + 1
        'INSERT INTO ztmpfaclin (codusu, codigo, Numfac,fecha "
        
        SQL = ContadorCli & ",'" & RS!NUmSerie & Format(RS!codfaccl, "0000000000") & "','" & Format(RS!fecfaccl, FormatoFecha) & "','"
        ', cta,  NIF, IVA,  Total   .- IVA= numero empresa
        Importe = RS!ba1faccl + DBLet(RS!ba2faccl, "N") + DBLet(RS!ba3faccl, "N")
        SQL = SQL & RS!codmacta & "','" & RS!NIF & "'," & I & "," & TransformaComasPuntos(CStr(Importe))
        
        
        SQL = SQL & ",'" & DevNombreSQL(RS!nommacta)
        SQL = RC & SQL & "')"
    
        
        Conn.Execute SQL
    
        RS.MoveNext
    Wend
    RS.Close
    
 
    'PROVEEDORES
    
    RC = "INSERT INTO Usuarios.ztmpfaclinprov (codusu, codigo, Numfac, FechaCon, cta,  NIF, "
    RC = RC & " IVA,  Total,Fechafac,cliente) VALUES (" & vUsu.Codigo & ","
    SQL = "SELECT  numregis,fecrecpr,fecfacpr,totfacpr,numfacpr,nif,nommacta,ba1facpr,ba2facpr,ba3facpr  from " & Contabilidad & ".cabfactprov,"
    'La tmp es la de la empresa local
    SQL = SQL & Contabilidad & ".Cuentas,"
    SQL = SQL & "tmp347  where "
    SQL = SQL & " tmp347.cta=cabfactprov.codmacta AND tmp347.cta=Cuentas.codmacta "
    
    'Solo usuario 1
    SQL = SQL & " AND tmp347.codusu = " & vUsu.Codigo
    
    SQL = SQL & " AND fecrecpr >='" & Format(Text3(26).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecrecpr <='" & Format(Text3(27).Text, FormatoFecha) & "'"
    'Factura extranjero
    SQL = SQL & " AND extranje=1"
    
    'De compras / vetnas cojemos compras
    SQL = SQL & " AND cliprov = 1"
    
    'Importes negativos
    SQL = SQL & " AND totfacpr <0"

    
    'Modificacion del 27 Febrero 2006
    SQL = SQL & " AND tmp347.codusu = " & vUsu.Codigo
    
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        ContadorPro = ContadorPro + 1
        'INSERT INTO ztmpfaclin (codusu, codigo, Numfac,fecha "
        
        SQL = ContadorPro & ",'" & Format(RS!NumRegis, "0000000000") & "','" & Format(RS!fecrecpr, FormatoFecha) & "','"
        ', cta,  NIF, IVA,  Total   .- IVA= numero empresa    cta=cod factura
        
        'Abril 2006. Busco la base, no el total factura
        Importe = RS!ba1facpr + DBLet(RS!ba2facpr, "N") + DBLet(RS!ba3facpr, "N")
        SQL = SQL & Mid(RS!numfacpr, 1, 10) & "','" & RS!NIF & "'," & I & "," & TransformaComasPuntos(CStr(Importe))
        
        
        
        SQL = SQL & ",'" & Format(RS!fecfacpr, FormatoFecha) & "','" & DevNombreSQL(RS!nommacta)
        SQL = RC & SQL & "')"
    
        
        Conn.Execute SQL
    
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing
    ComprobarCuentas349_DOS = True
    Exit Function
EComprobarCuentas349:
    MuestraError Err.Number, "Comprobar Cuentas 349"
End Function



Private Sub CopiarFicheroHacienda(Modelo347 As Boolean)
    On Error GoTo ECopiarFichero347
    MsgBox "El archivo se ha generado con exito.", vbInformation
    SQL = ""
    cd1.CancelError = True
    cd1.ShowSave
    If Modelo347 Then
        SQL = App.path & "\Hacienda\mod347\mod347.txt"
    Else
        SQL = App.path & "\Hacienda\mod349\mod349.txt"
    End If
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El fichero ya existe. ¿Reemplazar?", vbQuestion + vbYesNo) = vbNo Then SQL = ""
        End If
        If SQL <> "" Then
            FileCopy SQL, cd1.FileName
            MsgBox Space(20) & "Copia efectuada correctamente" & Space(20), vbInformation
        End If
    End If
    Exit Sub
ECopiarFichero347:
    If Err.Number <> 32755 Then MuestraError Err.Number, "Copiar fichero 347"
    
End Sub

'Copiar Persa
Private Sub CopiarPersa()

    On Error GoTo ECopiarPersa
    Me.lblPersa.Caption = "Copiando "
    Me.lblPersa2.Caption = "Archivos"
    Me.Refresh
    RC = App.path & "\conta" & vEmpresa.codempre & "_maestro"
    Cad = "-" & RC & vbCrLf
    Cont = FileLen(RC)
    RC = App.path & "\conta" & vEmpresa.codempre & "_histor"
    Cad = Cad & "-" & RC & vbCrLf
    Cont = Cont + FileLen(RC)
    If Check4.Value = 0 Then
        Cont = 0
    Else
        Cont = Cont \ 1024
    End If
    If Cont > 1440 Then
        I = 1
        SQL = "Los archivos generados ocupan mas de la capacidad de un diskette." & vbCrLf & vbCrLf & Cad
        SQL = SQL & vbCrLf & vbCrLf & "Elija una carpeta donde copiarlos"
        MsgBox SQL, vbInformation
        SQL = ""
    Else
        'MsgBox "Archivos generados con exito. ", vbInformation
        If Check4.Value = 1 Then
            'A diskette
            SQL = "A:"
        Else
            'K pregunte carpeta
            SQL = ""
        End If
    End If
    
    If SQL <> "" Then
        Cad = SQL
    Else
        Cad = GetFolder("Carpeta destino")
    End If
    If Cad = "" Then Exit Sub
    
    'Copiamos archivos
    RC = App.path & "\conta" & vEmpresa.codempre & "_maestro"
    SQL = Cad & "\conta" & vEmpresa.codempre & "_maestro"
    FileCopy RC, SQL
    RC = App.path & "\conta" & vEmpresa.codempre & "_histor"
    SQL = Cad & "\conta" & vEmpresa.codempre & "_histor"
    FileCopy RC, SQL
    RC = Cad & "\conta" & vEmpresa.codempre & "_maestro" & vbCrLf
    RC = RC & Cad & "\conta" & vEmpresa.codempre & "_histor"
    MsgBox "Se han generado los ficheros: " & vbCrLf & RC, vbInformation
    
    Exit Sub
ECopiarPersa:
    MuestraError Err.Number, "Copiar fichero traspaso PERSA"
End Sub


Private Sub CopiarACE()

    On Error GoTo ECopiarAcE
    RC = App.path & "\trasace.txt"
    
    
    SQL = GetFolder("Destino traspaso ACE")
    'ANTES ERA A DISKETTE. AHORA OFERTO CARPETA
    '
   '
   ' Cont = FileLen(RC)
   ' Cont = Cont \ 1024
   ' If Cont > 1440 Then
   '     SQL = "El archivo generados ocupa más de la capacidad de un diskette." & vbCrLf & vbCrLf & Cad
   '     MsgBox SQL, vbInformation
   '     SQL = ""
   ' Else
   '     MsgBox "Archivo generado con éxito. " & vbCrLf & RC, vbInformation
   '     SQL = "A:\trasace.txt"
   ' End If
        
    If SQL = "" Then Exit Sub
    
    'Copiamos archivos
    FileCopy RC, SQL & "\trasace.txt"
    Exit Sub
ECopiarAcE:
    MuestraError Err.Number, "Copiar fichero traspaso ACE"
End Sub



Private Function DevuelveRegistrosProveedores() As Long
Dim C As String
    On Error GoTo EDevuelveRegistrosProveedores
    DevuelveRegistrosProveedores = 1
    If Opcion = 13 Then
        C = "Select count(*) from cabfactprov where fecrecpr< '" & Format(Text3(10).Text, FormatoFecha) & "'"
        C = C & " and fecrecpr>= '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Else
        C = "Select count(*) from cabfact where fecfaccl< '" & Format(Text3(10).Text, FormatoFecha) & "'"
        C = C & " and fecfaccl>= '" & Format(vParam.fechaini, FormatoFecha) & "'"
        C = C & " and numserie ='" & txtSerie(0).Text & "'"
    End If
    Set RS = New ADODB.Recordset
    RS.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        DevuelveRegistrosProveedores = DBLet(RS.Fields(0), "N") + 1
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EDevuelveRegistrosProveedores:
    MuestraError Err.Number, "Devuelve Registros Proveedores"
End Function



Private Function ObtenerFechasEjercicioContabilidad(Inicio As Boolean, Contabi As Integer) As Date

    On Error GoTo EObtenerFechasEjercicioContabilidad
    If Inicio Then
        ObtenerFechasEjercicioContabilidad = vParam.fechaini
    Else
        ObtenerFechasEjercicioContabilidad = vParam.fechafin
    End If
    
    RS.Open "Select fechaini,fechafin from Conta" & Contabi & ".parametros", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Inicio Then
            ObtenerFechasEjercicioContabilidad = RS.Fields(0)
        Else
            ObtenerFechasEjercicioContabilidad = RS.Fields(1)
        End If
        RS.Close
    Else
        RS.Close
        GoTo EObtenerFechasEjercicioContabilidad
    End If
    
    Exit Function
EObtenerFechasEjercicioContabilidad:
    MuestraError Err.Number, "Obtener Fechas Ejercicio Contabilidad: " & Contabi
End Function




'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'           Para la legalizacion de libros
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------

Private Sub GeneraLegalizaPRF(ByRef OtrosP As String, NumPara As Integer)
Dim NomArchivo As String

    'Estos informes los tengo k poner a mano
    'Si los cambiaramos habria k cambiarlos en imprime y aqui

    NomArchivo = App.path & "\InformesD\"
    Select Case Opcion
    Case 32
        NomArchivo = NomArchivo & "DiarioOf.rpt"
    Case 33
        NomArchivo = NomArchivo & "resumen.rpt"
    Case 34
        NomArchivo = NomArchivo & "ConsExtracL1.rpt"
    Case 35
        NomArchivo = NomArchivo & "AsientoHco.rpt"
    Case 36, 41
        NomArchivo = NomArchivo & "Sumas2.rpt"
    Case 37
        NomArchivo = NomArchivo & "faccli2.rpt"
    Case 38
        NomArchivo = NomArchivo & "facprov2.rpt"
    Case 39, 40
        If vParam.NuevoPlanContable Then
            'Nuevos balances
            If chkBalPerCompa.Value = 0 Then
                NomArchivo = NomArchivo & "balance1a.rpt"
            Else
                NomArchivo = NomArchivo & "balance2a.rpt"
            End If
        Else
            If chkBalPerCompa.Value = 0 Then
                NomArchivo = NomArchivo & "balance1.rpt"
            Else
                NomArchivo = NomArchivo & "balance2.rpt"
            End If

        End If
    End Select

   With frmVisReport
        If Opcion <> 34 Then
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
        Else
            .FormulaSeleccion = "{ado_lineas.codusu}=" & vUsu.Codigo
        End If
        .SoloImprimir = False
        .OtrosParametros = OtrosP
        .NumeroParametros = NumPara
        .MostrarTree = False
        .Informe = NomArchivo
        .ExportarPDF = True
        .Show vbModal
    End With
 
End Sub



Private Function CompararEmpresasBlancePerson(CONTA As Integer, ByRef E As Cempresa, FechaInicio As Date) As Boolean
Dim I As Integer
Dim J As Integer
    On Error GoTo ECompararEmpresasBlancePerson
    CompararEmpresasBlancePerson = False
    
    RC = "Empresa configurada"
    Set RS = New ADODB.Recordset
    
    SQL = "Select * from Conta" & CONTA & ".Empresa"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
       Cad = "Empresa no configurada"
       RS.Close
       Exit Function
    End If
    

    'OK empresa
    'Veamos los niveles
    RC = "Niveles"
    
    
    
    'Numero de niveles
    I = RS!numnivel
    If I <> vEmpresa.numnivel Then
        Cad = "numero de niveles distintos. " & vEmpresa.numnivel & " - " & I
        RS.Close
        Exit Function
    End If
    
    
    
    For J = 1 To vEmpresa.numnivel - 1
        I = DigitosNivel(J)
        NumRegElim = DBLet(RS.Fields(3 + J), "N")
        If I <> NumRegElim Then
            Cad = "Numero de digitos de nivel " & J & " son distintos. " & I & " - " & NumRegElim
            RS.Close
            Exit Function
        End If
    Next J
    
    RS.Close
    
    
    RC = "Parametros"
    Set RS = New ADODB.Recordset
    
    SQL = "Select * from Conta" & CONTA & ".Parametros"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
       Cad = "Parametros no configurados"
       RS.Close
       Exit Function
    End If
    
    
    If RS!fechaini <> FechaInicio Then
        Cad = "Fecha inicio ejercicios distintas: " & FechaInicio & " - " & RS!fechaini
        RS.Close
        Exit Function
    End If
    If FechaInicio <> vParam.fechaini Then
        Cad = "La fecha de inicio de ejercicio no coincide con los datos de la empresa actual. " & vParam.fechaini & " - " & FechaInicio
        RS.Close
        Exit Function
    End If
    RS.Close
    CompararEmpresasBlancePerson = True
    
ECompararEmpresasBlancePerson:
    If Err.Number <> 0 Then
        Cad = "Conta " & CONTA & vbCrLf & RC & vbCrLf & Err.Description
    Else
        Cad = ""
        CompararEmpresasBlancePerson = True
    End If
    Set RS = Nothing
End Function






'//////////////////////////////////////////////////////////////////
'       LISTADO FACTURAS PROVEEDORES
'  FechaOfactura -  0.- Nº Fac
'                   1.- Fecha emision
'                   2.- Fecha liqudiacion
'
'   NIF:   Quiere filtrar por NIF
Public Function ListadoFacturasProveedoresConsolidado2(vSQL As String, Ordenacion As String, NumeroFAC As Long, Agrupa As Boolean, FechaOfactura As Byte, Liquidacion As Boolean, Proveedores As Boolean, NIF As String) As Boolean
Dim V As Integer
    On Error GoTo EListadoFacturasProveedoresConsolidado
    ListadoFacturasProveedoresConsolidado2 = False
    
    Conn.Execute "DELETE FROM Usuarios.zFactConso where codusu =" & vUsu.Codigo
    Conn.Execute "Delete from Usuarios.ztmpresumenivafac where codusu =" & vUsu.Codigo
    
    
    If Proveedores Then
        
        For I = 1 To List9.ListCount
                InsertaFacturasEmpresa Proveedores, List9.ItemData(I - 1), vSQL, FechaOfactura, NIF
        Next I
        
    Else
    
        For I = 1 To List10.ListCount
                InsertaFacturasEmpresa Proveedores, List10.ItemData(I - 1), vSQL, FechaOfactura, NIF
        Next I
    
    End If
    
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select count(*) from Usuarios.zFactConso where codusu =" & vUsu.Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.EOF) Then
            If miRsAux.Fields(0) > 0 Then I = 1
        End If
    End If
    miRsAux.Close
    
    
    If I = 0 Then
        MsgBox "Ningún dato para mostrar", vbExclamation
        Set miRsAux = Nothing
        Exit Function
    End If
    
    
    'Ahora voy a cojer e insertar la sumas de los IVAS
    '
    Cont = 1
    For I = 1 To 3
    
        SQL = "select pi" & I & "faccl,sum(ba" & I & "faccl), sum(ti" & I & "faccl) from Usuarios.zfactconso where"
        SQL = SQL & " codusu =" & vUsu.Codigo & " group by pi" & I & "faccl    "
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If Not IsNull(miRsAux.Fields(0)) Then
                SQL = "INSERT INTO Usuarios.ztmpresumenivafac (codusu, orden, IVA, TotalIVA, sumabases) VALUES ("
                SQL = SQL & vUsu.Codigo & "," & Cont & ","
                SQL = SQL & TransformaComasPuntos(CStr(miRsAux.Fields(0))) & ","
                SQL = SQL & TransformaComasPuntos(CStr(miRsAux.Fields(1))) & ","
                SQL = SQL & TransformaComasPuntos(CStr(miRsAux.Fields(2)))
        
                SQL = SQL & ")"
                Conn.Execute SQL
                Cont = Cont + 1
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    Next I
    
    
    ListadoFacturasProveedoresConsolidado2 = True
    Set miRsAux = Nothing
    Exit Function
    
EListadoFacturasProveedoresConsolidado:
    MuestraError Err.Number
End Function



Private Function InsertaFacturasEmpresa(Provee As Boolean, NumEmpre As Integer, ByRef vSQL As String, FechaOfactura As Byte, NIF As String) As Boolean
Dim Cad As String
'       LISTADO FACTURAS PROVEEDORES
'  FechaOfactura -  0.- Nº Fac
'                   1.- Fecha emision
'                   2.- Fecha liqudiacion

    
    Cad = "INSERT INTO Usuarios.zfactconso (codusu, codempre,"
    Cad = Cad & "numserie,codfaccl, fecfaccl, codmacta, anofaccl, confaccl,"
    Cad = Cad & "ba1faccl, ba2faccl, ba3faccl, pi1faccl, pi2faccl, pi3faccl,"
    Cad = Cad & "pr1faccl, pr2faccl, pr3faccl, ti1faccl, ti2faccl, ti3faccl,"
    Cad = Cad & "tr1faccl, tr2faccl, tr3faccl, totfaccl, "
    Cad = Cad & "trefaccl, fecliqcl, nommacta)"
    'El select
    If Provee Then
        
    
        'PROVEEDORES
        Cad = Cad & "Select " & vUsu.Codigo & "," & NumEmpre & ",'0',"
        Cad = Cad & "numregis, "
        
        RC = "fecfacpr"
        If optListFacP(2).Value Then
            'Fecha rec o Liqu
            If Me.optSelFech(1).Value Then
                'LIQUIDA
                Cad = Cad & "fecliqpr"
            Else
                'FEc FAC
                Cad = Cad & "fecrecpr"
            End If
        Else
            'Fecha rec
            If optListFacP(0).Value Then
                Cad = Cad & "fecrecpr"
            Else
                Cad = Cad & "fecfacpr"
                RC = "fechaent"
            End If
        End If
        Cad = Cad & ", cuentas.codmacta, anofacpr,"
        'Numero factura o fecha entrada
        If optMostrarFecha(0).Value Then
            Cad = Cad & "numfacpr"
        Else
            Cad = Cad & RC
        End If
        Cad = Cad & ",ba1facpr, ba2facpr, ba3facpr, pi1facpr, pi2facpr, pi3facpr,"
        Cad = Cad & "pr1facpr, pr2facpr, pr3facpr, ti1facpr,ti2facpr, ti3facpr, "
        Cad = Cad & "tr1facpr, tr2facpr, tr3facpr, totfacpr,"
        Cad = Cad & "trefacpr, fecliqpr, nommacta from "
        Cad = Cad & "Conta" & NumEmpre & ".cabfactprov,Conta" & NumEmpre & ".cuentas where "
        Cad = Cad & "cabfactprov.codmacta = cuentas.codmacta "
        If vSQL <> "" Then Cad = Cad & "AND " & vSQL
        If NIF <> "" Then Cad = Cad & " AND nifdatos = '" & NIF & "'"
        
        
        
    Else
        '----------------------------------------------------
        '---------------------------------------------------
        ' CLIENTES
        '---------------------------------------------------
        
'            cad = cad & "numserie,codfaccl, fecfaccl, codmacta, anofaccl, confaccl,"
'            cad = cad & "ba1faccl, ba2faccl, ba3faccl, pi1faccl, pi2faccl, pi3faccl,"
'            cad = cad & "pr1faccl, pr2faccl, pr3faccl, ti1faccl, ti2faccl, ti3faccl,"
'            cad = cad & "tr1faccl, tr2faccl, tr3faccl, totfaccl, "
'            cad = cad & "trefaccl, fecliqcl, nommacta)"

        
        Cad = Cad & "Select " & vUsu.Codigo & "," & NumEmpre & ",numserie,"
        Cad = Cad & "codfaccl, "
        
        RC = "fecfaccl"
        If optListFac(3).Value Then RC = "fecliqcl"
        Cad = Cad & RC & ", cuentas.codmacta, anofaccl"

        'En el concepto voy a poner el NIF
        Cad = Cad & ",nifdatos"

        Cad = Cad & ",ba1faccl, ba2faccl, ba3faccl, pi1faccl, pi2faccl, pi3faccl,"
        Cad = Cad & "pr1faccl, pr2faccl, pr3faccl, ti1faccl,ti2faccl, ti3faccl, "
        Cad = Cad & "tr1faccl, tr2faccl, tr3faccl, totfaccl,"
        Cad = Cad & "trefaccl, fecliqcl, nommacta from "
        Cad = Cad & "Conta" & NumEmpre & ".cabfact,Conta" & NumEmpre & ".cuentas where "
        Cad = Cad & "cabfact.codmacta = cuentas.codmacta "
        If vSQL <> "" Then Cad = Cad & "AND " & vSQL
        If NIF <> "" Then Cad = Cad & " AND nifdatos = '" & NIF & "'"
        
        
    End If
    Conn.Execute Cad
End Function

Private Sub LeerSerieFacturas(Leer As Boolean)
On Error GoTo ELeerSerieFacturas
    SQL = App.path & "\ult349.xdf"
    
    If Leer Then
        If Dir(SQL) <> "" Then
            I = FreeFile
            Open SQL For Input As I
            Line Input #I, SQL
            Close #I
            txtSerie(4).Text = SQL
        End If
    Else
        If txtSerie(4).Text = "" Then
            If Dir(SQL) <> "" Then Kill SQL
        Else
            Open SQL For Output As I
            Print #I, txtSerie(4).Text
            Close #I
        End If
    End If
    Exit Sub
ELeerSerieFacturas:
    Err.Clear
End Sub

Private Sub ComprobarFechasBalanceQuitar6y7()
    On Error GoTo EComprobarFechasBalanceQuitar6y7
    Me.chkResetea6y7.visible = False
    If Not EjerciciosCerrados Then
        If CDate("01/" & cmbFecha(0).ListIndex + 1 & "/" & txtAno(0).Text) > vParam.fechafin Then Me.chkResetea6y7.visible = True
    End If
    Exit Sub
EComprobarFechasBalanceQuitar6y7:
    Err.Clear
End Sub




Public Sub ListadoKEYpress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 16 Then HacerF1
    End If
End Sub

Private Sub HacerF1()
    Select Case Opcion
    Case 1
        cmdListExtCta_Click
    Case 2
    
    Case 4
        cmdAceptarTotalesCta_Click
    Case 3
    
    Case 5
        cmdBalance_Click
    Case 6
        cmdAceptarHco_Click
    Case 7
        cmdCtaExplo_Click
    Case 8
        cmdFactCli_Click
    Case 9
        cmdPresupuestos_Click
    Case 10
        cmdBalPre_Click
    Case 13
        cmdFacProv_Click
        
    Case 14
        cmdLibroDiario_Click
    Case 15
        cmdSaldosCC_Click
    Case 16
        cmdCtaExpCC_Click
    Case 17
        cmdCCxCta_Click
    Case 18
        cmdDiarioRes_Click
    Case 19
        cmdCtapoCC_Click
        
    Case 21
        cmdComparativo_Click
        
    Case 54
        cmdEvolMensSald_Click
    Case Else
    
    End Select
End Sub


Private Function ListadoEvolucionMensual() As Boolean
Dim QuitarTambienElCierre As Boolean

    On Error GoTo EListadoEvolucionMensual
    ListadoEvolucionMensual = False


    'En cmbejercicios(0) tenemos las fechas
    '
    '   Con simple mid obtenemos inicio / fin
    
    
    SQL = cmbEjercicios(0).List(cmbEjercicios(0).ListIndex)
    RC = Mid(SQL, 1, 10)
    FechaIncioEjercicio = CDate(RC)
    RC = Mid(SQL, 14, 10)
    FechaFinEjercicio = CDate(RC)
    
    SQL = "Select hsaldos.codmacta,nommacta from hsaldos,cuentas where hsaldos.codmacta=cuentas.codmacta AND hsaldos.codmacta like '"
    'Nivel
    RC = ""
    For I = 1 To 10
        If ChkEvolSaldo(I).visible And (ChkEvolSaldo(I).Value = 1) Then
            If I = 10 Then
                Cont = vEmpresa.DigitosUltimoNivel
            Else
                'En el caption pone Digitos:  luego son 8 caractares
                Cad = Mid(ChkEvolSaldo(I).Caption, 9)
                If IsNumeric(Cad) Then
                    Cont = Val(Cad)
                Else
                    Cont = 0
                End If
            End If
            If Cont > 0 Then
                RC = Mid("__________", 1, Cont)
            End If
            Exit For
        End If
    Next I
    If RC = "" Then
        MsgBox "Error obteniendo nivel.", vbExclamation
        Exit Function
    End If
    
    SQL = SQL & RC & "'"
    'Si tienen desde /hasta
    If txtCta(23).Text <> "" Then SQL = SQL & " AND hsaldos.codmacta >= '" & txtCta(23).Text & "'"
    If txtCta(24).Text <> "" Then SQL = SQL & " AND hsaldos.codmacta <= '" & txtCta(24).Text & "'"

    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Año natural
        SQL = SQL & " AND anopsald = " & Year(FechaIncioEjercicio)
    Else
        'Años fiscales partidos . Coooperativas agricolas
        SQL = SQL & " AND ( (anopsald = " & Year(FechaIncioEjercicio) & " and mespsald >=" & Month(FechaIncioEjercicio) & ") OR "
        SQL = SQL & " (anopsald =" & Year(FechaIncioEjercicio) + 1 & " AND mespsald < " & Month(FechaIncioEjercicio) & "))"
    End If
    
    SQL = SQL & " GROUP BY codmacta"
    
    
    'stop
    QuitarTambienElCierre = FechaIncioEjercicio < vParam.fechaini
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        FijarValoresEvolucionMensualSaldos FechaIncioEjercicio, FechaFinEjercicio
        
        'PAra el SQL
        SQL = ""
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            'Año natural
            SQL = SQL & " AND anopsald = " & Year(FechaIncioEjercicio)
        Else
            'Años fiscales partidos . Coooperativas agricolas
            SQL = SQL & " AND ( (anopsald = " & Year(FechaIncioEjercicio) & " and mespsald >=" & Month(FechaIncioEjercicio) & ") OR "
            SQL = SQL & " (anopsald =" & Year(FechaIncioEjercicio) + 1 & " AND mespsald < " & Month(FechaIncioEjercicio) & "))"
        End If
        Cont = 0
        While Not RS.EOF
            Label2(29).Caption = RS!codmacta & " " & Mid(RS!nommacta, 1, 20) & " ..."
            Me.Refresh
            DatosEvolucionMensualSaldos2 RS!codmacta, RS!nommacta, SQL, chkEvolSalMeses.Value = 1, False, QuitarTambienElCierre
            RS.MoveNext
            If Cont > 150 Then
                Cont = 0
                DoEvents
                Screen.MousePointer = vbHourglass
            End If
            Cont = Cont + 1
        Wend
    End If
    RS.Close
    
    
    
    'Hacemos el conteo para ver si tiene o no movimientos
    Label2(29).Caption = "Comprobando valores"
    Label2(29).Refresh
    SQL = "Select count(*) from Usuarios.ztmpconextcab"
    SQL = SQL & " WHERE codusu =" & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not RS.EOF Then Cont = DBLet(RS.Fields(0), "N")
    RS.Close
    Set RS = Nothing
    If Cont > 0 Then
        ListadoEvolucionMensual = True
    Else
        MsgBox "No hay datos con estos valores", vbExclamation
    End If
    
    
    Label2(29).Caption = ""
    Exit Function
EListadoEvolucionMensual:
    MuestraError Err.Number
    Set RS = Nothing
    Label2(29).Caption = ""
End Function


Private Function GeneraRelacionCta_x_Bases() As Boolean
Dim F1 As Date
Dim F2 As Date


    Set RS = New ADODB.Recordset

    'Preparando tablas informe
    Label2(24).Caption = "Preparando tablas informe"
    Label2(24).Refresh
    DoEvents
    Cad = "DELETE from Usuarios.zentrefechas where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    
    Cont = 0
    F1 = CDate(Text3(36).Text)
    F2 = CDate(Text3(37).Text)
    GeneraRelacionCta_x_Bases2 True, F1, F2, ""


    
    If chkCliproxCtalineas.Value = 1 Then
        F1 = DateAdd("yyyy", -1, F1)
        F2 = DateAdd("yyyy", -1, F2)
        GeneraRelacionCta_x_Bases2 False, F1, F2, "ANTERIOR: "
    End If
        
        
        'Si no hay datos mostr
        Cad = "Select count(*) from Usuarios.zentrefechas where codusu =" & vUsu.Codigo
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cont = 0
        If Not RS.EOF Then Cont = DBLet(RS.Fields(0), "N")
        RS.Close
        If Cont = 0 Then
            MsgBox "No hay registros con estas opciones", vbExclamation
        Else
            GeneraRelacionCta_x_Bases = True
        End If
End Function


Private Function GeneraRelacionCta_x_Bases2(ProcesoNormal As Boolean, FI As Date, FF As Date, ParaElLabel As String) As Boolean
Dim J As Long

    On Error GoTo EGeneraRelacionCta_x_Bases2
    GeneraRelacionCta_x_Bases2 = False


    'Si no es procesonormal entonces es el comparativo, es decir, que para cada resultado tendra que comprobar que
    'existe la entrada y actualizar
    

    'He visto que cruzar las tablas cuando son muchos apuntes es duro(tiempo y carga de CPU)
    'Por lo tanto dividiremos el proceso en DOS
    '   1.- Cargar las facturas con la cta que esten en el desde /hasta
    '   2.- Cargar las lineas de facturas para esas cuentas
    Label2(24).Caption = ParaElLabel & "Obteniendo cabecera facturas"
    Label2(24).Refresh
    
    SQL = "DELETE FROM tmplinfactura WHERE codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    If False Then
        Cad = ""
    Else
        If Opcion = 55 Then
            SQL = "faccl"
        Else
            SQL = "facpr"
        End If
        Cad = ""
        For J = 1 To 3
            'Cad = Cad & ", tp" & Cont & SQL & " as tp" & Cont
            Cad = Cad & ", pi" & J & SQL & " as pi" & J
            'Cad = Cad & ", ti" & Cont & SQL & " as ti" & Cont
        Next J
    End If
    
    
    
    
    If Opcion = 55 Then
            'CLIENTES
        SQL = "select numserie,codfaccl as codigo,anofaccl as anyo,cabfact.codmacta,fecfaccl as fecha ,nommacta " & Cad & " from"
        SQL = SQL & " cabfact,cuentas WHERE cabfact.codmacta = cuentas.codmacta"
        RC = "fecfaccl"
        Tablas = "cabfact"
    Else
    
        'PROVEEDORES
        If optCta_x_gastos(0).Value Then
            RC = "fecfacpr"
        Else
            RC = "fecrecpr"
        End If
        Tablas = "cabfactprov"
        SQL = "select '' as numserie,numregis as codigo,anofacpr as anyo,cabfactprov.codmacta," & RC
        SQL = SQL & " as fecha ,nommacta " & Cad & " from"
        SQL = SQL & " cabfactprov,cuentas WHERE cabfactprov.codmacta = cuentas.codmacta"
        
    End If
    'Cta
    If txtCta(27).Text <> "" Then SQL = SQL & " AND " & Tablas & ".codmacta >= '" & txtCta(27).Text & "'"
    If txtCta(28).Text <> "" Then SQL = SQL & " AND " & Tablas & ".codmacta <= '" & txtCta(28).Text & "'"
    'Fecha
    If Text3(36).Text <> "" Then SQL = SQL & " AND " & RC & " >= '" & Format(FI, FormatoFecha) & "'"
    If Text3(37).Text <> "" Then SQL = SQL & " AND " & RC & " <= '" & Format(FF, FormatoFecha) & "'"
    
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    Cad = ""
    SQL = "INSERT INTO tmplinfactura (codusu, numserie, codfaccl, anofaccl,texto1,texto2,texto3) VALUES"
    While Not RS.EOF
        Label2(24).Caption = ParaElLabel & RS!Codigo & " (" & RS!NUmSerie & ")"
        Label2(24).Refresh
        
        J = J + 1
        Cad = Cad & "(" & vUsu.Codigo & ",'" & RS!NUmSerie & "'," & RS!Codigo & "," & RS!Anyo
        Cad = Cad & ",'" & RS!codmacta & "','" & DevNombreSQL(RS!nommacta) & "','" & Format(RS!Fecha, FormatoFecha) & "|" & AmpliacionRelacionCtaxCtaGastos(True) & "'),"
        
        If Len(Cad) > 100000 Then
            'Ejecutamos
            Cad = Mid(Cad, 1, Len(Cad) - 1)
            Cad = SQL & Cad
            Conn.Execute Cad
            Cad = ""
        End If
        DoEvents
        If PulsadoCancelar Then
            RS.Close
            Exit Function
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    
    If Cad <> "" Then
        'Ejecutamos
        Cad = Mid(Cad, 1, Len(Cad) - 1)
        Cad = SQL & Cad
        Conn.Execute Cad
        
    End If
    
    
    '----------------------------------------------------
    'Cruzamos con las bases
    '----------------------------------------------------
        'WHERE
        
    'Si detallamos el IVA o no


    If Opcion = 55 Then
        SQL = "select tmplinfactura.*,codtbase,impbascl as importe,linfact.impbascl as total, linfact.codtbase  from"
        SQL = SQL & " tmplinfactura,linfact where codusu = " & vUsu.Codigo
        SQL = SQL & " AND tmplinfactura.anofaccl=linfact.anofaccl and"
        SQL = SQL & " tmplinfactura.codfaccl=linfact.codfaccl AND tmplinfactura.numserie=linfact.numserie"
        
    Else
        SQL = "select tmplinfactura.*,codtbase,impbaspr as importe,linfactprov.impbaspr as total, linfactprov.codtbase from"
        SQL = SQL & " tmplinfactura,linfactprov where codusu = " & vUsu.Codigo
        SQL = SQL & " AND tmplinfactura.anofaccl=linfactprov.anofacpr and"
        SQL = SQL & " tmplinfactura.codfaccl=linfactprov.numregis "
        
    
    End If
    If txtCta(25).Text <> "" Then SQL = SQL & " AND codtbase >= '" & txtCta(25).Text & "'"
    If txtCta(26).Text <> "" Then SQL = SQL & " AND codtbase <= '" & txtCta(26).Text & "'"
    SQL = SQL & " ORDER BY codtbase"
    

    
    
    
    
    
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        RS.Close
        Exit Function
    End If
    Tablas = ""
    

    
    
    'SQL = "INSERT INTO Usuarios.zentrefechas (codusu, codigo, codccost, valoradq, fecventa, nomccost, nominmov, fechaadq,  impventa,nomconam)  VALUES "
    '                                                                                                                                           TIPOS IVA
    SQL = "INSERT INTO Usuarios.zentrefechas (codusu, codigo, codccost, valoradq, fecventa, nomccost, nominmov, fechaadq,  impventa,nomconam,conconam,amortacu,impperiodo)  VALUES "
    
    J = Cont   'El contador lo arrastro
    Cad = ""
    While Not RS.EOF
            Label2(24).Caption = ParaElLabel & Tablas & " - " & RS!codfaccl & RS!NUmSerie
            Label2(24).Refresh

            J = J + 1
            If Tablas <> RS!codtbase Then
                Tablas = RS!codtbase
                RC = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Tablas, "T")
                Label2(24).Caption = ParaElLabel & Tablas & " - " & RC
                Label2(24).Refresh
                NombreSQL RC
                DoEvents
            End If
            
        
            '(codusu, codigo, codccost, valoradq,nomccost,  nomconam,  nominmov, fechaadq,
            '         amortacu, fecventa, impventa)
            'Ejemplo
            If RS!Importe <> 0 Then
                
            
                Cad = Cad & "(" & vUsu.Codigo & "," & J & ",'" & RS!NUmSerie & "'," & RS!codfaccl & ",'" & RecuperaValor(RS!texto3, 1)
                Cad = Cad & "','" & RC & "','" & DevNombreSQL(RS!texto2) & "','" & Tablas & "',"
                If ProcesoNormal Then
                    Cad = Cad & TransformaComasPuntos(CStr(RS!Importe))
                Else
                    Cad = Cad & "0"
                End If
                
                Cad = Cad & ",'" & RS!texto1 & "',"
                'Los tipos de IVA
                Cad = Cad & AmpliacionRelacionCtaxCtaGastos(False, RS!texto3)
                
                'EN la funcion de arriba (ampliacionre.....) si comparativopone 2 valores. El terecero lo tendre que poner aqui
                If Me.chkCliproxCtalineas = 1 Then
                    'Si es en la primera pasada pongo el cero. En la segunda el importe (al reves que unas lineas mas arriba
                    If ProcesoNormal Then
                        Cad = Cad & "0"
                    Else
                        Cad = Cad & TransformaComasPuntos(CStr(RS!Importe))
                    End If
                End If
                Cad = Cad & "),"
            
                If Len(Cad) > 100000 Then
                    'Ejecutamos
                    Cad = Mid(Cad, 1, Len(Cad) - 1)
                    Cad = SQL & Cad
                    Conn.Execute Cad
                    Cad = ""
                End If
                DoEvents
                If PulsadoCancelar Then
                    RS.Close
                    Exit Function
                End If
                
            End If
            RS.MoveNext
            
            
        Wend
        
        RS.Close
        
        If Cad <> "" Then
            'Ejecutamos
            Cad = Mid(Cad, 1, Len(Cad) - 1)
            Cad = SQL & Cad
            Conn.Execute Cad
            Cad = ""
            DoEvents
        End If

    
    Cont = J
    
    GeneraRelacionCta_x_Bases2 = True
    Label2(24).Caption = ""
    Exit Function
EGeneraRelacionCta_x_Bases2:
    MuestraError Err.Number
    Label2(24).Caption = ""
End Function




Private Function AmpliacionRelacionCtaxCtaGastos(DesdeFactura As Boolean, Optional CADENA As String) As String

    'El primer tipop IVA es obligado
    If DesdeFactura Then
        AmpliacionRelacionCtaxCtaGastos = RS!pi1 * 100 & "|" 'El primero va en variable integer
        
        If Not IsNull(RS!pi2) Then
            AmpliacionRelacionCtaxCtaGastos = AmpliacionRelacionCtaxCtaGastos & TransformaComasPuntos(Format(RS!pi2, "0.00")) & "|"
        Else
            AmpliacionRelacionCtaxCtaGastos = AmpliacionRelacionCtaxCtaGastos & "NULL|"
        End If
        
        If Not IsNull(RS!pi3) Then
            AmpliacionRelacionCtaxCtaGastos = AmpliacionRelacionCtaxCtaGastos & TransformaComasPuntos(Format(RS!pi3, "0.00")) & "|"
        Else
            AmpliacionRelacionCtaxCtaGastos = AmpliacionRelacionCtaxCtaGastos & "NULL|"
        End If
        
    Else
        '-----------------------------------------------
        AmpliacionRelacionCtaxCtaGastos = RecuperaValor(CADENA, 2) & "," & RecuperaValor(CADENA, 3) & ","
        'Si no es compartivo entonces  pinto lo que haya en la ampliacion
        'Si es comparativo entonces vere si el valor es 0 o el importe calculado
        If Me.chkCliproxCtalineas.Value = 0 Then AmpliacionRelacionCtaxCtaGastos = AmpliacionRelacionCtaxCtaGastos & RecuperaValor(CADENA, 4)
        
    End If
    
End Function


Private Function Volcar347TablaTmp2() As Boolean
Dim Imp2 As Currency
Dim CuatroImportes(3) As Currency
On Error GoTo EVolcar
    Volcar347TablaTmp2 = False


    SQL = "DELETE from Usuarios.zsimulainm where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Select * from Usuarios.z347 where codusu = " & vUsu.Codigo & " ORDER BY nif"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = "Insert Into Usuarios.zsimulainm (codusu, codigo,  nomconam,  nominmov, fechaadq, valoradq, amortacu, totalamor) VALUES (" & vUsu.Codigo & ","
    Cad = ""
    Cont = 0
    While Not RS.EOF
        If RS!NIF <> Cad Then
            If Cad <> "" Then
                'Es otro NIF
                'Sera insert into
                Inserta347Agencias CuatroImportes(0), CuatroImportes(2), True
                Inserta347Agencias CuatroImportes(1), CuatroImportes(3), False
            End If
            Cad = RS!NIF
            RC = RS!razosoci
            CuatroImportes(0) = 0: CuatroImportes(1) = 0: CuatroImportes(2) = 0: CuatroImportes(3) = 0:
        End If
        'Sera UPDATE
        Select Case RS!cliprov
        Case 48
            I = 0
        Case 49
            I = 1
        Case 70
            I = 2
        Case 71
            I = 3
        End Select
        CuatroImportes(I) = RS!Importe
        
        
        
        RS.MoveNext
        
    Wend
    RS.Close
    'Metemos el ultimo registro
    Inserta347Agencias CuatroImportes(0), CuatroImportes(2), True
    Inserta347Agencias CuatroImportes(1), CuatroImportes(3), False
    Set RS = Nothing
    Volcar347TablaTmp2 = True
    Exit Function
EVolcar:
    MuestraError Err.Number
    Set RS = Nothing
End Function


Private Sub Inserta347Agencias(Importe1 As Currency, importe2 As Currency, Ventas As Boolean)
Dim C As String
    'SQL = "zsimulainm
    
    If Importe1 <> 0 Or importe2 <> 0 Then
        Cont = Cont + 1
        C = Cont & ",'" & Cad & "','" & DevNombreSQL(RC) & "','"
        If Ventas Then
            C = C & "VENTAS"
        Else
            C = C & "COMPRAS"
        End If
        C = C & "'," & TransformaComasPuntos(CStr(Importe1))
        C = C & "," & TransformaComasPuntos(CStr(importe2))
        C = C & "," & TransformaComasPuntos(CStr(Importe1 + importe2)) & ")"
        C = SQL & C
       Conn.Execute C
    End If
End Sub

'Cargo en el combo los ejercicios para que los seleccione
Private Sub CargaComboEjercicios(Indice As Integer)
Dim RS As Recordset
Dim PrimeraVez As Boolean
Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim Cad As String
        On Error GoTo ECargaComboEjericios
        
        Set RS = New ADODB.Recordset
        Cad = "Select min(fechaent) from hcabapu"  'FECHA MINIMA
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        FechaIncioEjercicio = vParam.fechaini
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then FechaIncioEjercicio = RS.Fields(0)
        End If
        RS.Close
        Set RS = Nothing
        
        'Cargo el combo
        '--------------------------------------------------------------------------------
        'Ajusto la primera fecha que devuelve a la que seria Inicio de ese ejercicio
        FechaFinEjercicio = CDate(Format(vParam.fechaini, "dd/mm/" & Year(FechaIncioEjercicio)))
        
        If FechaFinEjercicio > FechaIncioEjercicio Then
            'El ejercicio empieza un año antes
            FechaIncioEjercicio = DateAdd("yyyy", -1, FechaFinEjercicio)
        Else
            FechaIncioEjercicio = FechaFinEjercicio
        End If
        
        
        
        FechaFinEjercicio = DateAdd("yyyy", 1, vParam.fechafin)  'Final de año siguiente
        cmbEjercicios(Indice).Clear
        Cont = 0
        While FechaIncioEjercicio <= FechaFinEjercicio
                Cad = Format(FechaIncioEjercicio, "dd/mm/yyyy")
                FechaIncioEjercicio = DateAdd("yyyy", 1, FechaIncioEjercicio)
                FechaIncioEjercicio = DateAdd("d", -1, FechaIncioEjercicio)
                Cad = Cad & " - " & Format(FechaIncioEjercicio, "dd/mm/yyyy")
                'Le pongo una marca de actual o ssiguiente
                I = 0 'pAra memorizar cual es el que apunta
                If FechaIncioEjercicio > vParam.fechaini Then
                    If FechaIncioEjercicio = vParam.fechafin Then
                        Cad = Cad & "     Actual"
                        I = 1
                    Else
                        Cad = Cad & "     Siguiente"
                    End If
                End If
                'Meto en el combo
                cmbEjercicios(Indice).AddItem Cad
                If I = 1 Then Cont = cmbEjercicios(Indice).NewIndex
                'Paso a inicio del ejercicio siguiente sumandole un dia
                'al fin del anterior
                FechaIncioEjercicio = DateAdd("d", 1, FechaIncioEjercicio)
        Wend
             
        'En cont tengo actual
        Me.cmbEjercicios(Indice).ListIndex = Cont
        
        Exit Sub
ECargaComboEjericios:
    MuestraError Err.Number, "CargaComboEjericios"
    
End Sub



Private Sub DatosConsultaExtractoExtendida()

    On Error GoTo ED
    
    Label12.Caption = "Elimina datos aux"
    Label12.Refresh
    SQL = "DELETE FROM usuarios.zexplocomp where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "DELETE FROM usuarios.zsaldoscc where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    Label12.Caption = "Aux.  Contrapartida"
    Label12.Refresh
    SQL = "INSERT INTO usuarios.zexplocomp (codusu, cta, cuenta)"
    SQL = SQL & " Select " & vUsu.Codigo & ",contra,nommacta from tmpconext left join cuentas "
    SQL = SQL & " on tmpconext.codusu=" & vUsu.Codigo & " and tmpconext.contra=cuentas.codmacta "
    SQL = SQL & " where not (contra  is null) GROUP BY 1,2"
    Conn.Execute SQL
    
    
    '
    Label12.Caption = "Aux. centro coste"
    Label12.Refresh
    SQL = "INSERT INTO usuarios.zsaldoscc (codusu, ano, mes,codccost, nomccost) "
    SQL = SQL & " Select " & vUsu.Codigo & ",0,0,ccost,nomccost from tmpconext left join cabccost"
    SQL = SQL & " on tmpconext.codusu=" & vUsu.Codigo & " and tmpconext.ccost=cabccost.codccost"
    SQL = SQL & " where not (ccost  is null) GROUP BY 1,4"   'Agrupa por codusu, ccost
    Conn.Execute SQL
    
    Exit Sub
ED:
    MuestraError Err.Number, "Datos consulta extendido"
    Cont = 0
End Sub





Private Sub PonerFechaBorre()
Dim F As Date
    'Seleccionaremos la min fra y ese sera el ejercicio a borrar
    F = DateAdd("yyyy", -5, vParam.fechafin)
    Text3(24).Text = Format(F, "dd/mm/yyyy")
    
    If Opcion = 22 Then
        SQL = "Select min(fecfaccl) FROM cabfact"
    Else
        SQL = "Select min(fecrecpr) FROM cabfactprov"
    End If
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Label3(75).Tag = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If DateDiff("yyyy", F, RS.Fields(0)) <= 0 Then
                'OK. Podemos borrar algun ejercicio
                FechaFinEjercicio = DateAdd("yyyy", -5, vParam.fechafin)
                F = DateAdd("yyyy", -5, vParam.fechaini)
                
                Do
                    If F <= RS.Fields(0) Then
                        SQL = ""
                    Else
                        F = DateAdd("yyyy", -1, F)
                        FechaFinEjercicio = DateAdd("yyyy", -1, FechaFinEjercicio)
                    End If
                Loop Until SQL = ""
                Label3(75).Tag = Format(FechaFinEjercicio, "dd/mm/yyyy")
                Label3(75).Caption = "Fin ejercicio: " & Label3(75).Tag
            Else
                Label3(75).Caption = "NO hay datos para borrar"
            End If
        End If
     End If

    RS.Close
            
End Sub




Private Sub HacerBorreFacturasEjercicio()

    On Error GoTo EV
    
    'Tablas de trabajo
    Tablas = ""
    If Opcion = 23 Then Tablas = "prov"
    
    
    SQL = " WHERE anofac"
    If Opcion = 22 Then
        SQL = SQL & "cl = "
    Else
        SQL = SQL & "pr = "
    End If
    SQL = SQL & Year(CDate(Label3(75).Tag))
    
    pb8.Value = 0
    pb8.visible = True
    Me.Refresh
    
    'Cabceras
    RC = "INSERT INTO cabfact" & Tablas & "1 SELECT * from cabfact" & Tablas & SQL
    Conn.Execute RC
    pb8.Value = 250
    'lineas
    RC = "INSERT INTO linfact" & Tablas & "1 SELECT * from linfact" & Tablas & SQL
    Conn.Execute RC
    pb8.Value = 500
    
    'ELiminar
    RC = "DELETE from linfact" & Tablas & SQL
    Conn.Execute RC
    pb8.Value = 750
    
    'lineas
    RC = "DELETE from cabfact" & Tablas & SQL
    Conn.Execute RC
    pb8.Value = 1000
    
    Exit Sub
EV:
    MuestraError Err.Number, Err.Description
End Sub



Private Sub HacerBalanceInicio()
    
        'Numero de niveles
        'Para cada nivel marcado veremos si tiene cuentas en la tmp
        RC = ""
        For I = 1 To 10
            SQL = "0"
            If Check2(I).visible Then
                If Check2(I).Value = 1 Then SQL = "1"
            End If
            RC = RC & SQL
        Next I
    
        
    
    
                
        'Borramos los temporales
        SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute SQL
    
        'Precargamos el cierre
        PrecargaApertura  'Carga en ur RS la apertura
    
        Cont = 1
        If Not CargaBalanceInicioEjercicio(RC) Then Cont = 0
        CerrarPrecargaPerdidasyGanancias
        If Cont = 0 Then Exit Sub
                
        SQL = "Titulo= ""Balance inicio ejercicio""|"
        SQL = SQL & "NumPag= 0|"
        
        
        '------------------------------
        'Numero de niveles
        'Para cada nivel marcado veremos si tiene cuentas en la tmp
        Cont = 0
        For I = 1 To 10
            If Check2(I).visible Then
'                If Check2(I).Value = 1 Then Cont = Cont + 1
                If Check2(I).Value = 1 Then
                    If I = 10 Then
                        Cad = vEmpresa.DigitosUltimoNivel
                    Else
                        Cad = CStr(DigitosNivel(I))
                    End If
                    If TieneCuentasEnTmpBalance(Cad) Then Cont = Cont + 1
                End If
            End If
        Next I
        Cad = "numeroniveles= " & Cont & "|"
        SQL = SQL & Cad
        
        
        'Fecha de impresion
        SQL = SQL & "FechaImp= """ & Text3(7).Text & """|"
        
        
        'Remarcar
        If Combo3.ListIndex >= 0 Then
            SQL = SQL & "Salto= " & Combo3.ItemData(Combo3.ListIndex) & "|"
            Else
            SQL = SQL & "Salto= 11|"
        End If

        
        With frmImprimir
                .OtrosParametros = SQL
                .NumeroParametros = 7
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                'Opcion dependera del combo
                .Opcion = 18
                .Show vbModal
        End With
    
    
End Sub


Private Sub CargarComboPeriodo()
    Combo6.Clear
    If vParam.Presentacion349Mensual Then
        
        For I = 1 To 12
            SQL = Format(CDate("01/" & I & "/2000"), "mmmm")
            RC = UCase(Left(SQL, 1))
            SQL = RC & Mid(SQL, 2)
            Combo6.AddItem SQL
        Next I
        I = Month(Now) - 1
        If I > 1 Then I = I - 1
    Else
        'trimesres
        For I = 1 To 4
            SQL = "Trimestre " & I
            Combo6.AddItem SQL
        Next I
        
        I = Month(Now) \ 3
    End If
    
    'Añado el ultimo ITEM.  Anual
    'M.Carmen dice que este NO
    'Combo6.AddItem "ANUAL"
    RC = ""
    'Combo6.ListIndex = I
    
End Sub





