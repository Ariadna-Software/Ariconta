VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCentroControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "frmCentroControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameMovCtas 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CheckBox chkActualizarTesoreria 
         Caption         =   "Actualizar cobros/pagos"
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   2
         Left            =   2880
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdMovercta 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtcta 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txtcta 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   1
         Left            =   4680
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5760
         TabIndex        =   13
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   0
         Left            =   6720
         Picture         =   "frmCentroControl.frx":000C
         ToolTipText     =   "Ayuda"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   2520
         Picture         =   "frmCentroControl.frx":0A0E
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Bloquear cuenta de ORIGEN"
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
         Index           =   3
         Left            =   360
         TabIndex        =   24
         Top             =   3000
         Width           =   2025
      End
      Begin VB.Label Label16 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Mover cuentas "
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
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ORIGEN"
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
         Index           =   8
         Left            =   360
         TabIndex        =   21
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DESTINO"
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
         Index           =   9
         Left            =   360
         TabIndex        =   20
         Top             =   2520
         Width           =   660
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
         Index           =   2
         Left            =   3720
         TabIndex        =   19
         Top             =   1200
         Width           =   420
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
         Index           =   4
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   1680
         Width           =   690
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmCentroControl.frx":0A99
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmCentroControl.frx":149B
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4320
         Picture         =   "frmCentroControl.frx":1E9D
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   795
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmCentroControl.frx":1F28
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrameRenumFRAPRO 
      Height          =   4815
      Left            =   840
      TabIndex        =   76
      Top             =   0
      Width           =   6375
      Begin VB.Frame FrameTapaRenum 
         Height          =   1935
         Left            =   240
         TabIndex        =   93
         Top             =   1200
         Width           =   5895
         Begin VB.TextBox txtInformacion 
            BackColor       =   &H80000018&
            Height          =   375
            Index           =   0
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   94
            Text            =   "Text1"
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Ultimo periodo liquidado"
            Height          =   255
            Left            =   1200
            TabIndex        =   95
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.OptionButton optRenumPeriodo 
         Caption         =   "Desde ult periodo liquidado"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   92
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optRenumPeriodo 
         Caption         =   "Reumerar todo"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   91
         Top             =   960
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton cmdRenumFra 
         Caption         =   "Renumerar"
         Height          =   375
         Left            =   3840
         TabIndex        =   83
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtRenumFrapro 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   79
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtRenumFrapro 
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   80
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtRenumFrapro 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   81
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkCompruebaContab 
         Caption         =   "Comprobar contabilizadas"
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CheckBox chkUpdateNumDocum 
         Caption         =   "Updatear el campo numcoum"
         Height          =   255
         Left            =   3000
         TabIndex        =   78
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CheckBox chkSALTO_numerofactura 
         Caption         =   "El numero es un SALTO factura"
         Height          =   255
         Left            =   3000
         TabIndex        =   82
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5040
         TabIndex        =   85
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   1320
         TabIndex        =   96
         Top             =   1440
         Width           =   3015
         Begin VB.OptionButton optFrapro 
            Caption         =   "Siguiente"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   98
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optFrapro 
            Caption         =   "Actual"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   97
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   6000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label LabelIndF 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   90
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label LabelIndF 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   89
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5880
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label11 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   88
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   87
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Numero:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   86
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Renumerar número registro proveedor"
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
         Index           =   3
         Left            =   240
         TabIndex        =   84
         Top             =   240
         Width           =   5655
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   3
         Left            =   6000
         Picture         =   "frmCentroControl.frx":1FB3
         ToolTipText     =   "Ayuda"
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame FrameCambioIVA 
      Height          =   4575
      Left            =   1800
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   5
         Left            =   3480
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   4
         Left            =   1200
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   3060
         Width           =   1215
      End
      Begin VB.CommandButton cmdIVA 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2520
         TabIndex        =   66
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   4
         Left            =   3720
         TabIndex        =   67
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox txtDescIVA 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtIVA 
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtDescIVA 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtIVA 
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   3240
         Picture         =   "frmCentroControl.frx":29B5
         Top             =   3090
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
         Index           =   18
         Left            =   2760
         TabIndex        =   75
         Top             =   3120
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   840
         Picture         =   "frmCentroControl.frx":2A40
         Top             =   3097
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   17
         Left            =   240
         TabIndex        =   74
         Top             =   2760
         Width           =   795
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
         Index           =   16
         Left            =   240
         TabIndex        =   73
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label lblIVA 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Image imgIVA 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmCentroControl.frx":2ACB
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgIVA 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmCentroControl.frx":34CD
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   15
         Left            =   240
         TabIndex        =   71
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   14
         Left            =   240
         TabIndex        =   70
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cambiar IVA "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Index           =   2
         Left            =   1320
         TabIndex        =   61
         Top             =   240
         Width           =   2355
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   2
         Left            =   4680
         Picture         =   "frmCentroControl.frx":3ECF
         ToolTipText     =   "Ayuda"
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Frame FrDesbloq 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton cmdDesbloq 
         Caption         =   "Desbloquear"
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   3720
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Diario"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Nº Asiento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Obser"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6360
         TabIndex        =   1
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   6840
         Picture         =   "frmCentroControl.frx":48D1
         ToolTipText     =   "Quitar al haber"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   7200
         Picture         =   "frmCentroControl.frx":4A1B
         ToolTipText     =   "Puntear al haber"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desbloquear asientos"
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame FrameCeros 
      Height          =   3975
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   5415
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   2760
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   2760
         TabIndex        =   57
         Text            =   "Text2"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   2760
         TabIndex        =   56
         Text            =   "Text2"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2760
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdCeros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   50
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   49
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   2520
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "Digitos nivel anterior"
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   54
         Top             =   1960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Digitos último nivel"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   53
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Nº Niveles"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   52
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Aumentar digitos PGC  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Index           =   1
         Left            =   360
         TabIndex        =   51
         Top             =   240
         Width           =   4140
      End
      Begin VB.Image ImgAyuda 
         Height          =   240
         Index           =   1
         Left            =   5040
         Picture         =   "frmCentroControl.frx":4B65
         ToolTipText     =   "Ayuda"
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Frame frameNuevaEmpresa 
      Height          =   5715
      Left            =   240
      TabIndex        =   25
      Top             =   -120
      Width           =   6435
      Begin VB.CheckBox Check1 
         Caption         =   "Formas de pago"
         Height          =   195
         Index           =   7
         Left            =   2820
         TabIndex        =   37
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   39
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   660
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   1500
         Width           =   555
      End
      Begin VB.CommandButton cmdNuevaEmpresa 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   38
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia plan contable"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   3120
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia conceptos"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   3480
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia diarios"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia Tipos IVA"
         Height          =   195
         Index           =   3
         Left            =   2820
         TabIndex        =   35
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Asientos predefinidos"
         Height          =   195
         Index           =   4
         Left            =   2820
         TabIndex        =   31
         Top             =   3120
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia centros de coste"
         Height          =   195
         Index           =   5
         Left            =   2820
         TabIndex        =   33
         Top             =   3480
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia configuracion balances"
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   36
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmCentroControl.frx":5567
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre empresa"
         Height          =   255
         Index           =   11
         Left            =   300
         TabIndex        =   47
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre corto"
         Height          =   255
         Index           =   10
         Left            =   300
         TabIndex        =   46
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Número empresa"
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   45
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Creación nueva empresa ***"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Index           =   4
         Left            =   180
         TabIndex        =   44
         Top             =   180
         Width           =   5160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Insertar datos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   5
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha inicio"
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   42
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   41
         Top             =   2760
         Width           =   5715
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   4680
         Width           =   2835
      End
   End
End
Attribute VB_Name = "frmCentroControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
    '0.-  Desbloquear asientos
    '1.- Mover ctas
    '2.- Crear empresa nueva
    '3.- Aumento ce deros
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmI As frmIVA
Attribute frmI.VB_VarHelpID = -1

Dim I As Integer
Dim SQL As String
Dim PrimeraVez As Boolean


Dim TablaAnt As String
Dim Tam2 As Long
Dim Tamanyo As Long
Dim NumTablas As Integer
Dim ParaElLog As String
Dim Insert As String
Dim Campos()


Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkActualizarTesoreria_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCompruebaContab_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkSALTO_numerofactura_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkUpdateNumDocum_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCeros_Click()
Dim B As Boolean

    B = False

    If UsuariosConectados("Desbloqueando asientos" & vbCrLf, True) Then Exit Sub
    
    SQL = "Este programa aumentara el numero de digitos a ultimo nivel" & vbCrLf
    SQL = SQL & vbCrLf & vbCrLf & "Deberia hacer una copia de seguridad." & vbCrLf & vbCrLf
    SQL = SQL & "             ¿ Desea continuar?    "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    
    
    
    
        
    SQL = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(SQL) <> "ARIADNA" Then
        If SQL <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    If ComprobarOk(CByte(Text2(5).Text)) Then
        Label3.Caption = ""
        Label3.Visible = True
        pb1.Value = 0
        Me.pb1.Max = 1000
        Me.pb1.Visible = True
        B = HacerInsercionDigitoContable
        pb1.Visible = False
        Label3.Visible = False
        If B Then
            frmActualizar.OpcionActualizar = 15 'Recalculo automatico
            frmActualizar.Show vbModal
            
            
        End If
        
        'Insertamos el LOG
        ParaElLog = "Nº nivel: " & Text2(3).Text & vbCrLf
        ParaElLog = ParaElLog & "Digitos último nivel: " & Text2(4).Text & vbCrLf
        ParaElLog = ParaElLog & "Digitos nivel anterior: " & Text2(5).Text & vbCrLf
        ParaElLog = "Aumentar CERO (" & CStr(B) & ")" & vbCrLf & ParaElLog
        vLog.Insertar 16, vUsu, ParaElLog
        ParaElLog = ""
        
    End If
    Screen.MousePointer = vbDefault
    
    If B Then Unload Me
    
End Sub

Private Sub cmdDesbloq_Click()
    SQL = "Seleccione algún asiento para desbloquear"
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            SQL = ""
            Exit For
        End If
    Next I
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    If UsuariosConectados("Desbloqueando asientos" & vbCrLf, True) Then Exit Sub
        
    SQL = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(SQL) <> "ARIADNA" Then
        If SQL <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    ParaElLog = ""
    For I = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(I).Checked Then
            SQL = "UPDATE cabapu SET bloqactu = 0 WHERE numdiari =" & ListView1.ListItems(I).Text
            SQL = SQL & " AND fechaent = '" & Format(ListView1.ListItems(I).SubItems(1), FormatoFecha) & "'"
            SQL = SQL & " AND numasien = " & Val(ListView1.ListItems(I).SubItems(2))
            
            EjecutaSQL SQL
            'Para el LOG
            SQL = ListView1.ListItems(I).Text & "," & ListView1.ListItems(I).SubItems(1) & "," & Val(ListView1.ListItems(I).SubItems(2))
            
            ParaElLog = ParaElLog & ", [" & SQL & "]"
            ListView1.ListItems.Remove I
            
        End If
    Next I
    'Insertamos el LOG
    ParaElLog = "DESBLOQUEAR" & vbCrLf & ParaElLog
    vLog.Insertar 16, vUsu, ParaElLog
    ParaElLog = ""
    
    'Si nO queda ninguno cierro ventana
    If ListView1.ListItems.Count = 0 Then Unload Me
    
End Sub

Private Sub cmdIVA_Click()
Dim B As Boolean

        If txtIVA(0).Text = "" Or txtIVA(1).Text = "" Then
            MsgBox "IVA origen y destino requeridos", vbExclamation
            Exit Sub
        End If

        If txtIVA(0).Text = txtIVA(1).Text Then
            MsgBox "IVA origen no puede ser igual al IVA destino", vbExclamation
            Exit Sub
        End If
        SQL = "Deberia tener una copia de seguridad." & vbCrLf & "El proceso puede tardar mucho tiempo" & vbCrLf
        SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub


        If UsuariosConectados("Cambiando IVA" & vbCrLf, True) Then Exit Sub


        SQL = InputBox("Password de seguridad")
        If UCase(SQL) <> "ARIADNA" Then Exit Sub
    
        Screen.MousePointer = vbHourglass
        B = HacerCambioIVA
        lblIVA.Caption = ""
        Screen.MousePointer = vbDefault
        If B Then
            SQL = "Proceso finalizado con éxito." & vbCrLf & vbCrLf & "¿Desea realizar otro cambio?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then
                Unload Me
            Else
                Limpiar Me
                PonleFoco txtIVA(0)
            End If
        End If
End Sub

Private Sub cmdMovercta_Click()

         'Hacemos lo que tengamos que hacer
        If txtCta(0).Text = "" Or txtCta(1).Text = "" Then
            MsgBox "Ponga cuentas contables", vbExclamation
            Exit Sub
        End If
        If txtCta(0).Text = txtCta(1).Text Then
            MsgBox "Misma cuenta origen destino", vbExclamation
            Exit Sub
        End If
        
        If txtFecha(0).Text = "" Then
            MsgBox "Ponga la fecha ""Desde""", vbExclamation
            Exit Sub
        End If
        
        
        'Diciemnre 2012
        'Pequeñas comprobaciones
        'Si tiene pagos cobros Preguntara
        If vEmpresa.TieneTesoreria Then
            If Me.chkActualizarTesoreria.Value = 1 Then
                
                    I = 0
                    SQL = DevuelveDesdeBD("count(*)", "scobro", "codmacta", txtCta(0).Text, "T")
                    If Val(SQL) > 0 Then Insert = "cobros"
                    SQL = DevuelveDesdeBD("count(*)", "spagop", "ctaprove", txtCta(0).Text, "T")
                    If Val(SQL) > 0 Then Insert = Insert & " pagos"
                    If Insert <> "" Then
                        SQL = "Existen " & Insert & " relacionados con la cuenta. Continuar?"
                        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                            
                    End If
      
            End If
        End If
        Insert = ""
        
        SQL = "Deberia tener una copia de seguridad." & vbCrLf & "El proceso puede tardar mucho tiempo" & vbCrLf
        SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub

        SQL = InputBox("Password de seguridad")
        If UCase(SQL) <> "ARIADNA" Then Exit Sub
    
        If HacerCambioCuenta Then Unload Me
         Label16.Caption = ""
End Sub

Private Sub cmdNuevaEmpresa_Click(Index As Integer)

Dim Ok As Boolean
Dim T As TextBox


    For Each T In Text2
        If T.Visible Then
            T = Trim(T)
            If T = "" Then
                
                MsgBox "Todos los campos obligatorios", vbExclamation
                Exit Sub
            End If
        End If
    Next



    If Not IsNumeric(Text2(2).Text) Then
        MsgBox "Número de empresa tiene que ser numérico, obviamente", vbExclamation
        Exit Sub
    End If
    
    If Not IsDate(txtFecha(3).Text) Then
        MsgBox "Fecha inicio incorrecta", vbExclamation
        Exit Sub
    End If
    

    
    
    
    
    'Si marca el asipre tiene k tener marcados cuetas, y tal y tal
     Tam2 = Check1(0).Value + Check1(1).Value + Check1(2).Value + Check1(5).Value
     If Check1(4).Value = 1 Then
        If Tam2 <> 4 Then
            MsgBox "Si marca asientos predefinidos tiene que marcar cuentas, diarios, conceptos y centros de coste.", vbExclamation
            Exit Sub
        End If
    End If
    
    'Si marca IVA tiene que llevarse el plan contable, ya que los tipos de IVA estan
    'asociados a cuentas contables
    If Check1(3).Value Then
        If Check1(0).Value = 0 Then
            MsgBox "Los tipos de IVA estan asociados a cuentas contables de ultimo nivel.", vbExclamation
            Exit Sub
        End If
    End If
    
    SQL = "Va a generar una nueva empresa: " & Text2(0).Text
    SQL = SQL & vbCrLf & "Desea continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Ok = False
    Label6.Caption = "Generando estructura de BD"
    Me.Refresh
    
    Ok = GeneracionNuevaBD
    
    Label6.Caption = ""
    Screen.MousePointer = vbDefault
    If Ok Then
        MsgBox "Proceso finalizado con exito", vbExclamation
        Unload Me
    End If

End Sub



Private Sub cmdRenumFra_Click()
Dim Ok As Boolean


    If Me.optRenumPeriodo(0).Value Then
        'Renumeracion antigua
        
        If txtRenumFrapro(0).Text = "" Then
           MsgBox "Falta numero factura", vbExclamation
            Exit Sub
        End If
                    
        For I = 0 To 2
            If txtRenumFrapro(I).Text <> "" Then
                If Not IsNumeric(txtRenumFrapro(I).Text) Then
                    MsgBox "Campo numerico incorrecto", vbExclamation
                    Exit Sub
                End If
             End If
        Next I
                    
            
    Else
    
        'Varias comprobaciones.
        If Not SePuedeRenumerarPorPeriodo Then Exit Sub
    
    End If
    Me.LabelIndF(0).Caption = ""
    Me.LabelIndF(1).Caption = ""
    
    If MsgBox("Deberia hacer una copia de seguridad." & vbCrLf & vbCrLf & vbCrLf & "El proceso puede durar muchisimo tiempo. ¿Desea continuar igualmente?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        
    If UsuariosConectados("Renumerar nºReg. en factura proveedor" & vbCrLf, True) Then Exit Sub


    SQL = InputBox("Password de seguridad")
    If UCase(SQL) <> "ARIADNA" Then Exit Sub
    
        
        
    'OK---------------------------------
    'A renumerar
    Screen.MousePointer = vbHourglass
    cmdRenumFra.Enabled = False
    
    If Me.optRenumPeriodo(0).Value Then
        Ok = HacerRenumeracionFacturas
    Else
        Ok = RenumerarDesdeUltimoPeriodoLiquidacion
    End If
    
    If Ok Then

        'Insertamos el LOG
        If Me.optRenumPeriodo(0).Value Then
            'Antugio
            ParaElLog = "actual"
            If Me.optFrapro(1).Value Then ParaElLog = "siguiente"
              
            ParaElLog = "Ejercicio " & ParaElLog & vbCrLf
            
            ParaElLog = ParaElLog & "Nº registro " & txtRenumFrapro(0).Text & vbCrLf
            
        Else
            ParaElLog = "Ultimo periodo liquidacion: " & txtInformacion(0).Text
        End If
        ParaElLog = "Renumerar facturas proveedor." & vbCrLf & ParaElLog
        
        vLog.Insertar 16, vUsu, ParaElLog
        
        
        ParaElLog = String(40, "*") & vbCrLf
        ParaElLog = ParaElLog & ParaElLog & ParaElLog
        ParaElLog = ParaElLog & vbCrLf & vbCrLf & "Compruebe el contador de facturas de proveedor" & vbCrLf & vbCrLf & vbCrLf & ParaElLog
        MsgBox ParaElLog, vbExclamation
        

        ParaElLog = ""
    End If
    
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    Me.LabelIndF(0).Caption = ""
    Me.LabelIndF(1).Caption = ""
    cmdRenumFra.Enabled = True
    
    
    
    
    
    If Ok Then Unload Me
        
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            CargarAsientosBloqueados
        Case 2
            SugerirValoresNuevaEmpresa
        Case 3
            'Cargo los valores
            Text2(3).Text = vEmpresa.numnivel
            Text2(4).Text = vEmpresa.DigitosUltimoNivel
            
            I = vEmpresa.numnivel
            I = I - 1
            I = DigitosNivel(I)
            Text2(5).Text = I
        Case 4
            PonleFoco txtIVA(0)
        Case 5
            
            If vParam.CodiNume = 1 Then Me.chkUpdateNumDocum.Value = 1
        End Select
    End If
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    PrimeraVez = True
    Caption = "Herramientas"
    Me.Icon = frmPpal.Icon
    Limpiar Me
    FrameMovCtas.Visible = False
    Me.FrDesbloq.Visible = False
    frameNuevaEmpresa.Visible = False
    FrameCeros.Visible = False
    FrameCambioIVA.Visible = False
    FrameRenumFRAPRO.Visible = False
    Select Case Opcion
    Case 0
            PonerFrameVisible Me.FrDesbloq, H, W
    Case 1
            PonerFrameVisible Me.FrameMovCtas, H, W
            Me.chkActualizarTesoreria.Visible = vEmpresa.TieneTesoreria
    Case 2
            PonerFrameVisible frameNuevaEmpresa, H, W
    Case 3
            PonerFrameVisible FrameCeros, H, W
            pb1.Visible = False
            
    Case 4
            lblIVA.Caption = ""
            PonerFrameVisible FrameCambioIVA, H, W
                
    Case 5
            PonerFrameVisible FrameRenumFRAPRO, H, W
            Me.LabelIndF(0).Caption = ""
            Me.LabelIndF(1).Caption = ""
            
            
            'No puede actualizar el campo NUMDOCUM con el numregis si no esta
            'marcada la opcion Numeroregisro en documento(vParam.CodiNume = 1)
            If vParam.CodiNume <> 1 Then
                chkUpdateNumDocum.Value = 0
                chkUpdateNumDocum.Enabled = False
            End If
            
            optRenumPeriodo(0).Value = True
            FrameTapaRenum.Visible = False
            FrameTapaRenum.BorderStyle = 0
            'Ultimo periodo liquidado
            If vParam.perfactu > 0 Then
                If vParam.periodos = 1 Then
                    'IVA MENSUAL
                    I = vParam.perfactu
                Else
                    I = vParam.perfactu * 3
                End If
                NumTablas = DiasMes(CByte(I), vParam.anofactu)
                
                txtInformacion(0).Text = Format(NumTablas, "00") & "/" & Format(I, "00") & "/" & vParam.anofactu
            
            End If
    End Select
    
    Me.Height = H
    Me.Width = W
    Me.cmdCancelar(Opcion).Cancel = True





End Sub


Private Sub PonerFrameVisible(ByRef Fr As Frame, ByRef He As Integer, ByRef Wi As Integer)
    Fr.Top = 30
    Fr.Left = 30
    Fr.Visible = True
    He = Fr.Height + 540
    Wi = Fr.Width + 120
End Sub



Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmI_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
    
End Sub

Private Sub ImgAyuda_Click(Index As Integer)
    Select Case Index
    Case 0
        SQL = "MOVER CUENTAS" & vbCrLf & String(40, "-") & vbCrLf
        SQL = SQL & "Cambia la cuenta Origen por la de destino en las tablas:" & vbCrLf
        SQL = SQL & " -apuntes, hco apuntes, facturas, facturas provedor, presupesto, inmovilizado" & vbCrLf
        SQL = SQL & " -cobros, pagos, remesas, transferencias, reclamaciones" & vbCrLf & vbCrLf
        SQL = SQL & "Si se le indica fecha bloqueo, bloquearemos la cuenta origen" & vbCrLf
        SQL = SQL & "Si desmarca actualizar cobros pagos NO actualizara en esas tablas" & vbCrLf & vbCrLf
    
    Case 1
        'i=1
        SQL = "INSERTAR DIGITO" & vbCrLf & String(40, "-") & vbCrLf
        SQL = SQL & "Inserta un digito a las cuentas de ultimo nivel. "
        SQL = SQL & vbCrLf & "Lo añade en la posicion siguiente al digito de nivel anterior al ultimo"
        SQL = SQL & vbCrLf & vbCrLf
        SQL = SQL & "No deberia trabajar nadie y habria que cambiar las aplicaciones vinculadas (ARIGES, ARIAGRO...)" & vbCrLf
    
    Case 2
        SQL = "CAMBIAR IVA" & vbCrLf & String(40, "-") & vbCrLf
        SQL = SQL & "Cambia el tipo de IVA(NO el porcentaje) para las facturas, tanto de clientes "
        SQL = SQL & vbCrLf & "como de proveedores, comprendidas entre las fechas."
        SQL = SQL & vbCrLf & "Clientes: fecha factura"
        SQL = SQL & vbCrLf & "Proveedores: fecha recepcion"
        SQL = SQL & vbCrLf & vbCrLf
        SQL = SQL & "Habria que revisar el IVA en las aplicaciones vinculadas (ARIGES, ARIAGRO...)" & vbCrLf
    
    
    Case 3
        'Renumerar frapro
        ParaElLog = ""
        
        
        SQL = "---> " & chkCompruebaContab.Caption & vbCrLf
        TablaAnt = "Comprobara que todas las facturas, si estan contabilizadas, tienen el asiento correspondiente"
        ParaElLog = ParaElLog & SQL & TablaAnt & vbCrLf & vbCrLf
        
        SQL = "---> " & chkUpdateNumDocum.Caption & vbCrLf
        TablaAnt = "Pondra en el asiento, en el campo numdocum, el nuevo numero de regristro"
        ParaElLog = ParaElLog & SQL & TablaAnt & vbCrLf & vbCrLf
        
        ParaElLog = vbCrLf & "Renumerar  registro facturas proveedor" & vbCrLf & vbCrLf & ParaElLog & vbCrLf
        ParaElLog = ParaElLog & "***** OPCION RENUMERAR" & vbCrLf
        
        SQL = "---> " & chkSALTO_numerofactura.Caption & vbCrLf
        TablaAnt = "El numero de factura es un salto que se le ha producido." & vbCrLf
        ParaElLog = ParaElLog & SQL & TablaAnt & vbCrLf & vbCrLf
        
        SQL = "---> " & "Desde / hasta" & vbCrLf
        TablaAnt = "Por si dentro del ejercicio seleccionado, solo quiero renumerar un rango. P.ej. cambio en contador" & vbCrLf
        ParaElLog = ParaElLog & SQL & TablaAnt & vbCrLf & vbCrLf
        
        ParaElLog = ParaElLog & vbCrLf & vbCrLf & "***** DESDE ULTIMO PERIODO FACTURADO" & vbCrLf
        TablaAnt = "Desde el ultimo numero de registro del perido liquidado" & vbCrLf
        TablaAnt = TablaAnt & "ira renumerando desde ahi(1 para año completo)" & vbCrLf
        ParaElLog = ParaElLog & TablaAnt & vbCrLf & vbCrLf
        
        TablaAnt = String(100, "-")
        SQL = TablaAnt & ParaElLog & TablaAnt
        ParaElLog = ""
        TablaAnt = ""
    End Select
    MsgBox SQL, vbInformation
    SQL = ""
End Sub

Private Sub imgCheck_Click(Index As Integer)

    For I = 1 To Me.ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = Index = 1
    Next
End Sub


Private Sub CargarAsientosBloqueados()
Dim IT As ListItem
    Set miRsAux = New ADODB.Recordset
    ListView1.ListItems.Clear
    SQL = "Select * from cabapu WHERE bloqactu=1  ORDER BY fechaent,numdiari,fechaent"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = miRsAux!NumDiari
        IT.SubItems(1) = Format(miRsAux!fechaent, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!Numasien, "00000")
        SQL = DBLet(miRsAux!obsdiari, "T") & "  "
        SQL = Mid(SQL, 1, 20)
        IT.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Me.cmdDesbloq.Enabled = ListView1.ListItems.Count > 0
    
End Sub

Private Sub imgcta_Click(Index As Integer)
   Set frmCta = New frmColCtas
   SQL = ""
   frmCta.DatosADevolverBusqueda = "0|1|"
   frmCta.Show vbModal
   Set frmCta = Nothing
   If SQL <> "" Then
        txtCta(Index).Text = RecuperaValor(SQL, 1)
        DtxtCta(Index).Text = RecuperaValor(SQL, 2)
        SQL = ""
        PonleFoco txtCta(Index)
    End If
    
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Set frmC = New frmCal
    SQL = ""
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    If SQL <> "" Then
        txtFecha(Index).Text = SQL
        SQL = ""
        PonleFoco txtFecha(Index)
    End If
    Set frmC = Nothing
    
End Sub

Private Sub imgiva_Click(Index As Integer)
    SQL = ""
    Set frmI = New frmIVA
    frmI.DatosADevolverBusqueda = "0|1|"
    frmI.Show vbModal
    Set frmI = Nothing
    If SQL <> "" Then
        txtIVA(Index).Text = RecuperaValor(SQL, 1)
        txtDescIVA(Index).Text = RecuperaValor(SQL, 2)
        If Index = 0 Then
            PonleFoco txtIVA(1)
        Else
            PonleFoco txtFecha(4)
        End If
    End If
End Sub

Private Sub optFrapro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optRenumPeriodo_Click(Index As Integer)
    Me.FrameTapaRenum.Visible = optRenumPeriodo(1).Value
    
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 2 Then
        If Text2(2).Text <> "" Then
            If Not IsNumeric(Text2(2).Text) Then
                Text2(2).Text = ""
            Else
                Text2(2).Text = Val(Text2(2).Text)
                If Val(Text2(2).Text) > 99 Then
                    MsgBox "De uno a 99", vbExclamation
                    Text2(2).Text = ""
                End If
            End If
            If Text2(2).Text = "" Then PonleFoco Text2(2)
        End If
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
            imgcta_Click Index
        End If
    End If
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    
    KEYpress KeyAscii
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
    Case 10000   'Las que no sean obligadas de ultimo nivel
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


Private Sub txtFecha_GotFocus(Index As Integer)
    PonFoco txtFecha(Index)
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index))
    If txtFecha(Index) = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index), vbExclamation
        txtFecha(Index).Text = ""
        txtFecha(Index).SetFocus
    End If
End Sub



Private Sub HacerF1()
    Select Case Opcion
    Case 0
        
    Case 1
        
    End Select
End Sub


'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'
'   Cambio cuenta contable
'
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------

Private Function HacerCambioCuenta() As Boolean


Dim NombreArchivo As String
Dim NF As Integer
Dim Final As String

    On Error GoTo EHacerCambioCuenta
    
     HacerCambioCuenta = False
    
    'Veamos cuantos updates hay que hacer
    'Los fijos
            'ctabancaria    4    Dicimebre 2012. La cta ppal bancaria NO se puede cambiar y se añaden ctaingreso,ctaefectosdesc,ctagastostarj, DIC 2012
            'sinmov         3
        
        
    'Variables
            'linapu         2
            'hlinapu        2
            'linasipre      2
            'cabfactprov    1
            'cabfact        1
            'linfact        1
            'linfactprov    1
            'presupuestos   1
    
    'Total                  16
    Tamanyo = 16
    
    
    'Si tiene tesoreria
            'scaja          2

            'scobro         3
            'spagop         3
            'shacaja        2
            'sgastfij       2
            'stransfer      1
            'stransfercob   1
            'remesas        1
            'shcocob        1
            '_________________
            '               17
    If vEmpresa.TieneTesoreria Then Tamanyo = Tamanyo + 16
    Tam2 = 0
    Label16.Caption = "Comienzo proceso"
    Label16.Visible = True
    'Los que no llevan fechas
    'CTAbancaria,sinmov
    PonerTabla "ctabancaria"
    'EjecutaSQLCambio "codmacta", ""     'DIC2012    NO se puuede cambiar
    'ctaingreso,ctaefectosdesc,ctagastostarj, DIC 2012
    EjecutaSQLCambio "ctagastos", ""
    EjecutaSQLCambio "ctaingreso", ""
    EjecutaSQLCambio "ctaefectosdesc", ""
    EjecutaSQLCambio "ctagastostarj", ""
    
    PonerTabla "sinmov"
    EjecutaSQLCambio "codmact1", ""
    EjecutaSQLCambio "codmact2", ""
    EjecutaSQLCambio "codmact3", ""
    
    
    'linapu         2
    'hlinapu        2
    'linasipre      2
    NombreArchivo = "linapu|hlinapu|linasipre|"
    For NF = 1 To 3
        PonerTabla RecuperaValor(NombreArchivo, NF)
        Final = "fechaent"
        If NF = 3 Then Final = ""
        EjecutaSQLCambio "codmacta", Final
        EjecutaSQLCambio "ctacontr", Final
    Next NF
    
    
    'Presupuestos
    PonerTabla "presupuestos"
    EjecutaSQLCambio "codmacta", ""
    
    'cabfactprov    1
    'cabfact        1
    PonerTabla "cabfact"
    EjecutaSQLCambio "codmacta", "fecfaccl"
    PonerTabla "cabfactprov"
    EjecutaSQLCambio "codmacta", "fecrecpr"
    
    
    
    'Lineas de facturas
    PonerTabla "Lineas fracli"
    EjecutaSQLCambioLineasFras True, "fecfaccl"
    PonerTabla "Lineas frapro"
    EjecutaSQLCambioLineasFras False, "fecrecpr"
    
    
    'Si tiene tesoreria
    'scaja,departamento,scobro,spagop,shacaja,shcobro,sgatfij,stransfer,stransfercob
    If vEmpresa.TieneTesoreria Then
'        PonerTabla "departamentos"
'        EjecutaSQLCambio "codmacta", ""
        
        PonerTabla "slicaja"
        EjecutaSQLCambio "codmacta", ""
        
        PonerTabla "scobro"
        EjecutaSQLCambio "codmacta", "fecvenci"
        EjecutaSQLCambio "ctabanc1", "fecvenci"
        EjecutaSQLCambio "ctabanc2", "fecvenci"
        
        PonerTabla "sgastfij"
        EjecutaSQLCambio "ctaprevista", ""
        EjecutaSQLCambio "contrapar", ""
        
        PonerTabla "shcaja"
        EjecutaSQLCambio "ctacaja", ""
        EjecutaSQLCambio "codmacta", ""
        
        PonerTabla "shcocob"
        EjecutaSQLCambio "codmacta", ""
        
        
        PonerTabla "spagop"
        EjecutaSQLCambio "ctaprove", "fecefect"
        EjecutaSQLCambio "ctabanc1", "fecefect"
        EjecutaSQLCambio "ctabanc2", "fecefect"
        
        
        PonerTabla "stransfer"
        EjecutaSQLCambio "codmacta", ""
        
        
        PonerTabla "stransfercob"
        EjecutaSQLCambio "codmacta", ""
        
        
        PonerTabla "remesas"
        EjecutaSQLCambio "codmacta", ""
        
    End If
    
    If txtFecha(2).Text <> "" Then
        SQL = "UPDATE cuentas SET fecbloq = '" & Format(txtFecha(2).Text, FormatoFecha)
        SQL = SQL & "' WHERE codmacta = '" & Me.txtCta(0).Text & "'"
        Conn.Execute SQL
    End If
    
    ParaElLog = "Origen: " & txtCta(0).Text & " " & Me.DtxtCta(0).Text & vbCrLf
    ParaElLog = ParaElLog & "Destino: " & txtCta(1).Text & " " & Me.DtxtCta(1).Text & vbCrLf & vbCrLf
    ParaElLog = ParaElLog & "Fechas: " & txtFecha(0).Text & " - " & txtFecha(1).Text & vbCrLf
    If txtFecha(2).Text <> "" Then ParaElLog = ParaElLog & "Bloqueo: " & txtFecha(2).Text
    ParaElLog = "MOVER CTAS" & vbCrLf & ParaElLog
    vLog.Insertar 16, vUsu, ParaElLog
    ParaElLog = ""
    
    
    
    Label16.Caption = "Finalizando"
    
    frmActualizar.OpcionActualizar = 15 'Recalculo automatico
    frmActualizar.Show vbModal
    
    
    
    HacerCambioCuenta = True
    Exit Function
EHacerCambioCuenta:
    MuestraError Err.Number, TablaAnt & vbCrLf & SQL
End Function

Private Sub PonerTabla(ByRef T As String)
    TablaAnt = T
    Label16.Caption = ""
    Me.Refresh
    DoEvents
End Sub

Private Function EjecutaSQLCambio(Campo As String, CampoFecha As String) As Boolean
    Tam2 = Tam2 + 1
    Label16.Caption = Campo & " - " & TablaAnt & "    (" & Tam2 & " / " & Tamanyo & ")"
    Label16.Refresh
    SQL = "UPDATE " & TablaAnt & " SET " & Campo & " = " & txtCta(1).Text & " WHERE "
    SQL = SQL & Campo & " = " & txtCta(0).Text
    'Si tiene fechas
    If CampoFecha <> "" Then
        
        SQL = SQL & " AND " & CampoFecha & " >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then SQL = SQL & " AND " & CampoFecha & " <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    End If
    Conn.Execute SQL
End Function




Private Function EjecutaSQLCambioLineasFras(Clientes As Boolean, CampoFecha As String) As Boolean
    Tam2 = Tam2 + 1
    Label16.Caption = TablaAnt & "    (" & Tam2 & " / " & Tamanyo & ")"
    Label16.Refresh
    If Clientes Then
        SQL = "UPDATE cabfact,linfact SET codtbase='" & txtCta(1).Text & "'"
        SQL = SQL & " where cabfact.numserie=linfact.numserie and cabfact.codfaccl=linfact.codfaccl and"
        SQL = SQL & " cabfact.anofaccl=linfact.anofaccl"
    Else
        SQL = "UPDATE cabfactprov,linfactprov SET codtbase='" & txtCta(1).Text & "'"
        SQL = SQL & " where cabfactprov.numregis=linfactprov.numregis and"
        SQL = SQL & " cabfactprov.anofacpr = linfactprov.anofacpr"
    End If
    'Si tiene fechas
    
    SQL = SQL & " AND " & CampoFecha & " >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    If txtFecha(1).Text <> "" Then SQL = SQL & " AND " & CampoFecha & " <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
     
    
    SQL = SQL & " AND codtbase = '" & txtCta(0).Text & "'"
    
    Conn.Execute SQL
End Function



'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'
'       NUEVA EMPRESA
'
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
Private Function GeneracionNuevaBD() As Boolean

    GeneracionNuevaBD = False
    If Not IsNumeric(Text2(2).Text) Then
        MsgBox "Nº BD debe ser campo numérico", vbExclamation
        Exit Function
    End If
    
    
    
    'Comprobamos k la clave no esta
    TablaAnt = "nomempre"
    SQL = DevuelveDesdeBD("codempre", "usuarios.empresas", "codempre", Text2(2).Text, "T", TablaAnt)
    If SQL <> "" Then
        MsgBox "El codigo de empresa " & Text2(2).Text & " esta asociado a " & TablaAnt, vbExclamation
        Exit Function
    End If
        
    'Hago un SQL para que de error si no existe la BD
    SQL = "UPDATE conta" & Text2(2).Text & ".cabapu SET numdiari=1 WHERE numdiari=-1"
    If EjecutaSQL(SQL) Then
        MsgBox "YA existe la BD ", vbExclamation
        Exit Function
    End If
    
    
    
    
    If Not GeneraNuevaBD Then Exit Function
    Screen.MousePointer = vbHourglass
    'Insertamos en tabla empresas
        SQL = "INSERT INTO usuarios.empresas (codempre, nomempre, nomresum, Usuario, Pass, Conta,Tesor) VALUES ("
        SQL = SQL & Text2(2).Text & ",'" & Text2(0).Text & "','" & Text2(1).Text
        SQL = SQL & "','','','conta" & Text2(2).Text & "'," & Abs(vEmpresa.TieneTesoreria) & ")"
        Conn.Execute SQL
    
    
   If Not CrearEstructura Then Exit Function
        
   If InsercionDatos Then GeneracionNuevaBD = True
        
    
    
    
    Screen.MousePointer = vbDefault
    
End Function



Private Function GeneraNuevaBD() As Boolean
On Error Resume Next
       GeneraNuevaBD = False
        SQL = "CREATE DATABASE conta" & Text2(2).Text
        Conn.Execute SQL
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Creando BD"
        Else
            GeneraNuevaBD = True
        End If
End Function


'--------------------------------------------------------------------
'
'                    Crear estructura BD
'
'--------------------------------------------------------------------

Private Function CrearEstructura() As Boolean
Dim ColTablas As Collection
Dim ColCreate As Collection
Dim Bucle As Integer

    CrearEstructura = False

    Set ColTablas = New Collection
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "SHOW TABLES", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        ColTablas.Add CStr(miRsAux.Fields(0))
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ya tengo todas las tablas. Ahora para cada tabla ire buscando el show create table
    Set ColCreate = New Collection
    For Tam2 = 1 To ColTablas.Count
        SQL = ColTablas.Item(Tam2)
        miRsAux.Open "SHOW CREATE TABLE " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        TablaAnt = miRsAux.Fields(1)
        ColCreate.Add SQL & "|" & TablaAnt & "|"
        miRsAux.Close
    Next
    
    'Ya tengo los create tables
    'AHora para un bucle de 10 veces
    Bucle = 1
    Do
        Tamanyo = ColCreate.Count
        For Tam2 = Tamanyo To 1 Step -1
            TablaAnt = ColCreate.Item(Tam2)
            SQL = RecuperaValor(TablaAnt, 2) 'create table...
            TablaAnt = RecuperaValor(TablaAnt, 1)
            'TEngo que añadir el conta text2 .
            'Le quito los `
            SQL = Replace(SQL, "`", "")
            SQL = Trim(Mid(SQL, 13))
            'LE añado el contax.
            SQL = "CREATE TABLE conta" & Text2(2).Text & "." & SQL
            
            Label6.Caption = "[" & Bucle & "]" & TablaAnt & " (" & Tam2 & " /" & Tamanyo & ")"
            Label6.Refresh
            
            If EjecutaSQL(SQL) Then ColCreate.Remove Tam2
            
        Next
        Me.Refresh
        espera 0.5
        Bucle = Bucle + 1
        If ColCreate.Count = 0 Then
            Label6.Caption = "Creacion finalizada. " & Bucle - 1
            Label6.Refresh
            Bucle = 11 'YA ESTA TODO CREADO
        End If
            
        
    Loop Until Bucle > 10   'Si en 10 iteraciones no ha acabado.... vamos mal
    ''
    'Aqui ya tiene que a ver finalizado
    If ColCreate.Count > 0 Then
        'Algo va mal
        MsgBox "ALGO HA IDO MAL. "
    Else
        CrearEstructura = True
    End If
    
End Function



Private Function InsercionDatos() As Boolean
Dim RS As Recordset
Dim Linea As String
Dim Origen As String
Dim Insert As String
Dim F As Date



    On Error GoTo EInsercionDatos
    InsercionDatos = False
    
    Insert = "conta" & Text2(2).Text & "."
    Origen = "conta" & vEmpresa.codempre & "."
    
    
    'Datos basico
    Tam2 = 10
    TablaAnt = "contadores|stipcaja|stipoformapago|tipoconceptos|tipoamortizacion|"
    TablaAnt = TablaAnt & "tiporemesa|tiporemesa2|tiposituacion|tiposituacionrem|scryst|"
    
    For I = 1 To Tam2
        SQL = RecuperaValor(TablaAnt, I)
        Label6.Caption = "Datos básicos: " & SQL & " (" & I & "/" & Tam2 & ")"
        Label6.Refresh
        Linea = SQL
        'Conn.Execute "DELETE FROM " & Insert & Linea
        If EjecutaSQL(SQL) Then
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            'Conn.Execute SQL
            If Not EjecutaSQL(SQL) Then
                SQL = "Error insertando en tabla " & Insert & Linea
            Else
                SQL = ""
            End If
        Else
            SQL = "Error borrando tabla" & Insert & Linea
        End If
        If SQL <> "" Then
            SQL = SQL & ": " & Insert & vbCrLf & "El proceso continuará"
            MsgBox SQL, vbExclamation
        End If
   Next
    
    
    
    
    
    
    'Cuentas
    I = 0
    If Check1(I).Value Then
        Linea = "cuentas"
        
        Label6.Caption = Check1(I).Caption
        Label6.Refresh
        Conn.Execute "DELETE FROM " & Insert & Linea
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        SQL = SQL & " WHERE apudirec='S'"
        Conn.Execute SQL
    End If
    
    Me.Refresh
    
    'Conceptos
    I = 1
    If Check1(I).Value Then
        Linea = "conceptos"
        Label6.Caption = Check1(I).Caption
        Label6.Refresh
        Conn.Execute "DELETE FROM " & Insert & Linea
        
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
    End If
    
    
    
    'Tipos diario
    I = 2
    If Check1(I).Value Then
        Linea = "tiposdiario"
        Label6.Caption = Check1(I).Caption
        Label6.Refresh
        Conn.Execute "DELETE FROM " & Insert & Linea
        
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
    End If
    
    
    'Tipos IVA
    I = 3
    If Check1(I).Value Then
        Linea = "tiposiva"
        Label6.Caption = Check1(I).Caption
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
    End If
    
    
    'Centros de coste
    I = 5
    If Check1(I).Value Then
        Linea = "cabccost"
        Label6.Caption = Check1(I).Caption
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
        Linea = "linccost"
        Label6.Caption = Check1(I).Caption & " lineas"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
    End If
    
    
    
    
    
    'Asientos predefinidos
    I = 4
    'Se hara, aparte de si esta marcado, si estan las cuentas, conceptos,centros de coste
    Tam2 = Check1(0).Value + Check1(1).Value + Check1(2).Value + Check1(5).Value
    If Check1(I).Value Then
        If Tam2 = 4 Then
            Linea = "cabasipre"
            Label6.Caption = Check1(I).Caption
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
            
            Linea = "linasipre"
            Label6.Caption = Check1(I).Caption & " lineas"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
        End If
    End If
    
    

        
    'Configuracion Balances
    I = 6
    If Check1(I).Value Then
        
            Linea = "sbalan"
            
            Label6.Caption = "Balances 1/3"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
            
            
            Linea = "sperdid"
            Label6.Caption = "Balances 2/3"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
            
            
            Linea = "sperdi2"
            Label6.Caption = "Balances 3/3"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL

    End If
        
    
    
    I = 7
    If Check1(I).Value Then
        
        Linea = "sforpa"
        Label6.Caption = Check1(I).Caption
        Label6.Refresh
        Conn.Execute "DELETE FROM " & Insert & Linea
        
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
            

    End If
    
    
    
    '-----------------------------------------------------------
    'dATOS FIJOS COMO EMPRESA,EMPRESA2, PARAMETROS
        'Asientos predefinidos

    
    'Empresa
        Linea = "empresa"
        Label6.Caption = "Datos Empresa"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL

        
        
        Linea = "empresa2"
        Label6.Caption = "Datos Empresa"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
        
   'Plan contables y actualizar contadores
    Label6.Caption = "Subcuentas"
    Label6.Refresh
    Linea = "cuentas"
    SQL = "INSERT INTO " & Insert & Linea
    SQL = SQL & " SELECT * FROM " & Origen & Linea
    SQL = SQL & " WHERE apudirec<>'S'"
    Conn.Execute SQL
    
    
    
    'Contadores
    Label6.Caption = "Contadores"
    Label6.Refresh
    SQL = "UPDATE " & Insert & "contadores Set contado1=1, contado2=1 WHERE tiporegi='0'"
    Conn.Execute SQL
    SQL = "UPDATE " & Insert & "contadores Set contado1=0, contado2=0 WHERE tiporegi<>'0'"
    Conn.Execute SQL
        
        
        
        
        
    '----------
    'parametros
    '----------------------------------------
    'Los parametros solo podran ser insertado SI se piden ctas, conce y diarios
    
    SQL = ""
    If Check1(1).Value = 0 Or Check1(3).Value = 0 Then SQL = "1"
    If Check1(0).Value = 0 Then SQL = SQL & "1"
    If Len(SQL) = 0 Then
        
    
        Linea = "parametros"
        Label6.Caption = "Parámetros"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
        espera 0.5
    
    
        'En parametros
        F = CDate(txtFecha(3).Text)
        F = DateAdd("yyyy", 1, F)
        F = DateAdd("d", -1, F)

            SQL = "UPDATE " & Insert & "parametros SET fechaini='" & Format(txtFecha(3).Text, "yyyy-mm-dd")
            SQL = SQL & "', fechafin='" & Format(F, "yyyy-mm-dd") & "'"
            Conn.Execute SQL
      
        
        
     End If
        
    'Y actualizamos a los valores k nuevos
    SQL = "UPDATE " & Insert & "empresa SET nomempre= '" & Text2(0).Text & "', nomresum= '" & Text2(1).Text & "',codempre =" & Text2(2).Text
    Conn.Execute SQL
        
        
        
    InsercionDatos = True
        
    Exit Function
    
    
EInsercionDatos:
        MuestraError Err.Number, Label6.Caption
End Function


Private Sub SugerirValoresNuevaEmpresa()
    SQL = "Select max(codempre) from usuarios.empresas where codempre<100"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Tam2 = DBLet(miRsAux.Fields(0), "N") + 1
    miRsAux.Close
    Set miRsAux = Nothing
    Text2(2).Text = Tam2
    
    Me.Check1(7).Visible = vEmpresa.TieneTesoreria
    Check1(7).Value = Abs(vEmpresa.TieneTesoreria)
End Sub



'--------------------------------------------------------------------
'
'                    Subir un CERO el digitos
'
'--------------------------------------------------------------------

Private Function ComprobarOk(ByRef vNivelAnterior As Byte) As Boolean
Dim vE As String
Dim UltimoNivel As Byte
    On Error GoTo EComprobarOk
    ComprobarOk = False
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '
    'Comprobamos k las tablas siguientes NO tiene registros
    '
    '
    SQL = "cabfacte|cabfactprove|linapu|linapue|"  '4
    vE = ""
    NumTablas = 1
    Set miRsAux = New ADODB.Recordset
    Do
        ParaElLog = RecuperaValor(SQL, NumTablas)
        ParaElLog = "Select count(*) from " & ParaElLog
        miRsAux.Open ParaElLog, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then
                If miRsAux.Fields(0) > 0 Then vE = vE & RecuperaValor(SQL, NumTablas) & vbCrLf
            End If
        End If
        miRsAux.Close
        NumTablas = NumTablas + 1
    Loop Until NumTablas > 4
    
    If vE <> "" Then
        SQL = "Las siguientes tablas tienen datos y deberian estar vacias" & vbCrLf
        SQL = SQL & vE
        MsgBox SQL, vbExclamation
        Exit Function
    End If
    'Comprobamos k el ultimo nivel no es 10
    miRsAux.Open "empresa", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    vE = ""
    If miRsAux.EOF Then
        vE = "No esta definida la empresa."
    Else
        UltimoNivel = DBLet(miRsAux.Fields(3), "N")
        If UltimoNivel = 0 Then
            vE = "Si definir ultimo nivel contable"
        Else
            NumTablas = DBLet(miRsAux.Fields(3 + UltimoNivel), "N")
            If NumTablas = 0 Then
                vE = "Ultimo nivel es 0. Datos incorrectos"
            Else
                If NumTablas = 10 Then
                    vE = "No se puede ampliar el ultimo nivel. Ya es 10"
                Else
                    'Fale vamos a devolver el nivel anterior al ultimo
                    vNivelAnterior = CByte(DBLet(miRsAux.Fields(3 + UltimoNivel - 1)))
                    If vNivelAnterior < 3 Or vNivelAnterior > 10 Then vE = "Error obteniendo nivel anterior"
                End If
            End If
        End If
    End If
    miRsAux.Close
    If vE <> "" Then
        MsgBox vE, vbExclamation
        Exit Function
    End If
    ComprobarOk = True
    Exit Function
EComprobarOk:
    MuestraError Err.Number, "ComprobarOk." & Err.Description
End Function





Private Function AgregarCuentasNuevas() As Boolean
Dim Izda As String
Dim Der As String

    Label3.Caption = "Crear nueva estructura PGC"
    Label3.Refresh
    
    AgregarCuentasNuevas = False
   
    Set miRsAux = New ADODB.Recordset
    SQL = "Select count(*) from cuentas where apudirec='S'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim = 0 Then
        MsgBox "Ninguna cuenta de ultimo nivel. La aplicacion finalizara", vbCritical
        End
    End If
    NumRegElim = NumRegElim + 1
    
    SQL = "Select * from cuentas where apudirec='S'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    BACKUP_TablaIzquierda miRsAux, Izda
    Izda = "INSERT INTO cuentas " & Izda & " VALUES "
    Tamanyo = 0
    pb1.Value = 0
    While Not miRsAux.EOF
           Tamanyo = Tamanyo + 1
           PonerProgressBar (CLng(Tamanyo / NumRegElim * 1000))
           DatosTabla miRsAux, Der
           SQL = Izda & Der
           Conn.Execute SQL
           espera 0.001
           miRsAux.MoveNext
           If (Tamanyo \ 75) = 0 Then DoEvents
    Wend
    miRsAux.Close
    AgregarCuentasNuevas = True
    pb1.Value = 0
End Function


Private Sub DatosTabla(ByRef RS As ADODB.Recordset, ByRef Derecha As String)
Dim I As Integer
Dim nexo As String
Dim Valor As String
Dim Tipo As Integer
    Derecha = ""
    nexo = ""
    For I = 0 To RS.Fields.Count - 1
        Tipo = RS.Fields(I).Type
        
        If IsNull(RS.Fields(I)) Then
            Valor = "NULL"
        Else
        
            'pruebas
            Select Case Tipo
            'TEXTO
            Case 129, 200, 201
                Valor = RS.Fields(I)
                NombreSQL Valor
                'Si el campo es el codmacta o apudirec lo cambiamos
                If I = 0 Then
                    Valor = CambioCta(Valor)
                Else
                    If I = 2 Then Valor = "P"                    'de PROVISIONAL
                End If
                Valor = "'" & Valor & "'"
            'Fecha
            Case 133
                Valor = CStr(RS.Fields(I))
                Valor = "'" & Format(Valor, "yyyy-mm-dd") & "'"
                
            'Numero normal, sin decimales
            Case 2, 3, 16 To 19
                Valor = RS.Fields(I)
            
            'Numero con decimales
            Case 131
                Valor = CStr(RS.Fields(I))
                Valor = TransformaComasPuntos(Valor)
            Case Else
                Valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                Valor = Valor & vbCrLf & "SQL: " & RS.Source
                Valor = Valor & vbCrLf & "Pos: " & I
                Valor = Valor & vbCrLf & "Campo: " & RS.Fields(I).Name
                Valor = Valor & vbCrLf & "Valor: " & RS.Fields(I)
                MsgBox Valor, vbExclamation
                MsgBox "El programa finalizara. Avise al soporte técnico.", vbCritical
                End
            End Select
                        
        End If
        Derecha = Derecha & nexo & Valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
End Sub

Private Function CambioCta(Cta As String) As String
Dim Cad As String



    Cad = Mid(Cta, 1, CInt(Text2(5).Text))
    Cad = Cad & "0" & Mid(Cta, CInt(Text2(5).Text) + 1)
    CambioCta = Cad
End Function

Private Function HacerInsercionDigitoContable() As Boolean
    
    On Error GoTo EHacerInsercionDigitoContable
    HacerInsercionDigitoContable = False
    
    'Agregamos las cuentas nuevas con el numero correspondiente
    If AgregarCuentasNuevas Then
        Me.Refresh
        'Ahora hemos creado las cuentas con un digito mas
        'Ahora tendremos k ir tabla por tabla cambiando las cuentas a nivel nuevo
            
        
       'Facturas
       '------------------------------------------
       'Cabeceras
       CambiaTabla "cabfact", "codmacta|cuereten|", 2
       CambiaTabla "cabfact1", "codmacta|cuereten|", 2
       CambiaTabla "cabfactprov", "codmacta|cuereten|", 2
       CambiaTabla "cabfactprov1", "codmacta|cuereten|", 2
       
       
       'Linapus
       CambiaTabla "hlinapu", "codmacta|ctacontr|", 2
       CambiaTabla "hlinapu1", "codmacta|ctacontr|", 2
       CambiaTabla "linasipre", "codmacta|ctacontr|", 2
    
        'Linea facturas
       CambiaTabla "linfact", "codtbase|", 1
       CambiaTabla "linfact1", "codtbase|", 1
       CambiaTabla "linfactprov", "codtbase|", 1
       CambiaTabla "linfactprov1", "codtbase|", 1
       
       
       CambiaTabla "departamentos", "codmacta|", 1
       CambiaTabla "parametros", "ctaperga|", 1
       CambiaTabla "presupuestos", "codmacta|", 1
       CambiaTabla "sbasin", "codmacta2|", 1
       CambiaTabla "sinmov", "codprove|codmact1|codmact2|codmact3|", 4
       CambiaTabla "norma43", "codmacta|", 1
   
       
       'Tipos de iva
       CambiaTabla "tiposiva", "cuentare|cuentarr|cuentaso|cuentasr|cuentasn|", 5
   
       'Saldos en hco y de analitica
       'Ahora no pq recalcuo los saldos
       Label3.Caption = "Eliminando saldos"
       Label3.Refresh
       Conn.Execute "DELETE FROM hsaldos"
       Conn.Execute "DELETE FROM hsaldos1"
       Conn.Execute "DELETE FROM hsaldosanal"
       Conn.Execute "DELETE FROM hsaldosanal1"
       
       
       
       If vEmpresa.TieneTesoreria Then
            CambiaTabla "scobro", "codmacta|ctabanc1|ctabanc2|", 3
            CambiaTabla "remesas", "codmacta|", 1
            CambiaTabla "scarecepdoc", "codmacta|", 1
            CambiaTabla "sgastfij", "ctaprevista|", 1
            CambiaTabla "shcocob", "codmacta|", 1
            CambiaTabla "slicaja", "codmacta|", 1
            CambiaTabla "spagop", "ctaprove|ctabanc1|ctabanc2|", 3
            CambiaTabla "stransfer", "codmacta|", 1
            CambiaTabla "stransfercob", "codmacta|", 1
            CambiaTabla "susucaja", "codmacta|", 1
            
       End If
            
       'Quitamos las cuentas 'S'
       SQL = "Delete from Cuentas where apudirec='S'"
       Conn.Execute SQL
       
       'Las k eran apuntes directos P pasan a ser S
       SQL = "UPDATE Cuentas SET apudirec='S' where apudirec='P'"
       Conn.Execute SQL
       
       'Actualizamos en empresas
       AumentarEmpresaDigitoUltimoNivel
       
       
       
       'Creamos las cuentas de subnivel
       CrearSubNivel
       pb1.Value = 0
       Label3.Caption = ""
       Label3.Refresh
       vEmpresa.Leer vEmpresa.codempre
       vParam.Leer
       
       HacerInsercionDigitoContable = True
       
       
    End If
    Exit Function
EHacerInsercionDigitoContable:
    MuestraError Err.Number, "Errorfatal." & vbCrLf & Err.Description
End Function



Private Function CambiaTabla(Tabla As String, VCampos As String, NCampos As Integer)
Dim I As Integer

    ReDim Campos(NCampos)
    
    For I = 1 To NCampos
        Campos(I) = RecuperaValor(VCampos, I)
    Next I
    
    Label3.Caption = Tabla
    pb1.Value = 0
    Me.Refresh
    CambiaValores Tabla, NCampos

End Function

'
'Como tengo que recalcular saldos NO updateo las cuentas            |
'-------------------------------------------------------------------|
'Private Sub CambiaHSaldos(Tabla As String)
'Dim Cad As String
'
'    SQL = "SELECT " & Tabla & ".codmacta FROM " & Tabla & " INNER JOIN cuentas ON "
'    SQL = SQL & Tabla & ".codmacta = cuentas.codmacta WHERE cuentas.apudirec='S'"
'    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not miRsAux.EOF
'        Cad = CambioCta(miRsAux.Fields(0))
'        SQL = "UPDATE " & Tabla & " SET codmacta = '" & Cad
'        SQL = SQL & "' WHERE codmacta = '" & miRsAux.Fields(0) & "'"
'        Conn.Execute SQL
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'
'End Sub



Private Sub AumentarEmpresaDigitoUltimoNivel()

    
    
    I = vEmpresa.numnivel
    SQL = "UPDATE empresa SET numdigi" & CStr(I) & " = "
    I = CInt(Text2(5).Text) + 1
    SQL = SQL & CStr(I)
    I = vEmpresa.numnivel + 1
    SQL = SQL & ", numdigi" & CStr(I) & " = " & vEmpresa.DigitosUltimoNivel + 1
    SQL = SQL & ", numnivel = numnivel +1"
    
   
    
    Conn.Execute SQL
End Sub


Private Function CambiaValores(Tabla As String, numCta As Integer)
Dim SQL As String
Dim Cad As String
Dim I As Integer
    Cad = ""
    SQL = ""
    On Error GoTo ECambia
    
    For I = 1 To numCta
        'Para bonito
        Label3.Caption = Tabla & " (" & I & " de " & numCta & ")"
        pb1.Value = 0
        Me.Refresh
        Tamanyo = 0
        'Contador  COUNT(distinct(codmacta))
        SQL = "SELECT COUNT(DISTINCT(" & Campos(I) & ")) from " & Tabla
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not miRsAux.EOF Then Tamanyo = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        

        If Tamanyo > 0 Then
            'Updateamos la primera cta
            Tamanyo = Tamanyo + 1
            SQL = "SELECT " & Campos(I) & " FROM " & Tabla & " GROUP BY " & Campos(I)
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                PonerProgressBar Val((NumRegElim / Tamanyo) * 1000)
                If Not IsNull(miRsAux.Fields(0)) Then
                    Cad = CambioCta(miRsAux.Fields(0))
                    SQL = "UPDATE " & Tabla & " SET " & Campos(I) & " = '" & Cad & "'"
                    SQL = SQL & " WHERE " & Campos(I) & " = '" & miRsAux.Fields(0) & "'"
                    Conn.Execute SQL
                End If
                'Sig
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
    Next I
    Exit Function
ECambia:
    MuestraError Err.Number, Err.Description
End Function

Private Sub PonerProgressBar(Valor As Long)
    If Valor <= 1000 Then pb1.Value = Valor
End Sub



Private Sub CrearSubNivel()
Dim Col As Collection

    Label3.Caption = "Subniveles a crear (leyendo)"
    Label3.Refresh
    pb1.Value = 0
    I = CInt(Text2(5).Text) + 1
    SQL = "select substring(codmacta,1," & I & "),nommacta from cuentas where apudirec='S' group by 1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not miRsAux.EOF
        Col.Add CStr(miRsAux.Fields(0)) & "|" & DBLet(miRsAux.Fields(1), "T") & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ya tengo los subniveles que tengo que crear
    
    Label3.Caption = "Subniveles a crear (insertando)"
    Label3.Refresh
    espera 0.3
    DoEvents
    
    I = CInt(Text2(5).Text)
    SQL = String(I, "_")
    SQL = "Select codmacta,nommacta from cuentas where codmacta like '" & SQL & "'"
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    ParaElLog = "INSERT INTO cuentas(apudirec,model347,codmacta,nommacta,razosoci) VALUES ('N',0,'"
    Tam2 = CInt(Text2(5).Text)
    For I = 1 To Col.Count
        PonerProgressBar CLng((I / Col.Count) * 1000)
        TablaAnt = RecuperaValor(Col.Item(I), 1)
        SQL = ""
        miRsAux.Find "Codmacta = '" & Mid(TablaAnt, 1, Tam2) & "'", , adSearchForward, 1
        If Not miRsAux.EOF Then SQL = DBLet(miRsAux!nommacta, "T")
        If SQL = "" Then SQL = RecuperaValor(Col.Item(I), 1)
        If SQL = "" Then SQL = "Aumentando ceros"
        SQL = DevNombreSQL(SQL)
        TablaAnt = ParaElLog & TablaAnt & "','" & SQL & "','" & SQL & "')"
        Conn.Execute TablaAnt
    Next I
End Sub



Private Sub txtIVA_GotFocus(Index As Integer)
    PonFoco txtIVA(Index)
End Sub

Private Sub txtIVA_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtIVA_LostFocus(Index As Integer)
    SQL = ""
    txtIVA(Index).Text = Trim(txtIVA(Index).Text)
    If txtIVA(Index).Text <> "" Then
        If EsNumerico(txtIVA(Index).Text) Then
            SQL = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", txtIVA(Index).Text)
            If SQL = "" Then MsgBox "No existe el tipo de iva: " & txtIVA(Index).Text, vbExclamation
               
        End If
        If SQL = "" Then
            txtIVA(Index).Text = ""
            PonleFoco txtIVA(Index)
        End If
    End If
    txtDescIVA(Index).Text = SQL
End Sub



Private Function HacerCambioIVA() As Boolean

    HacerCambioIVA = False
    If CambioIVA(True) Then
        If CambioIVA(False) Then
            HacerCambioIVA = True
            
            'EL LOG
            ParaElLog = "IVA origen:     " & txtIVA(0).Text & "  -  " & txtDescIVA(0).Text & vbCrLf
            ParaElLog = ParaElLog & "IVA destino:    " & txtIVA(1).Text & "  -   " & txtDescIVA(1).Text & vbCrLf
            ParaElLog = ParaElLog & "Fechas: " & txtFecha(4).Text & " - " & txtFecha(5).Text
            ParaElLog = "CAMBIO IVA" & vbCrLf & ParaElLog
            vLog.Insertar 16, vUsu, ParaElLog
            ParaElLog = ""
            
        End If
            
    End If
    
End Function



Private Function CambioIVA(Clientes As Boolean) As Boolean
    
    'NO HACE FALTA transaccionar.
'    Conn.CommitTrans
    CambioIVA = False
    For I = 1 To 3
        If Clientes Then
            lblIVA.Caption = "Clientes"
        Else
            lblIVA.Caption = "Proveedores"
        End If
        lblIVA.Caption = lblIVA.Caption & ".  Iva " & I
        lblIVA.Refresh
        
        If Clientes Then
            SQL = "UPDATE cabfact SET tp" & I & "faccl = " & txtIVA(1).Text
            SQL = SQL & " WHERE tp" & I & "faccl = " & txtIVA(0).Text
            TablaAnt = "fecfaccl"
        Else
            SQL = "UPDATE cabfactprov SET tp" & I & "facpr = " & txtIVA(1).Text
            SQL = SQL & " WHERE tp" & I & "facpr = " & txtIVA(0).Text
            TablaAnt = "fecrecpr"
        End If
        If txtFecha(4).Text <> "" Then SQL = SQL & " AND " & TablaAnt & ">= '" & Format(txtFecha(4).Text, FormatoFecha) & "'"
        If txtFecha(5).Text <> "" Then SQL = SQL & " AND " & TablaAnt & "<= '" & Format(txtFecha(5).Text, FormatoFecha) & "'"
    
    
        If Not EjecutaSQL(SQL) Then
            'Se ha producido un error
            TablaAnt = "Error grave." & vbCrLf & "Cambiando IVA " & I & vbCrLf & vbCrLf
            TablaAnt = TablaAnt & "Desc: " & SQL & vbCrLf & "Avise a soporte técnico con el error"
            MsgBox TablaAnt, vbCritical
            Exit Function
        End If
    Next I
    
    CambioIVA = True
        
End Function


'-------------------------------------------------------------------
'
'RENUMERAR FRA PROVEEDORES
'
'-------------------------------------------------------------------
Private Sub txtRenumFrapro_GotFocus(Index As Integer)
    PonFoco txtRenumFrapro(Index)
End Sub

Private Sub txtRenumFrapro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtRenumFrapro_LostFocus(Index As Integer)
    txtRenumFrapro(Index).Text = Trim(txtRenumFrapro(Index).Text)
    If txtRenumFrapro(Index).Text = "" Then Exit Sub
    
    If Not IsNumeric(txtRenumFrapro(Index).Text) Then
        MsgBox "Campo debe ser numérico", vbExclamation
        txtRenumFrapro(Index).Text = ""
        PonleFoco txtRenumFrapro(Index)
    End If
End Sub




Private Function HacerRenumeracionFacturas() As Boolean
Dim Fecha As Date
Dim F2 As Date
Dim Finicio As Date
Dim AnoPartido As Boolean
Dim Ok As Boolean

    On Error GoTo EHacerRenumeracionFacturas
    HacerRenumeracionFacturas = False
    

    SQL = "Select fechaini,codinume from parametros"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Finicio = miRsAux!fechaini
    'Si no graba numdocum, SEGURO que no lo updateamos
    If miRsAux!CodiNume = 2 Then Me.chkUpdateNumDocum = 0
    miRsAux.Close
    
    'Fecha INCIO en actual o siguiente
    If Me.optFrapro(1).Value Then Finicio = DateAdd("yyyy", 1, Finicio)


    LabelIndF(0).Caption = "Realizando comprobaciones"
    LabelIndF(1).Caption = ""
    Me.Refresh
    DoEvents
    
    Fecha = Finicio
    F2 = DateAdd("yyyy", 1, Fecha)
    F2 = DateAdd("d", -1, F2)
    AnoPartido = Year(Fecha) <> Year(F2)
    
    'ContadorInserciones --> Numregelim
    NumRegElim = Val(txtRenumFrapro(0).Text)
    SQL = "Select count(*) from cabfactprov where fecrecpr>='" & Format(Fecha, FormatoFecha) & "'"
    SQL = SQL & " AND fecrecpr<='" & Format(F2, FormatoFecha) & "'"
    If Me.chkSALTO_numerofactura.Value = 1 Then SQL = SQL & " AND numregis >= " & txtRenumFrapro(0).Text
    'Desde hasta
    AnyadeDesdeHastaNumregis
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Tam2 = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    'Ya se cuantas facturas hay.
    '1.- Si hay 0 cierro y me largo
    '2.- Si hay mas de una veo si entre las fechas del ejerccio hay alguna factura con numregis entre los valores
    '   que voy a renumerar
    If Tam2 = 0 Then
        MsgBox "ninguna factura a renumerar", vbExclamation
        Set miRsAux = Nothing
        Exit Function
    End If




    Tamanyo = 0
    If Me.chkSALTO_numerofactura.Value = 0 Then
        'Proceso normal. No voy a partir de un numero de factura
            If AnoPartido Then
                '        AÑO PARTIDO
    
                SQL = "Select count(*) from cabfactprov where anofacpr = " & Year(Fecha)
                SQL = SQL & " AND numregis >= " & NumRegElim & " and numregis<= " & NumRegElim + Tam2
                AnyadeDesdeHastaNumregis
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not miRsAux.EOF Then Tamanyo = DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
            
                SQL = "Select count(*) from cabfactprov where anofacpr = " & Year(F2)
                SQL = SQL & " AND numregis >= " & NumRegElim & " and numregis<= " & NumRegElim + Tam2
                AnyadeDesdeHastaNumregis
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not miRsAux.EOF Then Tamanyo = Tamanyo + DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
            
            
            Else
                'AÑO NORMAL
                SQL = "Select count(*) from cabfactprov where anofacpr = " & Year(Fecha)
                SQL = SQL & " AND numregis >= " & NumRegElim & " and numregis<= " & NumRegElim + Tam2
                AnyadeDesdeHastaNumregis
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Tamanyo = 0
                If Not miRsAux.EOF Then Tamanyo = DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
                
            End If
                
    Else
    
        'Voy a renumerar A partir de un salto
        SQL = "Select count(*) from cabfactprov where anofacpr = " & Year(Fecha)
        SQL = SQL & " AND numregis = " & txtRenumFrapro(0).Text
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then Tamanyo = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        
    End If
    
    If Tamanyo > 0 Then
        MsgBox "Se solaparán números de factura", vbExclamation
        Exit Function
    End If
    
    
    
    
    
    If chkCompruebaContab.Value = 1 Then
        If Not ComprobarFRAPROContabilizadas(Fecha, False) Then
            Exit Function
        End If
    End If


    'AQUI SE HACE LA RENUMERACION PROIPAMENTE DICHA
    'Proceso laaargo donde los haya
    'Puesto que hay que hacer
    '   Crear la factura 0
    '   UPDATEAR LAS lineas de FACTURA A LA 0
    '      "        la factura a su nuevo numero
    '       "    las lineas al nuevo numero
    '    Si procede, updatear NUMDOCUM
    Fecha = Finicio
    Ok = RenumeraFacturas(Fecha)


    '-----------------------------
    'Insertamos el LOG
    SQL = "siguiente"
    If optFrapro(0).Value Then SQL = "actual"
    SQL = "Ejercicio " & SQL & vbCrLf
    
    ParaElLog = SQL & "Updatear numdocum: " & chkUpdateNumDocum.Value & vbCrLf
    SQL = ""
    AnyadeDesdeHastaNumregis
    If SQL <> "" Then
        ParaElLog = ParaElLog & "Desde/Hasta: " & Mid(SQL, 4) & vbCrLf
        SQL = ""
    End If
    
    ParaElLog = ParaElLog & "Registros: " & CStr(NumRegElim) & vbCrLf
    ParaElLog = ParaElLog & vbCrLf & "NºFactura: " & txtRenumFrapro(0).Text & vbCrLf
    ParaElLog = ParaElLog & "Salto fra: " & CStr(chkUpdateNumDocum.Value) & vbCrLf
    ParaElLog = "Renumerar nºregistro (" & CStr(Ok) & ")" & vbCrLf & ParaElLog
    
    
    vLog.Insertar 16, vUsu, ParaElLog
    ParaElLog = ""


    If Ok Then HacerRenumeracionFacturas = True
    
    Exit Function
EHacerRenumeracionFacturas:
    MuestraError Err.Number, Err.Description
End Function

Private Sub AnyadeDesdeHastaNumregis()
    If txtRenumFrapro(1).Text <> "" Then SQL = SQL & " AND numregis >= " & txtRenumFrapro(1).Text
    If txtRenumFrapro(2).Text <> "" Then SQL = SQL & " AND numregis <= " & txtRenumFrapro(2).Text
End Sub





'Comprobaremos que todas las facturas que estan contbilizadas tiene asiento
Private Function ComprobarFRAPROContabilizadas(Fecha As Date, DesdePeriodo As Boolean) As Boolean
Dim F As Date
Dim NF As Integer
Dim Bucles As Byte

    On Error GoTo EComprobarFRAPROContabilizadas
    ComprobarFRAPROContabilizadas = False
    
    
    'Por velocidad dividremos el ejhercicio end tres cuatrimestres
    LabelIndF(0).Caption = "Comprobar contabilizacion facturas"
    F = Fecha
    
    If DesdePeriodo Then
        'Renumeracion del periodo. Es decir, desde fecha en adelante. Solo se hace una vez
        Bucles = 1
        Fecha = "31/12/" & Year(Fecha)  'ejercicios naturales
        F = DateAdd("d", 1, F)
        Fecha = DateAdd("d", 1, Fecha)  'luego la resta aqui abajo
    Else
        'Renumeracion normal
        Fecha = DateAdd("m", 4, Fecha)
        Bucles = 3
    End If
    
    Insert = ""
    For NF = 1 To Bucles
        
        SQL = "select numregis,fecrecpr,cabfactprov.numasien,hcabapu.numasien as na,cabfactprov.numdiari,anofacpr "
        SQL = SQL & " from cabfactprov left join hcabapu"
        SQL = SQL & " on cabfactprov.numasien=hcabapu.numasien and cabfactprov.fechaent=hcabapu.fechaent and cabfactprov.numdiari=hcabapu.numdiari"
        SQL = SQL & " where fecrecpr>='" & Format(F, FormatoFecha)
        LabelIndF(1).Caption = "Desde : " & F & "   "
        
        F = DateAdd("d", -1, Fecha)
        SQL = SQL & "' and fecrecpr<='" & Format(F, FormatoFecha) & "'"
        LabelIndF(1).Caption = LabelIndF(1).Caption & "  hasta:  " & F
        
        
        F = Fecha
        
        Fecha = DateAdd("m", 4, Fecha)
        
        DoEvents
        
        'AHora tenog el res
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If IsNull(miRsAux!Numasien) Then
                'Factura sin contabilizar
                
            Else
                If IsNull(miRsAux!NA) Then
                    'ERRROR GRAVE
                    'La factura tiene numero asiento, pero el asiento NO existe
                    If miRsAux!Numasien = 0 Then
                        'Es posible ya que hay frapro que no se contabilizan
                    
                    Else
                        Insert = Insert & miRsAux!NumRegis & " / " & miRsAux!anofacpr & ": " & Format(miRsAux!Numasien, "00000") & ";"
                    End If
                End If
            End If
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
    Next NF
    
    
    If Insert <> "" Then
        'HAY ERRORES
        NF = FreeFile
        SQL = App.path & "\" & Format(Now, "yymmdd") & "_" & Format(Now, "hhmmss") & ".txt"
        Open SQL For Output As #NF
        Print #NF, Insert
        Close #NF
        Insert = "Se han producido errores. Vea el archivo: " & vbCrLf & vbCrLf & SQL
        Insert = Insert & " Desea continuar?"
        If MsgBox(Insert, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    If DesdePeriodo Then Set miRsAux = Nothing
    
    ComprobarFRAPROContabilizadas = True
    Exit Function
EComprobarFRAPROContabilizadas:
    MuestraError Err.Number, "ComprobarFRAPROContabilizadas"
End Function



Private Function RenumeraFacturas(fec As Date) As Boolean
Dim Ok As Boolean

    On Error GoTo ERenumeraFacturas
    RenumeraFacturas = False
    
    'Creo la factura 0
    LabelIndF(0).Caption = "Generando factura 0 / 00001"
    LabelIndF(1).Caption = ""
    Me.Refresh
    SQL = "INSERT INTO cabfactprov (numregis, fecfacpr, anofacpr, fecrecpr, numfacpr, codmacta, "
    SQL = SQL & " confacpr, ba1facpr, ba2facpr, ba3facpr, pi1facpr, pi2facpr, pi3facpr, pr1facpr,"
    SQL = SQL & " pr2facpr, pr3facpr, ti1facpr, ti2facpr, ti3facpr, tr1facpr, tr2facpr, tr3facpr,"
    SQL = SQL & "totfacpr, tp1facpr, tp2facpr, tp3facpr, extranje, retfacpr, trefacpr, cuereten,"
    SQL = SQL & " numdiari, fechaent, numasien, fecliqpr, nodeducible) VALUES "
    SQL = SQL & "(0, '0000-00-00', 1, '0000-00-00', '1', '1', 'RENUM', 0, NULL, NULL, NULL, NULL, NULL, "
    SQL = SQL & "NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, "
    SQL = SQL & "NULL, NULL, NULL, NULL, '0000-00-00', 0)"
    Conn.Execute SQL
    
    
    
    
    'En esta function RENUMERA
    LabelIndF(0).Caption = "Renumerando"
    DoEvents
    
    Ok = RenumeracionReal(fec)
    
    
    
    
    'Borro la factura
    SQL = "DELETE FROM cabfactprov WHERE numregis=0 AND anofacpr=1"
    Conn.Execute SQL
    
    If Ok Then
        MsgBox "Proceso finalizado con exito", vbInformation
        RenumeraFacturas = True
    End If
    
    Exit Function
ERenumeraFacturas:
    MuestraError Err.Number, Err.Description
End Function


Private Function RenumeracionReal(fec As Date) As Boolean


    On Error GoTo ERenumeracionReal
    RenumeracionReal = False
    SQL = "Select numregis,anofacpr,numasien,fechaent,numdiari from cabfactprov where fecrecpr>='" & Format(fec, FormatoFecha)
    fec = DateAdd("yyyy", 1, fec)
    fec = DateAdd("d", -1, fec)
    SQL = SQL & "' AND fecrecpr <='" & Format(fec, FormatoFecha) & "' "
    If Me.chkSALTO_numerofactura.Value = 1 Then SQL = SQL & " AND numregis > " & Me.txtRenumFrapro(0).Text
    
    SQL = SQL & " ORDER BY fecrecpr,numregis"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Tam2 = Val(Me.txtRenumFrapro(0).Text)
    While Not miRsAux.EOF
            
            LabelIndF(1).Caption = miRsAux!NumRegis & " / " & miRsAux!anofacpr & " --> " & Tam2
            LabelIndF(1).Refresh
            NumRegElim = NumRegElim + 1
            If NumRegElim > 60 Then
                NumRegElim = 0
                Me.Refresh
                DoEvents
            End If
            
            'Updateo las lineas a la 0/1
            SQL = "UPDATE linfactprov set numregis = 0 , anofacpr=1 where numregis =" & miRsAux!NumRegis & " AND anofacpr =" & miRsAux!anofacpr
            Conn.Execute SQL
            
            'Updateo la factura
            SQL = "UPDATE cabfactprov set numregis = " & Tam2 & " where numregis =" & miRsAux!NumRegis & " AND anofacpr =" & miRsAux!anofacpr
            Conn.Execute SQL
            
            'Reestablezco las lieas
            SQL = "UPDATE linfactprov set numregis = " & Tam2 & ", anofacpr =" & miRsAux!anofacpr & " where numregis = 0 AND anofacpr = 1"
            Conn.Execute SQL
            
            If Me.chkUpdateNumDocum.Value = 1 Then
                If Not IsNull(miRsAux!Numasien) And Not IsNull(miRsAux!NumDiari) Then
                    SQL = "UPDATE hlinapu set numdocum = '" & Format(Tam2, "0000000000") & "' WHERE numasien =" & miRsAux!Numasien
                    SQL = SQL & " AND numdiari =" & miRsAux!NumDiari & " AND fechaent = '" & Format(miRsAux!fechaent, FormatoFecha) & "'"
                    Conn.Execute SQL
                End If
            End If
        
            miRsAux.MoveNext
            Tam2 = Tam2 + 1
    Wend
    miRsAux.Close
    RenumeracionReal = True
    Exit Function
    
ERenumeracionReal:
    Insert = "Error grave: " & Err.Number & vbCrLf & vbCrLf & SQL & vbCrLf & "Desc: " & Err.Description
    MsgBox Insert, vbCritical
    Insert = ""
End Function



Private Function SePuedeRenumerarPorPeriodo() As Boolean

    SePuedeRenumerarPorPeriodo = False
    
    On Error GoTo eSePuedeRenumerarPorPeriodo:
    
    '************  LA VARIABLE   "I" se asigna aqui, no tocar
    
    
    'A años patidos partidos no esta todavia realizado
    If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
        MsgBox "Años partidos. No se puede renumerar por periodo liquidado", vbExclamation
        Exit Function
    End If
    
    'Ultimo peridodo no esta entre año inicio y año fin+1
                   'Hay que comprobar que las fechas estan
                'en los ejercicios y si
                '       0 .- Año actual
                '       1 .- Siguiente
                '       2 .- Ambito
                '       3 .- Anterior al inicio
                '       4 .- Posterior al fin
    I = CInt(FechaCorrecta2(DateAdd("d", 1, CDate(txtInformacion(0).Text)))) 'un dia mas del peridod
    If I > 1 Then
        MsgBox "Fecha incorrecta (ambito o ejercicios)", vbExclamation
        Exit Function
    End If
        
        
        
        
        
    If chkCompruebaContab.Value = 1 Then
        Set miRsAux = New ADODB.Recordset
        
        If Not ComprobarFRAPROContabilizadas(CDate(txtInformacion(0).Text), True) Then Exit Function
        
        
    End If
        
        
    ' Tam2  Tamanyo
    'Voy a ver la factura mas alta a fecha (seria la ultima si estuviera bien numerado)
    '

    'EJERCICIO ACTUAL
    SQL = "fecrecpr >='" & Format(DateAdd("yyyy", I, vParam.fechaini), FormatoFecha) & "' AND"
    SQL = SQL & " fecrecpr <='" & Format(txtInformacion(0).Text, FormatoFecha)
    SQL = SQL & "' AND 1"
    SQL = DevuelveDesdeBD("max(numregis)", "cabfactprov", SQL, "1")
    If SQL = "" Then SQL = "0"
    
    If Val(SQL) = 0 Then
        If CDate(txtInformacion(0).Text) > vParam.fechafin Then
        
        
            MsgBox "Maximo valor devuelto =0", vbExclamation
            Exit Function
        End If
    End If
    
    Tam2 = Val(SQL)   'uLTIMO NUMERO DEL PERIODO
    
    'Ahora vere si en fecha posterior hay alguna factura menor que esa
    SQL = ""
    SQL = " fecrecpr <='" & Format(DateAdd("yyyy", I, vParam.fechafin), FormatoFecha) & "' AND"
    SQL = SQL & " fecrecpr > '" & Format(txtInformacion(0).Text, FormatoFecha) & "' AND 1"
    SQL = DevuelveDesdeBD("min(numregis)", "cabfactprov", SQL, "1")
    If SQL = "" Then SQL = "0"
    
    If Val(SQL) = 0 Then
        MsgBox "Minimo valor del periodo =0", vbExclamation
        Exit Function
    End If
    Tamanyo = Val(SQL)   'minimo valor de la renumeracion
    
    If Tamanyo < Tam2 Then
        MsgBox "Numero de registro menor que mayor numero del periodo", vbExclamation
        Exit Function
    End If
    
    
    
    'Comprobacion. Todos los anofacpr de estas facturas son el mismo
    SQL = "Select anofacpr,count(*) FROM cabfactprov WHERE"
    SQL = SQL & " fecrecpr <='" & Format(DateAdd("yyyy", I, vParam.fechafin), FormatoFecha) & "' AND"
    SQL = SQL & " fecrecpr > '" & Format(txtInformacion(0).Text, FormatoFecha) & "' GROUP BY anofacpr"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Tamanyo = 0
    While Not miRsAux.EOF
        Tamanyo = Tamanyo + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Tamanyo > 1 Then
        MsgBox "Facturas con años de facturas distinto. ", vbExclamation
        Exit Function
    End If
    
    
    'ULTMO, ver la utlima del año
    SQL = "anofacpr = " & Year(DateAdd("yyyy", I, vParam.fechafin)) & " AND 1"
    SQL = DevuelveDesdeBD("max(numregis)", "cabfactprov", SQL, "1")
    If SQL = "" Then SQL = "0"
    If Val(SQL) = 0 Then
        MsgBox "Error buscando ultima factura del año", vbExclamation
        Exit Function
    End If
    
    Tamanyo = Val(SQL) + 1 'A PARTIR DE aqui estan libre para el año
    If Tamanyo > 1000000000 Then
        MsgBox "No se puede establacer el hueco(>long)", vbExclamation
        Exit Function
    End If
    SePuedeRenumerarPorPeriodo = True 'se puede, adelante
    
eSePuedeRenumerarPorPeriodo:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
End Function



Private Sub PrepararParaRenumUltPer()

    'No lleva control de errores. Si explota se va al del abjo

    LabelIndF(0).Caption = "Generando factura 0 / 00001"
    LabelIndF(1).Caption = ""
    Me.Refresh
    SQL = "INSERT INTO cabfactprov (numregis, fecfacpr, anofacpr, fecrecpr, numfacpr, codmacta, "
    SQL = SQL & " confacpr, ba1facpr, ba2facpr, ba3facpr, pi1facpr, pi2facpr, pi3facpr, pr1facpr,"
    SQL = SQL & " pr2facpr, pr3facpr, ti1facpr, ti2facpr, ti3facpr, tr1facpr, tr2facpr, tr3facpr,"
    SQL = SQL & "totfacpr, tp1facpr, tp2facpr, tp3facpr, extranje, retfacpr, trefacpr, cuereten,"
    SQL = SQL & " numdiari, fechaent, numasien, fecliqpr, nodeducible) VALUES "
    SQL = SQL & "(0, '0000-00-00', 1, '0000-00-00', '1', '1', 'RENUM', 0, NULL, NULL, NULL, NULL, NULL, "
    SQL = SQL & "NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 1, NULL, NULL, 0, NULL, NULL, "
    SQL = SQL & "NULL, NULL, NULL, NULL, '0000-00-00', 0)"
    Conn.Execute SQL
    
    SQL = "DELETE from tmprenumfrapro WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
End Sub


Private Function RenumerarDesdeUltimoPeriodoLiquidacion() As Boolean
Dim Fin As Boolean
Dim UltimaFactura As String  'Llevara la ultima fact
Dim BorreFrasTMP As String

    'Si la factura destino ya existe tendremos que meterla en un hueco
    '  |12:300|15:310|   significa que la fra 12 la hemos movido al 300 y la 15 al 310
    

    On Error GoTo eRenumerarDesdeUltimoPeriodoLiquidacion
    
    PrepararParaRenumUltPer  'preara factura
    
    
    
    LabelIndF(0).Caption = "Buscando huecos"
    DoEvents
    
    
    RenumerarDesdeUltimoPeriodoLiquidacion = False
    
    '************  LAS VARIABLES
    'I          se ha asignado en el comprobar, no tocar, ers para el caluclo de fechas
    'Tam2       uLTIMO NUMERO DEL PERIODO
    'Tamanyo    primer hueco libre
    Tam2 = Tam2 + 1 'primer nuemero del nuevo periodo
    
    
    'Grabamos la temporal
    SQL = "Select " & vUsu.Codigo & ",anofacpr,numregis,0,fecrecpr,numdiari,fechaent,numasien FROM cabfactprov WHERE"
    SQL = SQL & " fecrecpr <='" & Format(DateAdd("yyyy", I, vParam.fechafin), FormatoFecha) & "' AND"
    SQL = SQL & " fecrecpr > '" & Format(txtInformacion(0).Text, FormatoFecha) & "' "

    SQL = "INSERT INTO tmprenumfrapro(codusu,anofacpr,numregisold,numregisnew,fecrecpr,numdiari,fechaent,numasien) " & SQL
    Conn.Execute SQL

    Set miRsAux = New ADODB.Recordset
    SQL = "Select * from tmprenumfrapro WHERE codusu = " & vUsu.Codigo & " ORDER BY fecrecpr,numregisold"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Insert = "UPDATE tmprenumfrapro SET numregisnew = "
    TablaAnt = " WHERE codusu = " & vUsu.Codigo & " AND numregisold ="
    While Not miRsAux.EOF
        SQL = Insert & Tam2 & TablaAnt & miRsAux!numregisold
        Conn.Execute SQL
        Tam2 = Tam2 + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
       
    'Ahora, borro los que el nuevo y el viejo sean el mismo
    LabelIndF(0).Caption = "Fras sin renumerar(I)"
    LabelIndF(1).Caption = ""
    DoEvents
    
    espera 0.5
    
    SQL = "DELETE from tmprenumfrapro where codusu = " & vUsu.Codigo
    SQL = SQL & "  AND numregisnew=numregisold"
    Conn.Execute SQL
    espera 0.5
    
    
    
    
    
    'NO DEBERIA HACERSE , pero mes es comodo. Cuando subamos version volver a revisar este trozo
InicioProceso:
    
    
    
    
    'Vamos con la renumeracion de aquellas facturas que se realizan con un unico update
    Fin = False
    I = 0
    Do
        I = I + 1
        LabelIndF(0).Caption = "Fase I.    " & I
        LabelIndF(0).Refresh
        DoEvents
        SQL = "Select * from tmprenumfrapro where codusu = " & vUsu.Codigo
        SQL = SQL & " AND not numregisnew in (select numregisold from tmprenumfrapro where codusu =" & vUsu.Codigo & ")"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText  'este ire buscando datos
        If miRsAux.EOF Then
            Fin = True
        Else
            While Not miRsAux.EOF
                RenumeraUnaSolaFrapro False
                
                'RenumeraUNAFrapro
                Conn.Execute "DELETE from tmprenumfrapro where codusu = " & vUsu.Codigo & " AND numregisold = " & miRsAux!numregisold
                
                miRsAux.MoveNext
            Wend
        End If
        miRsAux.Close
    Loop Until Fin
    
 
    
    'Renumeracion del resto
    '       Pasos:   Mover la ultima factura a Tamanyo
    '                   y acutalizamos la factura oriengen que era la ultima a esta de tamanyo
    '                EJEMPLO:
    '                   antiguo     nuevo
    '                      1          3
    '                      3          2
    '                      2          1
    '                       a)   3 pasa a la 4, y la 3 antigua pasa a la 3   (4 -->2)
    '                       b)   1   "       3
    '                       c)   2   "       1  (era la origin en el punto anterior)
    '                       d)   4(antes3)   2
    '                whl not eof vamos moviendo
    '                   cuando movemos la 3012   a la 3020  luego buscamos la factura cuyo destino sera la 3012
    '
    
    LabelIndF(0).Caption = "Abriendo hueco"
    LabelIndF(0).Refresh
    
    SQL = "Select max(numregisnew) from tmprenumfrapro WHERE codusu = " & vUsu.Codigo & " ORDER BY numregisnew DESC"
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText  'este ire buscando datos
    
    'ESTA ES NUMREGIS mas alto
    BorreFrasTMP = ""
    Fin = True
    SQL = ""
    If Not miRsAux.EOF Then
        SQL = ""
        If Not IsNull(miRsAux.Fields(0)) Then
            'Renumero la factura mas alta:
            SQL = "numregisold= " & miRsAux.Fields(0) & " AND codusu = " & vUsu.Codigo
            miRsAux.Close  'lo cierro pq lo voy a volver a abrir
            
        End If
        
        
        
        If SQL <> "" Then
            SQL = "Select * from tmprenumfrapro WHERE " & SQL
           
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RenumeraUnaSolaFrapro True
            
            SQL = "UPDATE tmprenumfrapro SET numregisold= " & Tamanyo & " WHERE codusu = " & vUsu.Codigo & " AND numregisold=" & miRsAux!numregisold
            Conn.Execute SQL
            
            'Cierro y vuelvo a abrir
            miRsAux.Close
            
            
            'VOLVEMOS AL INICIO del proceso.
            GoTo InicioProceso
            LabelIndF(0).Caption = "Renumeracion (II)"
            LabelIndF(0).Refresh
            espera 0.6
            SQL = "Select * from tmprenumfrapro WHERE codusu = " & vUsu.Codigo & " ORDER BY numregisnew DESC"
            miRsAux.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText  'este ire buscando datos
            Fin = False
        End If
    End If
        
    LabelIndF(0).Caption = "Renumeracion"
    I = 0
    While Not Fin
        I = I + 1
        If (I Mod 20) = 0 Then DoEvents
        
        
        
        RenumeraUnaSolaFrapro False
        BorreFrasTMP = BorreFrasTMP & ", " & miRsAux!numregisold
    
    
        'Busco la de origen
        SQL = "numregisnew = " & miRsAux!numregisold
        miRsAux.Find SQL, , adSearchForward, 1
        
        If Not miRsAux.EOF Then
            
            If I > 20000 Then
                'Buscare un parametros aqui
                
            End If
        Else
            LabelIndF(0).Caption = "Eliminar tmp traspasadas  "
            LabelIndF(0).Refresh
            
            SQL = miRsAux.Source
            miRsAux.Close
            
            'Borro del la tmp los datos renumerados
            If BorreFrasTMP <> "" Then
                BorreFrasTMP = "(" & Mid(BorreFrasTMP, 2) & ")" 'quito la primera coma
                BorreFrasTMP = "DELETE FROM tmprenumfrapro WHERE codusu = " & vUsu.Codigo & " AND numregisold IN " & BorreFrasTMP
                Conn.Execute BorreFrasTMP
                BorreFrasTMP = ""
                espera 0.5
            End If
            
            'Siempre deberia ser EOF
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then Fin = True
        End If
    Wend
    miRsAux.Close
    LabelIndF(0).Caption = "Fin proceso"
    RenumerarDesdeUltimoPeriodoLiquidacion = True
    MsgBox "Proceso finalizado con exito", vbInformation
    
eRenumerarDesdeUltimoPeriodoLiquidacion:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        MsgBox "AVISE SOPORTE TECNICO", vbExclamation
    End If
    Set miRsAux = Nothing

    
    'Borro esta
    SQL = "DELETE FROM cabfactprov WHERE numregis=0 AND anofacpr=1"
    Conn.Execute SQL
    
    TablaAnt = ""
End Function

'La temporal es la primera factura de todo que se renumera a count+1
Private Sub RenumeraUnaSolaFrapro(EsLaTemporal As Boolean)
        'Sin control de errores
        
        If EsLaTemporal Then
            Tam2 = Tamanyo  'al a la ultima mas uno
        Else
            Tam2 = miRsAux!numregisnew
        End If
        
        
        LabelIndF(1).Caption = miRsAux!numregisold & " / " & miRsAux!anofacpr & " --> " & Tam2
        LabelIndF(1).Refresh
        NumRegElim = NumRegElim + 1
        If NumRegElim > 40 Then
            NumRegElim = 0
            Me.Refresh
            DoEvents
        End If
            
        'Updateo las lineas a la 0/1
        SQL = "UPDATE linfactprov set numregis = 0 , anofacpr=1 where numregis =" & miRsAux!numregisold & " AND anofacpr =" & miRsAux!anofacpr
        Conn.Execute SQL
        
        'Updateo la factura
        SQL = "UPDATE cabfactprov set numregis = " & Tam2 & " where numregis =" & miRsAux!numregisold & " AND anofacpr =" & miRsAux!anofacpr
        Conn.Execute SQL
        
        'Reestablezco las lieas
        SQL = "UPDATE linfactprov set numregis = " & Tam2 & ", anofacpr =" & miRsAux!anofacpr & " where numregis = 0 AND anofacpr = 1"
        Conn.Execute SQL
        
        If Not EsLaTemporal Then
            If Me.chkUpdateNumDocum.Value = 1 Then
                If Not IsNull(miRsAux!Numasien) And Not IsNull(miRsAux!NumDiari) Then
                    SQL = "UPDATE hlinapu set numdocum = '" & Format(Tam2, "0000000000") & "' WHERE numasien =" & miRsAux!Numasien
                    SQL = SQL & " AND numdiari =" & miRsAux!NumDiari & " AND fechaent = '" & Format(miRsAux!fechaent, FormatoFecha) & "'"
                    Conn.Execute SQL
                Else
                    'Stop
                End If
            End If
        End If
       


'sin control de errores
End Sub
