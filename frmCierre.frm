VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmCierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   Icon            =   "frmCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fPyG 
      BorderStyle     =   0  'None
      Height          =   7395
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10155
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   6240
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   3
         Left            =   6210
         TabIndex        =   2
         Top             =   1110
         Width           =   975
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   7290
         TabIndex        =   87
         Top             =   1110
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   6960
         TabIndex        =   86
         Text            =   "Text1"
         Top             =   1650
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   5160
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   2850
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   6360
         TabIndex        =   84
         Top             =   2850
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   9000
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   1650
         Width           =   855
      End
      Begin VB.CommandButton cmdSimula 
         Caption         =   "&Simulacion"
         Height          =   375
         Left            =   7860
         TabIndex        =   6
         Top             =   6720
         Width           =   945
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   8940
         TabIndex        =   7
         Top             =   6720
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   50
         Top             =   6120
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   6120
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   5100
         Width           =   1095
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   47
         Top             =   4620
         Width           =   2535
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   3
         Top             =   4620
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8940
         TabIndex        =   8
         Top             =   6720
         Width           =   885
      End
      Begin VB.CommandButton cmdCierreEjercicio 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7860
         TabIndex        =   5
         Top             =   6720
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   7260
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   5670
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   7260
         TabIndex        =   45
         Top             =   6150
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   6060
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   6150
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   7260
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   5130
         Width           =   1095
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   7230
         TabIndex        =   42
         Top             =   4650
         Width           =   2535
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   2
         Left            =   6150
         TabIndex        =   4
         Top             =   4650
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   27
         Top             =   3600
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   24
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtDescDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2250
         TabIndex        =   18
         Top             =   1140
         Width           =   2535
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   0
         Left            =   1170
         TabIndex        =   1
         Top             =   1140
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb3 
         Height          =   375
         Left            =   5040
         TabIndex        =   52
         Top             =   6720
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label5 
         Caption         =   "Nº asiento"
         Height          =   195
         Index           =   18
         Left            =   8160
         TabIndex        =   93
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "ESTADO:"
         Height          =   195
         Index           =   14
         Left            =   5160
         TabIndex        =   91
         Top             =   2220
         Width           =   1545
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   5760
         Picture         =   "frmCierre.frx":030A
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Diario"
         Height          =   255
         Index           =   17
         Left            =   5160
         TabIndex        =   90
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha contabilización"
         Height          =   195
         Index           =   16
         Left            =   5160
         TabIndex        =   89
         Top             =   1650
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   15
         Left            =   5160
         TabIndex        =   88
         Top             =   2640
         Width           =   1545
      End
      Begin VB.Line GELine 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   7920
         X2              =   9840
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label GElabel1 
         Caption         =   "REGULARIZACION 8 y 9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   5040
         TabIndex        =   82
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "SIMULACION.      De perdidas y ganancias,  y del cierre de ejercicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   64
         Top             =   0
         Width           =   9540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         TabIndex        =   30
         Top             =   6720
         Width           =   1995
      End
      Begin VB.Label Label5 
         Caption         =   "Número de asiento"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   63
         Top             =   5700
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   62
         Top             =   6180
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha contabilización"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   61
         Top             =   5160
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Diario"
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   60
         Top             =   4620
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmCierre.frx":0D0C
         Top             =   4590
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Número de asiento"
         Height          =   195
         Index           =   7
         Left            =   5160
         TabIndex        =   59
         Top             =   5730
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   10
         Left            =   5160
         TabIndex        =   58
         Top             =   6210
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha contabilización"
         Height          =   195
         Index           =   11
         Left            =   5160
         TabIndex        =   57
         Top             =   5220
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Diario"
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   56
         Top             =   4650
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   5700
         Picture         =   "frmCierre.frx":170E
         Top             =   4620
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "CIERRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   55
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "APERTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5100
         TabIndex        =   54
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   1080
         X2              =   4800
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   6480
         X2              =   9720
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   315
         Left            =   180
         TabIndex        =   53
         Top             =   6720
         Width           =   4275
      End
      Begin VB.Label Label7 
         Caption         =   "PÉRDIDAS Y GANANCIAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   600
         Width           =   3015
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   2880
         X2              =   4800
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label5 
         Caption         =   "Grupo excepción"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   32
         Top             =   2160
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Nº asiento"
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   29
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   3360
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta perd. y Gan."
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha contabilización"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Diario"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   1110
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmCierre.frx":2110
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "CIERRE ejercicio. Asientos de regularización y cierre."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   17
         Top             =   60
         Width           =   9660
      End
   End
   Begin VB.Frame frameBorrarEjer 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   3615
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2160
         TabIndex        =   77
         Text            =   "Text4"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminarEjerCerr 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1440
         TabIndex        =   78
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   6
         Left            =   2520
         TabIndex        =   79
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label16 
         Caption         =   "Eliminar ejercicios cerrados."
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
         Index           =   2
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label17 
         Caption         =   "Eliminar hasta la fecha"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   1080
         Width           =   1920
      End
   End
   Begin VB.Frame fRenumeracion 
      BorderStyle     =   0  'None
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   4485
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   2790
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.CommandButton cmdRenumera 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   3330
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   3060
         TabIndex        =   12
         Top             =   3330
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ejercicio SIGUIENTE"
         Height          =   285
         Index           =   1
         Left            =   2250
         TabIndex        =   11
         Top             =   2070
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ejercicio ACTUAL"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   2070
         Width           =   1725
      End
      Begin VB.Label Label2 
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   2520
         Width           =   3525
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "No debe haber nadie mas trabajando contra esta contabilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   60
         TabIndex        =   9
         Top             =   630
         Width           =   4245
      End
   End
   Begin VB.Frame fTraspaso 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   0
      TabIndex        =   33
      Top             =   60
      Width           =   5835
      Begin VB.ComboBox cmbCierre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   81
         Text            =   "Combo1"
         Top             =   1680
         Width           =   3735
      End
      Begin ComCtl2.Animation Animation1 
         Height          =   735
         Left            =   180
         TabIndex        =   39
         Top             =   2820
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   269
         FullHeight      =   49
      End
      Begin VB.CommandButton cmdTraspasar 
         Caption         =   "&Traspasar"
         Height          =   375
         Left            =   4440
         TabIndex        =   38
         Top             =   2580
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   34
         Top             =   3180
         Width           =   1185
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   315
         Left            =   180
         TabIndex        =   40
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label14 
         Caption         =   "HASTA:"
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
         Left            =   180
         TabIndex        =   37
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label Label12 
         Caption         =   "No debe haber nadie trabajando en esta contabilidad."
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Top             =   1020
         Width           =   5475
      End
      Begin VB.Label Label11 
         Caption         =   $"frmCierre.frx":2B12
         Height          =   435
         Left            =   240
         TabIndex        =   35
         Top             =   180
         Width           =   5475
      End
   End
   Begin VB.Frame frameDescierre 
      BorderStyle     =   0  'None
      Caption         =   $"frmCierre.frx":2B9D
      Height          =   4035
      Left            =   60
      TabIndex        =   65
      Top             =   60
      Width           =   4035
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   73
         Text            =   "Text3"
         Top             =   1920
         Width           =   3435
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   2820
         TabIndex        =   71
         Top             =   3000
         Width           =   885
      End
      Begin VB.CommandButton cmdDescerrar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1680
         TabIndex        =   70
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Trabajar nadie en esta contabiliad"
         Height          =   195
         Left            =   180
         TabIndex        =   72
         Top             =   2640
         Width           =   3540
      End
      Begin VB.Label Label16 
         Caption         =   "Haber apuntes del ejercicio siguiente"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   69
         Top             =   1380
         Width           =   2610
      End
      Begin VB.Label Label18 
         Caption         =   "Estar descuadrada"
         Height          =   195
         Left            =   780
         TabIndex        =   68
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label17 
         Caption         =   "Trabajar nadie en esta contabiliad"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   67
         Top             =   600
         Width           =   2400
      End
      Begin VB.Label Label16 
         Caption         =   "Para deshacer el cierre NO debe:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame FrameTraerCerrados 
      Height          =   2415
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdTraer 
         Caption         =   "Traer"
         Height          =   375
         Left            =   4080
         TabIndex        =   96
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5280
         TabIndex        =   95
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label21 
         Caption         =   "Label21"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   99
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label21 
         Caption         =   "Label21"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   98
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label Label13 
         Caption         =   "TRAER DE EJERCICIOS CERRADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   1320
         TabIndex        =   97
         Top             =   240
         Width           =   3870
      End
   End
End
Attribute VB_Name = "frmCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0.- Renumeracion
    '1.- Perdidas y ganancias Y Cierre
    '3.- Traspasar a ejercicios cerrados
    '4.- Simulacion de cierre, es decir, mostrará un listado con
    '    los apuntes de pyg, cierre, y apertura
    '5.- DESCIERRRE
    '6.- Eliminar ejercicios cerrados
    
    '7.- TRAER de ejercicios cerrados
    
    
    
    
    '--------------------------------------------------------------
    'REGULARIZACION grupos 8 9: Concepto=961
        
    
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
    
Private PrimeraVez As Boolean
Dim Cad As String
Dim SQL As String
Dim RS As Recordset

Dim I As Integer
Dim NumeroRegistros As Long
Dim MaxAsiento As Long
Dim ImporteTotal As Currency
Dim ImportePyG As Currency






Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCierreEjercicio_Click()
Dim Ok As Boolean

    
    Ok = True
    For I = 0 To 3
        If txtDiario(I).Text = "" Then
           If I <= 2 Or (I = 3 And vParam.GranEmpresa And vParam.NuevoPlanContable) Then
                MsgBox "Seleccione el diario.", vbExclamation
                Ok = False
                Exit For
            End If
        End If
    Next I
    If Not Ok Then Exit Sub
    
    
    'Coamprobamos las cuentas 8 9
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        If Not ComprobarCierreCuentas8y9 Then Exit Sub
    End If
    
    
    
    Ok = UsuariosConectados("")
    If Not Ok Then
        SQL = "Seguro que desea cerrar el ejercicio?"
        If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
        
    Else
        'Hay usuarios conectados
        If vUsu.Nivel > 1 Then
            'NO TIENE PERMISOS
            Exit Sub
        Else
            SQL = "No es recomendado, pero, ¿desea continuar con el proceso?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
        End If
    End If
    
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    
    'Mensaje de lineas en introduccion de asientos
    Ok = True
    If IntroduccionDeApuntes(True) Then
        Ok = False
        SQL = "Hay asientos en la introducción de apuntes del ejercicio en curso."
        MsgBox SQL, vbExclamation
        
    End If
    
    
    Screen.MousePointer = vbHourglass
    

    
    'Comprobar cuadre
    If Ok Then
        'QUITAR EL COMETNARIO
        Label10.Caption = "Comprobar cuadre"
        Label10.Refresh
        Ok = ComprobarCuadre
    End If
    cmdCierreEjercicio.Enabled = False
    Me.Refresh
    espera 0.3
    Me.Refresh
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Perd. y ganacias"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        Ok = ASientoPyG
    End If
    
    
    If Ok Then
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            pb3.Value = 0
            pb3.visible = True
            Label10.Caption = "Regularizacion 8 y 9"
            Label10.Refresh
            Label3.Caption = ""
            Label3.Refresh
            Ok = ASiento8y9
        End If
    End If

    
    
    
    Me.Refresh
    DoEvents
    espera 0.2
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Cierre"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        'Hacer el cierre
        Ok = HacerElCierre
    End If
    

    If Ok Then
        'GRABO LOG
        
        vLog.Insertar 18, vUsu, "[CIERRE]"
        
    End If
    
    Screen.MousePointer = vbHourglass
    Label10.Caption = ""
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Bloquear_DesbloquearBD False
    pb3.visible = False
    
    If Ok Then Unload Me
    Screen.MousePointer = vbDefault
End Sub





Private Function ComprobarCuadre() As Boolean
    Screen.MousePointer = vbHourglass
    ComprobarCuadre = True
    CadenaDesdeOtroForm = ""
    frmMensajes.Opcion = 5
    frmMensajes.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        ComprobarCuadre = False
        If vUsu.Codigo > 0 Then
            MsgBox "Error en la comprobación del cuadre.", vbExclamation
        Else
            If MsgBox("Error en cuadre. ¿Continuar igualmente pese al riesgo?", vbQuestion + vbYesNoCancel) = vbYes Then ComprobarCuadre = True
        End If
    End If
    Me.Refresh
End Function



Private Sub cmdDescerrar_Click()
Dim Ok As Boolean
On Error GoTo EDescierre

    Label10.Caption = ""

    SQL = "Seguro que desea deshacer el cierre?"
    If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados("Deshacer cierre.", True) Then Exit Sub
    
    
    If Not ExisteAsientosDescerrar Then
        MsgBox "No existe el asiento de apertura, luego no existe el cierre para el ejercicio anterior", vbExclamation
        Exit Sub
    End If
      
   
   
   
    'Veremos si existen los apuntes en hco .
    'Queremos ver si no se han llevado a ejercicios cerrados
    If Not ExistenApuntesEjercicioAnterior Then
        MsgBox "No se han encontrado apuntes del ejercicios anteriores", vbExclamation
        Exit Sub
    End If
   
   '---------------
    'Veremos si existen los apuntes de ejercicio siguiente
    If ExistenApuntesEjercicioSiguiente Then
        MsgBox "Hay apuntes del ejercicio siguiente ", vbExclamation
        Exit Sub
    End If
   
   
    If IntroduccionDeApuntes(False) Then
        Cad = "Hay asientos en la introducción de apuntes del ejercicio siguiente."
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
   
    'pASSSWORD MOMENTANEO
    Cad = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(Cad) <> "ARIADNA" Then
        If Cad <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    'Mensaje de lineas en introduccion de asientos
    Ok = True
    If IntroduccionDeApuntes(False) Then
        Ok = False
        SQL = "Hay asientos en la introducción de apuntes del ejercicio siguiente."
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    cmdDescerrar.Enabled = False
    cmdCancel(5).Enabled = False
    Me.Refresh
    'Comprobar cuadre
    If Ok Then
        'QUITAR EL COMETNARIO
        Label10.Caption = "Comprobar cuadre"
        Label10.Refresh
        Ok = ComprobarCuadre
    End If
    
    Me.Refresh
    espera 0.3
    Me.Refresh
    If Ok Then
        Label19.Caption = "Eliminar asientos"
        Label19.Refresh
        'NO se pueden hacer mas transacciones
        'Conn.BeginTrans
        Ok = HacerDescierre
        'If Ok Then
        '    Conn.BeginTrans
        'Else
        '    Conn.RollbackTrans
        'End If
    End If
    
    'Desbloqueamos BD
    Bloquear_DesbloquearBD False
    
    If Ok Then
    
        
        vLog.Insertar 18, vUsu, "[DESHACER]"
        
    
    
        Unload Me
    Else
        cmdDescerrar.Enabled = False
        cmdCancel(5).Enabled = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EDescierre:
    MuestraError Err.Number, "Descierre ejercicio:"
End Sub

Private Sub cmdEliminarEjerCerr_Click()
Dim F As Date

    If Text4.Text = "" Then Exit Sub
    F = CDate(Text4.Text)
    NumeroRegistros = Month(F)
    MaxAsiento = Month(DateAdd("d", 1, F))
    If NumeroRegistros = MaxAsiento Then
        MsgBox "Fecha no abarca mes completo", vbExclamation
        Exit Sub
    End If
    
    Label20.Caption = ""

    SQL = "Seguro que desea eliminar de ejercicios cerrados hasta la fecha: " & Text4.Text & "?"
    If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados("") Then Exit Sub
    
    
    'pASSSWORD MOMENTANEO
    Cad = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(Cad) <> "ARIADNA" Then
        If Cad <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    
    'Ahora eliminamos de las tablas
    Conn.BeginTrans
    If EliminarEjerciciosCerrados(F) Then
        Conn.CommitTrans
        MsgBox "Proceso efectuado con éxito", vbExclamation
        cmdCancel(6).SetFocus
        
        
        vLog.Insertar 18, vUsu, "[ELIMINAR]  Ejercicios cerrados anterior: " & Text4.Text
        
        
    Else
        Conn.RollbackTrans
    End If
    'Desbloqueamos BD
    Bloquear_DesbloquearBD False
    Label20.Caption = ""
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdRenumera_Click()
Dim Ok As Boolean
    'Comprobacion si hay alguien trabajando
    
    
    
    If UsuariosConectados("", (vUsu.Nivel = 0)) Then Exit Sub
    
    
    
    If IntroduccionDeApuntes(Option1(0).Value) Then
        SQL = "Hay asientos en la introducción de apuntes pertenecientes"
        SQL = SQL & vbCrLf & "al ejercicio a renumerar. "
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    SQL = "Deberia hacer una copia de seguridad." & vbCrLf & vbCrLf
    SQL = SQL & "¿ Desea continuar igualmente ?" & vbCrLf
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    'BLOQUEAMOS LA BD
    If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    
    Ok = True
    
    
        
    If Ok Then
        'Hemos bloqueado la tbla y esta preparado para la renumeración. No hay nadie trabajando, ni lo va a haber
        Screen.MousePointer = vbHourglass
        'LOG
        If Option1(0).Value Then
            SQL = "actual"
        Else
            SQL = "siguiente"
        End If
        SQL = "[RENUMERAR]  Ejercicio " & SQL
        vLog.Insertar 18, vUsu, SQL
        SQL = ""
        
        
        pb1.visible = True
        Label2.Caption = ""
        'Ocultanmos el del fondo , para que no pegue pantallazos
        Me.Hide
        frmPpal.Hide
        Me.Show
        
        
        'Renumeramos aqui dentro
        RenumerarAsientos
        
        'Volvemos a mostrar
        Me.Hide
        frmPpal.Show
        
        
        
        
        
        pb1.visible = False
        Screen.MousePointer = vbDefault
    End If
    
    Bloquear_DesbloquearBD False
    If Ok Then Unload Me
End Sub



'Private Function UsuariosConectados() As Boolean
'Dim i As Integer
'
'Cad = OtrosPCsContraContabiliad
'UsuariosConectados = False
'If Cad <> "" Then
'    UsuariosConectados = True
'    i = 1
'    Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
'    Do
'        SQL = RecuperaValor(Cad, i)
'        If SQL <> "" Then
'            Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
'            i = i + 1
'        End If
'    Loop Until SQL = ""
'    MsgBox Me.Tag, vbExclamation
'End If
'End Function


Private Sub cmdSimula_Click()
Dim Ok As Boolean

    Ok = True
    For I = 0 To 3
        If txtDiario(I).Text = "" Then
           If I < 2 Or (I = 3 And vParam.GranEmpresa) Then
                MsgBox "Seleccione el diario.", vbExclamation
                Ok = False
                Exit For
            End If
        End If
    Next I
    If Not Ok Then Exit Sub
    
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        If Not ComprobarCierreCuentas8y9 Then Exit Sub
    End If
    
    Ok = True
    ImporteTotal = 0
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Perd. y ganacias"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        Ok = ASientoPyG
    End If
    Screen.MousePointer = vbHourglass
    Me.Refresh
    espera 0.2
    
    
    
    If Ok Then
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            pb3.Value = 0
            pb3.visible = True
            Label10.Caption = "Regularizacion 8 y 9"
            Label10.Refresh
            Label3.Caption = ""
            Label3.Refresh
            Ok = ASiento8y9
        End If
    End If
        
    
    
    
    
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Cierre"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        'Hacer el cierre
        Ok = SimulaCierreApertura
    End If
    pb3.visible = False
    Label10.Caption = ""
    Label10.Refresh
    Label3.Caption = ""
    Label3.Refresh
    Me.Refresh
    
    If Ok Then
     With frmImprimir
            .OtrosParametros = ""
            .NumeroParametros = 0
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            '.SoloImprimir = True
            'Opcion dependera del combo
            .Opcion = 39
            .Show vbModal
        End With
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdTraer_Click()
Dim B As Boolean
    'Comprobaciones
    'If Not ComprobacionesTraerCierre Then Exit Sub

    'Comprobacion si hay alguien trabajando
    If UsuariosConectados("") Then Exit Sub
    
    
    If MsgBox("El proceso puede llevar mucho tiempo" & vbCrLf & "Seguro que desea traer de ejercicios cerrados?.", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    B = TraerDeCerrados
    Bloquear_DesbloquearBD False
    If B Then
        
        vLog.Insertar 18, vUsu, "[TRAER]  Traer de ejercicios cerrados "
        
    
    
        Unload Me
    Else
        cmdTraer.Enabled = False
        Label21(0).Caption = "AVISE SOPORTE TECNICO"
        MsgBox Label21(0).Caption, vbExclamation
    End If
    
    
End Sub

Private Sub cmdTraspasar_Click()
Dim F As Date

    If cmbCierre.ListIndex < 0 Then Exit Sub
        
       
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados("") Then Exit Sub
    
    
    
    'Obtengo la fecha del cierre k me han indicado
    
    F = CDate(Format(vParam.fechafin, "dd/mm/") & cmbCierre.ItemData(cmbCierre.ListIndex))
    If F >= vParam.fechaini Then
        MsgBox "Error inesperado. FECHAS ERRONEAS EN TRASPASO", vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    If NoHayCierre(F) Then
        MsgBox "No se ha encontrado el asiento de cierre del ultimo ejercicio a traspasar", vbExclamation
        Exit Sub
    End If
    If YaHayDatosFechas(F) Then
        MsgBox "Ya existen datos en ejercicios cerrados con fechas del ejercicio a traspasar", vbExclamation
        Exit Sub
    End If
    
    'Vamos a preguntar
    Cad = "Ejercicio actual " & vParam.fechaini & "  -  " & vParam.fechafin & vbCrLf & vbCrLf
    Cad = Cad & "Fecha final traspaso: " & Format(F, "dd/mm/yyyy") & vbCrLf & vbCrLf
    Cad = Cad & "Esta seguro que desea traspasar datos a ejercicios cerrados hasta la fecha seleccionada?"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    'pASSSWORD MOMENTANEO
    Cad = InputBox("Escriba password de seguridad", "CLAVE")
    If LCase(Cad) <> "ariadna" Then
        If Cad <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    
      
        vLog.Insertar 18, vUsu, "[TRASPASAR]  Traer de ejercicios cerrados "
      

        'Hemos bloqueado la tbla y esta preparado para la renumeración. No hay nadie trabajando, ni lo va a haber
        Screen.MousePointer = vbHourglass
        Label15.Caption = ""
        Label15.Refresh
        PonerAVI
        HacerTraspaso F
        Animation1.Stop
        Animation1.visible = False
        Label15.Caption = ""
        Screen.MousePointer = vbDefault

    
    Bloquear_DesbloquearBD False
    Unload Me
    
    
End Sub



Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False
    DoEvents
    cmdCancel(Opcion).Cancel = True
    Select Case Opcion
    Case 1, 4
        PonerDatosPyG
        PonerDatosCierre
    Case 3
       CargaComboTraspaso
'    Case 2
'        PonerDatosPyG
'        PonerDatosCierre
    Case 6
        Text4.SetFocus
        
    Case 7
        PonerDatosTraerCierre
    End Select
    Screen.MousePointer = vbDefault
End If

End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer

limpiar Me
PrimeraVez = True
Me.fRenumeracion.visible = False
Me.fPyG.visible = False
Me.fTraspaso.visible = False
Me.frameDescierre.visible = False
frameBorrarEjer.visible = False
Me.FrameTraerCerrados.visible = False
Select Case Opcion
Case 0
    Me.fRenumeracion.visible = True
    H = fRenumeracion.Height
    W = fRenumeracion.Width
    Label2.Caption = ""
    pb1.visible = False
    Caption = "Renumeración de asientos"
    
Case 1, 4
    fPyG.visible = True
    H = fPyG.Height + 120
    W = fPyG.Width
    Label3.Caption = ""
    Label10.Caption = ""
    pb3.visible = False
    cmdSimula.visible = (Opcion = 4)
    Me.cmdCancel(4).visible = (Opcion = 4)
    Me.cmdCierreEjercicio.visible = (Opcion = 1)
    Me.cmdCancel(1).visible = (Opcion = 1)
    Label6.visible = (Opcion = 4)
    Label4.visible = Not Label6.visible
    If Opcion = 1 Then
        Caption = "Generar asiento de pérdidas y ganancias y el cierre de ejercicio"
    Else
        Caption = "Simulación del cierre de ejercicio."
    End If
    PonerGrandesEmpresas
Case 3
    Me.fTraspaso.visible = True
    H = fTraspaso.Height + 200
    W = fTraspaso.Width
    Caption = "TRASPASO"
    Label15.Caption = ""
    Me.Animation1.visible = False
    CargaComboTraspaso
Case 5
    frameDescierre.visible = True
    H = frameDescierre.Height
    W = frameDescierre.Width
    Caption = "Deshacer cierre"
    Label19.Caption = ""
    Text3.Text = "Ejercicio actual: " & Format(vParam.fechaini, "dd/mm/yyyy") & " - " & Format(vParam.fechafin, "dd/mm/yyyy")
Case 6
    'Eliminar de ejercicios cerrados
    'QUITAR: hcabapu1,hlinapu1, hsaldos1,hsaldosanal1
    frameBorrarEjer.visible = True
    H = frameBorrarEjer.Height
    W = frameBorrarEjer.Width
    Caption = "Eliminar ejercicios cerrados"
    Label20.Caption = ""
    Text4.Text = Format(DateAdd("yyyy", -5, vParam.fechafin), "dd/mm/yyyy")
    
Case 7
    
    FrameTraerCerrados.visible = True
    H = FrameTraerCerrados.Height + 240
    W = FrameTraerCerrados.Width + 120
    Caption = "Traer de cerrados"
    Label21(0).Caption = "": Label21(1).Caption = ""
End Select


Me.Height = H + 100
Me.Width = W + 100
End Sub


Private Function IntroduccionDeApuntes(Actual As Boolean) As Boolean
Dim Ok As Boolean
    IntroduccionDeApuntes = False
    Ok = False
    SQL = CadenaFechasActuralSiguiente(Actual)
    Cad = "Select numasien from cabapu where " & SQL
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
      Ok = True
      IntroduccionDeApuntes = True
    End If
    miRsAux.Close
    
    'Si no tiene cabceceras veo si tiene lineas
    If Not Ok Then
        SQL = CadenaFechasActuralSiguiente(Actual)
        Cad = "Select numasien from linapu where " & SQL
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not miRsAux.EOF Then
          IntroduccionDeApuntes = True
        End If
        miRsAux.Close
    End If
    
    
    Set miRsAux = Nothing
End Function


Private Function CadenaFechasActuralSiguiente(Actual As Boolean) As String
Dim SQL As String
    If Actual Then
        'ACTUAL
        SQL = "fechaent >='" & Format(vParam.fechaini, FormatoFecha) & "' AND "
        SQL = SQL & "fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
    Else
        'SIGUIENTE
        Cad = Format(DateAdd("yyyy", 1, vParam.fechaini), FormatoFecha)
        SQL = "fechaent >='" & Cad & "' AND "
        Cad = Format(DateAdd("yyyy", 1, vParam.fechafin), FormatoFecha)
        SQL = SQL & "fechaent <='" & Cad & "'"
    End If
    CadenaFechasActuralSiguiente = SQL
End Function

Private Sub RenumerarAsientos()
Dim ContAsientos As Long
Dim NumeroAntiguo As Long
Dim fec As String
Dim RA As Recordset


    Set RA = New ADODB.Recordset
    
   
    'obtner el maximo
    Cad = CadenaFechasActuralSiguiente(Option1(0).Value)
    SQL = "Select max(numasien) from hcabapu where " & Cad
    RA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    MaxAsiento = 0
    If Not RA.EOF Then
        MaxAsiento = DBLet(RA.Fields(0), "N")
    End If
    RA.Close


    

    'Obtener contador
    SQL = "Select count(numasien) from hcabapu where " & Cad
    RA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ContAsientos = 0
    If Not RA.EOF Then
        ContAsientos = DBLet(RA.Fields(0), "N")
    End If
    RA.Close


    
    If MaxAsiento + ContAsientos > 99999999 Then
        MsgBox "La aplicación no tiene espacio suficiente para renumerar. Numero registros posibles mayor que la capacidad disponible.", vbCritical
        Exit Sub
    End If

    
    
    
    'Para la progresbar
    NumeroRegistros = ContAsientos

    
    'Tendremos el incremeto
    MaxAsiento = MaxAsiento + ContAsientos + 1
    
    
    Label2.Caption = "Preparación"
    Me.Refresh
    
    'actualizamos todas las tablas sumandole maxasiento al numero de asiento donode proceda
    'es decir en el ejercicio y si tiene asiento
    PreparacionAsientos MaxAsiento

    
    DoEvents
    Me.Refresh
    espera 0.01
    
    '-----------------------------------------------------------------
    ' Ahora iremos cogiendo cada registro y los iremos actualizando con
    ' los nuevos valores de numasien, tb para las tblas relacionadas
    ' Solo cambia NUMASIEN
    Cad = CadenaFechasActuralSiguiente(Option1(0).Value)
    SQL = "Select numasien,fechaent,numdiari from hcabapu where " & Cad
    SQL = SQL & " ORDER BY fechaent,numasien"
    RA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ContAsientos = MaxAsiento
    'Y maxasiento lo utilizo como contador
    If Option1(1).Value Then
        MaxAsiento = 2
        NumeroRegistros = NumeroRegistros + 1
    Else
        MaxAsiento = 1
    End If

    While Not RA.EOF
           NumeroAntiguo = RA.Fields(0)
           fec = "'" & Format(RA.Fields(1), FormatoFecha) & "'"
           If Not CambiaNumeroAsiento(NumeroAntiguo, MaxAsiento, fec, RA!NumDiari) Then
                MsgBox "Error muy grave. Se ha producido un error en la renumeracion del asiento: " & vbCrLf & NumeroAntiguo & "  --> " & MaxAsiento & " // Fecha: " & fec, vbExclamation
                End
           End If
           
           
           'progressbar
           Label2.Caption = NumeroAntiguo & " / " & RA.Fields(1)
           Label2.Refresh
           I = Int((MaxAsiento / NumeroRegistros) * pb1.Max)
           pb1.Value = I
           
           
         
           'Siguiente
           MaxAsiento = MaxAsiento + 1
           ContAsientos = ContAsientos + 1
           RA.MoveNext
           
           If (MaxAsiento Mod 50) = 0 Then
                'Me.Refresh
                DoEvents
                Me.Refresh
                espera 0.01
            End If
           
    Wend
    RA.Close
    Set RA = Nothing
    
    
    
    'En contadores ponemos el contador al numero k le toca
    MaxAsiento = NumeroRegistros
    SQL = "UPDATE contadores set "
    If (Option1(0).Value) Then
        SQL = SQL & " contado1=" & MaxAsiento
    Else
        If MaxAsiento = 1 Then MaxAsiento = 2
        SQL = SQL & " contado2=" & MaxAsiento
    End If
    SQL = SQL & " WHERE TipoRegi = '0'"
    Conn.Execute SQL
    
    
    
    '--------------------------------------------------------------------
    ' VA COMENTADO TODO LO DE ABAJO
    'Comprobacion del error
'    Cad = CadenaFechasActuralSiguiente
'    SQL = "Select numasien,fechaent from hcabapu where " & Cad
'    SQL = SQL & " ORDER BY fechaent,numasien"
'
'
'    Set RA = New ADODB.Recordset
'    RA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Set Rs = New ADODB.Recordset
'    While Not RA.EOF
'        NumeroAntiguo = RA.Fields(0)
'        Fec = "'" & Format(RA.Fields(1), FormatoFecha) & "'"
'        SQL = "Select count(linliapu) from hlinapu where numasien=" & NumeroAntiguo
'        SQL = SQL & " AND fechaent =" & Fec
'        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        MaxAsiento = 0
'        If Not Rs.EOF Then
'            MaxAsiento = DBLet(Rs.Fields(0), "N")
'        End If
'        Rs.Close
'
'        RA.MoveNext
'
'    Wend
'

End Sub



Private Function CambiaNumeroAsiento(Antiguo As Long, Nuevo As Long, Fecha As String, NuDi As Integer) As Boolean

On Error GoTo ECambia
    CambiaNumeroAsiento = False
    
    'AUX
    Cad = " SET numasien = " & Nuevo & " WHERE numasien = " & Antiguo
    Cad = Cad & " AND fechaent = " & Fecha & " AND numdiari = " & NuDi
    
    
    'Actualizamos el registro de facturas
    SQL = "UPDATE cabfact" & Cad
    Conn.Execute SQL
    
    
    'Actualizamos el registro de facturas
    SQL = "UPDATE cabfactprov" & Cad
    Conn.Execute SQL
    
    'lineas
    SQL = "UPDATE hlinapu" & Cad
    Conn.Execute SQL
    
    'cabeceras
    SQL = "UPDATE hcabapu" & Cad
    Conn.Execute SQL
    
    
    CambiaNumeroAsiento = True
    Exit Function
ECambia:
    MuestraError Err.Number, "Renumeracion tipo 1, asiento: " & Antiguo

End Function


Private Sub PreparacionAsientos(Suma As Long)
On Error GoTo EPreparacionAsientos

    SQL = CadenaFechasActuralSiguiente(Option1(0).Value)
    Cad = " Set NumASien = NumASien + " & Suma
    Cad = Cad & " WHERE numasien>0 AND " & SQL
    
    pb1.Max = 4
    
    'Facturas clientes
    Label2.Caption = "Facturas clientes"
    Label2.Refresh
    pb1.Value = 1
    SQL = "UPDATE cabfact" & Cad
    Conn.Execute SQL
    
    
        
    'Facturas proveedores
    Label2.Caption = "Facturas proveedores"
    Label2.Refresh
    pb1.Value = 2
    SQL = "UPDATE cabfactprov" & Cad
    Conn.Execute SQL
    


    'Lineas hco asiento
    Label2.Caption = "Lineas asientos"
    Label2.Refresh
    pb1.Value = 3
    SQL = CadenaFechasActuralSiguiente(Option1(0).Value)
    SQL = "Select distinct(numasien) from hlinapu WHERE " & SQL
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = CadenaFechasActuralSiguiente(Option1(0).Value)
    Cad = " Set NumASien = NumASien + " & Suma
    Cad = Cad & " WHERE numasien>0 AND " & SQL
    
    
    
    'Esto es factible de revisar
'    SQL = "UPDATE hlinapu " & Cad
'    While Not RS.EOF
'        Conn.Execute SQL & " AND numasien = " & RS!Numasien
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
    
    'Ejecutaremos esto
    SQL = "UPDATE hlinapu " & Cad
    Conn.Execute SQL
    
    
    
    'ASientos
    Label2.Caption = "Cabeceras asientos"
    Label2.Refresh
    pb1.Value = 4
    SQL = "UPDATE hcabapu " & Cad
    Conn.Execute SQL
    

    pb1.Max = 1000
    Exit Sub
EPreparacionAsientos:
    MuestraError Err.Number
    MsgBox "Error grave. Soporte técnico", vbExclamation
    Set RS = Nothing
End Sub



Private Sub PonerDatosPyG()

    'Fecha siempre la del final de ejercicio
    Text1(0).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    '8y9
    Text1(13).Text = Text1(0).Text


    NumeroRegistros = 1
    SQL = DevuelveDesdeBD("contado1", "contadores", "tiporegi", "0", "T")
    If SQL = "" Then
        MsgBox "Error obteniendo numero de asiento."
        cmdCierreEjercicio.Enabled = False
    Else
        Text1(3).Text = Val(SQL) + 1
    End If
    If vParam.GranEmpresa Then Text1(11).Text = Val(Text1(3).Text) + 1
    
    'PyG
    SQL = CuentaCorrectaUltimoNivel(vParam.ctaperga, Cad)
    If SQL = "" Then
        MsgBox "Error en la cuenta de pérdidas y ganancias de parametros.", vbExclamation
    Else
        Text1(1).Text = vParam.ctaperga
        Text2(0).Text = Cad
    End If
    
    'Concepto  --> Siempre sera nuestro 960
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "960")
    If SQL = "" Then
        MsgBox "No existe el concepto 960.", vbExclamation
    Else
        Text1(2).Text = "960"
        Text2(1).Text = SQL
    End If
    
    'Concepto para grandes empresas
    If vParam.GranEmpresa Then
        SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "961")
        If SQL = "" Then
            MsgBox "No existe el concepto 961.", vbExclamation
        Else
            Text1(12).Text = "961"
            Text2(4).Text = SQL
        End If
    End If
    

    
    'Simulacion nos salimos ya
    If Opcion = 4 Then Exit Sub
    
    
    'Si ya hay un 960 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    SQL = "Select numasien from hlinapu WHERE codconce=960 and fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        MsgBox "Ya se ha efectuado el asiento de Perdidas y ganancias : " & RS.Fields(0), vbExclamation
        cmdCierreEjercicio.Enabled = False
    End If
    RS.Close
    
    
    'Si ya hay un 961 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    If vParam.GranEmpresa Then
        SQL = "Select numasien from hlinapu WHERE codconce=961 and fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            MsgBox "Ya se ha efectuado el asiento de regularizacion : " & RS.Fields(0), vbExclamation
            cmdCierreEjercicio.Enabled = False
        End If
        RS.Close
    End If
    
    
    'Comprobamos k tampoc haya asiento 1 en ejercicio siguiente
    SQL = "Select numasien from hcabapu WHERE fechaent>'" & Format(vParam.fechafin, FormatoFecha)
    SQL = SQL & "' and numasien=1"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        MsgBox "Ya existe el asiento numero 1 para el año siguiente.", vbExclamation
        cmdCierreEjercicio.Enabled = False
    End If
    RS.Close
    Set RS = Nothing
End Sub



Private Sub PonerDatosCierre()
Dim Ok As Boolean

    'Fecha CIERRE siempre la del final de ejercicio
    Text1(7).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    Text1(9).Text = Format(DateAdd("d", 1, vParam.fechafin), "dd/mm/yyyy")
    
    'NUmero asiento
    'Es uno mas k Perdidas y ganancias
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        Text1(4).Text = Val(Text1(3).Text) + 2
    Else
        Text1(4).Text = Val(Text1(3).Text) + 1
    End If


    
    'Concepto  --> Siempre sera nuestro 980  CIERRE
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "980")
    If SQL = "" Then
        MsgBox "No existe el concepto 980.", vbExclamation
    Else
        Text1(5).Text = "980"
        Text2(2).Text = SQL
    End If
    
    'Nº asiento apertura.---- El uno
    Text1(6).Text = 1
    
    'Concepto  --> Siempre sera nuestro 970  apertura
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "970")
    If SQL = "" Then
        MsgBox "No existe el concepto 970.", vbExclamation
    Else
        Text1(8).Text = "970"
        Text2(3).Text = SQL
    End If
    
    
    
    'Si es simulacion busco el numero de diario mas pequeño
    If Opcion = 4 Then
        SQL = "Select numdiari,desdiari from tiposdiario order by numdiari"
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            For I = 0 To 3
                txtDescDiario(I).Text = RS.Fields(1)
                txtDiario(I).Text = RS.Fields(0)
            Next I
        End If
        RS.Close
        Set RS = Nothing
        'Ponemos el control en simula
        cmdSimula.SetFocus
    End If
    
    
    
    'Simulacion nos salimos ya
    If Opcion = 4 Then Exit Sub
    
    
    'Si ya hay un 980 y/o 970 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Cad = CadenaFechasActuralSiguiente(True)
    SQL = "Select numasien from hlinapu WHERE codconce=980"
    SQL = SQL & " AND " & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        MsgBox "Ya se ha efectuado el asiento de cierre de ejercicio : " & RS.Fields(0), vbExclamation
        Ok = False
    Else
        Ok = True
    End If
    RS.Close
    
    'Apertura
    If Ok Then
            Cad = CadenaFechasActuralSiguiente(False)
            SQL = "Select numasien from hlinapu WHERE codconce=980"
            SQL = SQL & " AND " & Cad
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                MsgBox "Ya se ha efectuado el asiento de cierre de ejercicio : " & RS.Fields(0), vbExclamation
                Ok = False
            Else
                Ok = True
            End If
            RS.Close
    End If
    Set RS = Nothing
    cmdCierreEjercicio.Enabled = Ok
End Sub









Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    txtDiario(I).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescDiario(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    I = Index
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub





Private Sub Text4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text4_LostFocus()
    Text4.Text = Trim(Text4.Text)
    If Text4.Text <> "" Then
        If Not EsFechaOK(Text4) Then Text4.SetFocus
    End If
    
End Sub



Private Sub txtDiario_GotFocus(Index As Integer)
    PonFoco txtDiario(Index)
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    Me.txtDescDiario(Index).Text = ""
    If txtDiario(Index).Text = "" Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtDiario(Index).Text) Then
        MsgBox "Diario debe ser numérico: " & txtDiario(Index).Text, vbExclamation
        txtDiario(Index).Text = ""
        txtDiario(Index).SetFocus
        Exit Sub
    End If
    SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text)
    If SQL = "" Then
        MsgBox "No existe el diario : " & txtDiario(Index).Text, vbExclamation
        txtDiario(Index).Text = ""
        txtDiario(Index).SetFocus
    Else
        Me.txtDescDiario(Index).Text = SQL
    End If
    
    
End Sub


Private Sub PonerAVI()
On Error GoTo EPonerAVI
Me.Animation1.Open App.path & "\actua.avi"
Me.Animation1.Play
Me.Animation1.visible = True
Exit Sub
EPonerAVI:
    MuestraError Err.Number, "Poner Video"
End Sub


Private Function ASientoPyG() As Boolean
Dim Ok As Boolean
Dim NoTieneLineas As Boolean
Dim Cuantos As Long
    
    On Error GoTo EASientoPyG
    
    ASientoPyG = False

    'Generamos los  apuntes, sobre cabapu, y luego los actualizamos
    If Opcion <> 4 Then
        SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES (" & txtDiario(0).Text
        Cad = SQL & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(3).Text & ",0,NULL,NULL)"
        
    Else
        'Estamo simulando
        'Borramos los datos de tmp
        Cad = "Delete from Usuarios.zhistoapu where codusu = " & vUsu.Codigo
    End If
        
    Conn.Execute Cad
    Ok = Cuentas6y7
    
    If Ok Then
        
        'Entonces hacemos las cuentas de otros grupos de analitica
        If Text1(10).Text <> "" Then
            Ok = Cuentas9
        End If
    End If
    
    If Ok Then
        'Cuadramos el asiento
        Ok = CuadrarAsiento
    End If
    
    
    If Opcion = 1 And Ok Then
        
        'Veremos si hay algun registro insertado
        
        SQL = "numdiari=" & txtDiario(0).Text & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien  "
        espera 0.2
        SQL = DevuelveDesdeBD("count(*)", "linapu", SQL, Text1(3).Text)
        If SQL = "" Then SQL = "0"
        NoTieneLineas = Val(SQL) = 0
        
        
        If NoTieneLineas Then
            'NO HA INSERTADO NI UNA SOLA LINEA EN  LINAPU. CON LO CUAL TENDREMOS QUE AVISAR Y BORRAR
            If MsgBox("Ningún apunte en lineas de perdidas y ganancias. ¿Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then
                Ok = False

            Else
                'COMO NO genera el apunte pq no hay lineas 6 y 7 entonces el contador de cierre disminuye
                If vParam.NuevoPlanContable And vParam.GranEmpresa Then
                    'El del 7 y 8 pasa al del CIERRE
                    Text1(4).Text = Text1(11).Text
                    'El contador del py g pasa al de 8 y 9
                    Text1(11).Text = Text1(3).Text
                    
                Else
                    Text1(4).Text = Text1(3).Text
                End If
            End If
            
            'De cualquier modo hay que borrar la cabecera que ha creado
            SQL = " where numdiari=" & txtDiario(0).Text
            SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
    
            'Borramos por si acaso ha insertado lineas
            Cad = "Delete FROM cabapu" & SQL
            Conn.Execute Cad
        End If
    End If
    
    
    If Opcion = 4 Then
        ASientoPyG = Ok
        Exit Function
    End If
    
    'Llamamos a actualizar el asiento, para pasarlo a hco
    If Ok Then
            'actualizaremos si tiene linea
            If Not NoTieneLineas Then
                Screen.MousePointer = vbHourglass
                frmActualizar.OpcionActualizar = 1
                frmActualizar.NumAsiento = CLng(Text1(3).Text)
                frmActualizar.FechaAsiento = vParam.fechafin
                frmActualizar.NumDiari = CInt(txtDiario(0).Text)
                AlgunAsientoActualizado = False
                frmActualizar.Show vbModal
                Me.Refresh
                If AlgunAsientoActualizado Then
                    Ok = True
                Else
                    Ok = False
                End If
            End If
    End If
        
        
    
    If Not Ok Then
        'Comun lineas y cabeceras
        SQL = " where numdiari=" & txtDiario(0).Text
        SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
        
        'Borramos por si acaso ha insertado lineas
        Cad = "Delete FROM linapu" & SQL
        Conn.Execute Cad
    
        'Borramos la cabcecera del apunte
        Cad = "DELETE FROM cabapu" & SQL
        Conn.Execute Cad
        
        Label3.Caption = ""
        Exit Function
    End If
    ASientoPyG = True
    
    Exit Function
EASientoPyG:
    MuestraError Err.Number, "Error en procedmiento ASientoPyG"
End Function


Private Function Cuentas6y7() As Boolean
On Error GoTo ECuentas6y7


    Cuentas6y7 = False
    Set RS = New ADODB.Recordset
    
    'Para todas las cuentas de los grupos 6 y 7  ----> Vienen en parametros
    ' calculamos su saldo y si es distinto de 0 lo insertamos en linapu
    MaxAsiento = 1
    If vParam.grupogto <> "" Then
        If Not Subgrupo(vParam.grupogto, "") Then Exit Function
    End If
    
    If vParam.grupovta <> "" Then
        If Not Subgrupo(vParam.grupovta, "") Then Exit Function
    End If
        
    Set RS = Nothing
    Cuentas6y7 = True
    Exit Function
ECuentas6y7:
    MuestraError Err.Number, "Cuentas   ventas / gastos "
End Function



Private Function Cuentas9() As Boolean
On Error GoTo ECuentas9


    Cuentas9 = False
    Set RS = New ADODB.Recordset
    
    'Para todas las cuentas de los grupos 6 y 7  ----> Vienen en parametros
    ' calculamos su saldo y si es distinto de 0 lo insertamos en linapu
    If vParam.grupoord <> "" And Text1(10).Text <> "" Then
        If Not Subgrupo(vParam.grupoord, Text1(10).Text) Then Exit Function
    End If
    

        
    Set RS = Nothing
    Cuentas9 = True
    Exit Function
ECuentas9:
    MuestraError Err.Number, "Cuentas   ventas / gastos "
End Function




Private Function Subgrupo(Primera As String, Excepcion As String) As Boolean
Dim Cont As Integer
Dim Importe As Currency
Dim AUX3 As String
Dim vCta As String

    Subgrupo = False
    'ATENCION AÑOS PARTIDOS
    Cad = Mid("__________", 1, vEmpresa.DigitosUltimoNivel - 1)
    AUX3 = Primera & Cad
    Cad = " from hsaldos"
    
    'Necesito sbaer tb nombre cta
    If Opcion = 4 Then Cad = Cad & ",cuentas"
    
    Cad = Cad & " WHERE "
    If Opcion = 4 Then Cad = Cad & " cuentas.codmacta = hsaldos.codmacta AND "
    'Por la ambiguedad del nombre
    vCta = " ("
    If Opcion = 4 Then vCta = vCta & " cuentas."
    'vCta = vCta & "codmacta like '" & AUX3 & "'"
    vCta = vCta & "codmacta like '" & AUX3 & "')"
    If Excepcion <> "" Then
        Excepcion = Mid(Excepcion & "__________", 1, vEmpresa.DigitosUltimoNivel)
        vCta = vCta & " and not ("
        If Opcion = 4 Then vCta = vCta & " cuentas."
        vCta = vCta & "codmacta like '" & Excepcion & "')"
    End If
    
    
    Cad = Cad & vCta
    
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Años son el mimo, luego en para mer, luego
        Cad = Cad & " AND ( anopsald = " & Year(vParam.fechafin) & ")"
    Else
        'Fecha inicio y fin no estan en el mismo año natural
        Cad = Cad & " AND (( anopsald = " & Year(vParam.fechaini) & " AND mespsald >=" & Month(vParam.fechaini) & ") OR "
        Cad = Cad & " ( anopsald = " & Year(vParam.fechafin) & " AND mespsald <=" & Month(vParam.fechafin) & "))"
        'ANTES del 8 marzo
        'Cad = Cad & " AND anopsald = " & Year(vParam.fechafin) & " AND mespsald <" & Month(vParam.fechafin) & ")"
    End If
    
    If Opcion = 4 Then
        Cad = "cuentas.codmacta,nommacta " & Cad & " GROUP BY cuentas.codmacta ORDER BY cuentas.codmacta"
    Else
        Cad = " codmacta " & Cad & " GROUP BY codmacta ORDER BY codmacta"
    End If
    
    'Contador
    SQL = "Select sum(impmesde)-sum(impmesha), " & Cad
    RS.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    NumeroRegistros = 0
    If RS.EOF Then
        RS.Close
        'Puede k asi este bien
        Subgrupo = True
        Exit Function
    End If
    
    'Contador
    While Not RS.EOF
        NumeroRegistros = NumeroRegistros + 1
        RS.MoveNext
    Wend
    
    'Preparamos el SQL para la insercion de lineas de apunte
    'Montamos la cadena casi al completo
    CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
    
    RS.MoveFirst
    Cont = 1
    AUX3 = "'" & Text1(1).Text & "'"
    While Not RS.EOF
    
        Label3.Caption = RS.Fields(1)
        Label3.Refresh
        I = Int((Cont / NumeroRegistros) * pb1.Max)
        pb1.Value = I
        Importe = RS.Fields(0)
        If Importe <> 0 Then
            If Opcion = 4 Then
                ' "linliapu, codmacta, nommacta, numdocum, ampconce    : numdocum. Nos ayudara con la ordenacion
                Cad = SQL & "," & MaxAsiento + Cont & ",'" & RS.Fields(1) & "','" & DevNombreSQL(RS.Fields(2)) & "','1','" & Text2(1).Text & "',"
            Else
                Cad = SQL & "," & MaxAsiento + Cont & ",'" & RS.Fields(1) & "','',960,'" & Text2(1).Text & "',"
            End If
            InsertarLineasDeAsientos Importe, AUX3
        End If
        'Sig
        Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.Close
    MaxAsiento = MaxAsiento + Cont - 1
    Subgrupo = True
End Function


Private Sub CadenaLINAPU(Diario As String, Fecha As Date, Num As String)

    If Opcion = 4 Then
        SQL = "INSERT INTO Usuarios.zhistoapu (codusu, numdiari, desdiari, fechaent, numasien,"
        SQL = SQL & "linliapu, codmacta, nommacta, numdocum, ampconce, timporteD, "
        SQL = SQL & "codccost,timporteH) VALUES ("
        SQL = SQL & vUsu.Codigo & ","
        SQL = SQL & Diario & ",'" & Me.txtDescDiario(1).Text & "'"
        SQL = SQL & ",'" & Format(Fecha, FormatoFecha) & "'," & Num
    Else
        SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
        SQL = SQL & "codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada) VALUES ("
        SQL = SQL & Diario
        SQL = SQL & ",'" & Format(Fecha, FormatoFecha) & "'," & Num
    End If
    'Primera parte es fija

End Sub

'////////////////////////////////////////////////////////////////////////
'///
'///
'/// Insertamos la linea de asiento correspondiente
Private Function InsertarLineasDeAsientos(ByRef Importe As Currency, ByRef Ctrapar As String) As Boolean
Dim Aux As String
    
        Aux = TransformaComasPuntos(CStr(Abs(Importe)))
        'Deb  centro coste haber
        If Importe < 0 Then     'si es negativo lo pongo al DEBE, si + al haber
            Aux = Aux & ",NULL,NULL"
        Else
            Aux = "NULL,NULL," & Aux
        End If
        ImporteTotal = ImporteTotal + Importe
        '                       contrapartida
        If Opcion = 4 Then
            
            Cad = Cad & Aux & ")"
        Else
            Cad = Cad & Aux & "," & Ctrapar & ",'CONTAB',0)"
            '26 Abril 2006. QUito lo de abjo
            'Cad = Cad & Aux & "," & Ctrapar & ",'idcontab',0)"
        End If
        'Ejecutamos
        Conn.Execute Cad
    
End Function


'//////////////////////////////////////////////////////////////////7
'///  Cuadramos el asiento de perdidas y ganancias
'///
Private Function CuadrarAsiento() As Boolean
Dim Importe As Currency

On Error GoTo ECua
    CuadrarAsiento = False
    
    If Opcion < 4 Then
        'SQL = "select sum(timporteD)-  sum(timporteH) from linapu "
        SQL = "select sum(timporteD),  sum(timporteH) from linapu "
        SQL = SQL & " where numdiari=" & txtDiario(0).Text
        SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Importe = 0
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then Importe = RS.Fields(0)
            If Not IsNull(RS.Fields(1)) Then Importe = Importe - RS.Fields(1)
            
           ' If Not IsNull(RS.Fields(0)) Then Importe = RS.Fields(0)
        End If
        RS.Close
        Set RS = Nothing
        
        If Importe <> 0 Then
            CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
            Cad = SQL & "," & MaxAsiento + 1 & ",'" & Text1(1).Text & "','',960,'" & Text2(1).Text & "',"
            InsertarLineasDeAsientos Importe, "NULL"
        End If
    Else
        
        'Simulación
        ' "linliapu, codmacta, nommacta, numdocum, ampconce "
        ImporteTotal = ImporteTotal * -1 'Para k cuadre
        ImportePyG = ImporteTotal
        CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
        Cad = SQL & "," & MaxAsiento + 1 & ",'" & Text1(1).Text & "','"
        ' lo meto en el uno, NO en maxasiento
        Cad = SQL & ",1,'" & Text1(1).Text & "','"
        Cad = Cad & Text2(0).Text & "','1','" & Text2(1).Text & "',"
        InsertarLineasDeAsientos ImporteTotal, ""
    End If
    CuadrarAsiento = True
    Exit Function
ECua:
    MuestraError Err.Number, "Cuadrando asiento"
End Function



Private Function HacerElCierre() As Boolean
Dim Ok As Boolean
On Error GoTo EHacerElCierre
    
    HacerElCierre = False
    Set RS = New ADODB.Recordset
    Conn.Execute "Delete from tmpCierre"  ' no hace falta codusu pq solo puede haber trabajando uno a la vez
    
    Label10.Caption = "Leyendo datos"
    Label10.Refresh
    If Not GeneraTmpCierre Then Exit Function
    
    'Esta grabado el fichero tmpCierre con los importes
    'Fijamos la pb3
    SQL = "Select count(*) from tmpcierre where importe<>0"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroRegistros = 0
    If Not RS.EOF Then NumeroRegistros = DBLet(RS.Fields(0), "N")
    RS.Close
    If NumeroRegistros = 0 Then Exit Function
    NumeroRegistros = NumeroRegistros * 2   'Apertura y cierre
    
    MaxAsiento = 0
    Label10.Caption = "Asiento cierre"
    Label10.Refresh
    Ok = GeneraAsientoCierre
    
    If Ok Then
        Label10.Caption = "Asiento apertura"
        Label10.Refresh
        Ok = GeneraAsientoApertura
    End If
    
    If Ok Then
        'actualizaasientos
        Ok = ActualizarAsientoCierreApertura
    End If
    
    
    
    'Si no se ha generado el asiento de cierre y o apertura tenemos k borrarlo
    If Not Ok Then
        'BORRAMOS CIERRE
        SQL = "DELETE FROM cabapu where fechaent = '" & Format(vParam.fechafin, FormatoFecha) & "' AND Numasien = " & Text1(4).Text
        Conn.Execute SQL
        SQL = "DELETE FROM linapu where fechaent = '" & Format(vParam.fechafin, FormatoFecha) & "' AND Numasien = " & Text1(4).Text
        Conn.Execute SQL

        'BORRAMOS APERTURA
        SQL = "DELETE FROM cabapu where fechaent = '" & Format(Text1(9).Text, FormatoFecha) & "' AND Numasien = 1"
        Conn.Execute SQL
        SQL = "DELETE FROM linapu where fechaent = '" & Format(Text1(9).Text, FormatoFecha) & "' AND Numasien = 1"
        Conn.Execute SQL
        
    Else
        Me.cmdCierreEjercicio.Enabled = False
        
        
        
        'Ahora insertamos los contadores en hco
        Cad = "INSERT INTO contadoreshco (anoregis, tiporegi, nomregis, contado1, contado2) select " & Year(vParam.fechaini) & ",tiporegi,nomregis,contado1,contado2 FROM contadores"
        Conn.Execute Cad
        
        
            
        'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
        Cad = Format(DateAdd("yyyy", 1, vParam.fechaini), FormatoFecha)
        SQL = "UPDATE parametros SET fechaini= '" & Cad
        
        vParam.fechafin = DateAdd("yyyy", 1, vParam.fechafin)
        vParam.FechaActiva = DateAdd("yyyy", 1, vParam.FechaActiva)
        If vParam.FechaActiva >= vParam.fechafin Then vParam.FechaActiva = vParam.fechaini
        If vParam.FechaActiva < DateAdd("yyyy", 1, vParam.fechaini) Then vParam.FechaActiva = DateAdd("yyyy", 1, vParam.fechaini)
        Cad = Format(vParam.fechafin, FormatoFecha)
        SQL = SQL & "' , fechafin='" & Cad & "'"
        Cad = Format(vParam.FechaActiva, FormatoFecha)
        SQL = SQL & " , fechaactiva='" & Cad & "'"
        
        SQL = SQL & " WHERE fechaini='" & Format(vParam.fechaini, FormatoFecha) & "'"
        
        'ANTES
        'Conn.Execute SQL
        If Not EjecutaSQL(SQL) Then MsgBox "Se ha producido un error insertanto contadores en HCO." & vbCrLf & "Cuando finalice avise a soporte técnico de Ariadna Software.", vbExclamation
            
        vParam.fechaini = DateAdd("yyyy", 1, vParam.fechaini)
       
        
        'los contadores
        'UPDATEAMOS LOS CONTADORES
        'con los nuevos valores
        'Es decir Contsiguiente pasa a actual, y en siguiente ponemos un 2, puesto k el 1 lo reservamos para apertura
        SQL = "UPDATE contadores SET contado1 =  contado2"
        Conn.Execute SQL
        
        'Ponemos en todos un 0
        SQL = "UPDATE contadores SET contado2 = 0"
        Conn.Execute SQL
        
        'Menos en asientos k podremos un 1, ya que se reservara para el cierre
        'del año siguiente
        SQL = "UPDATE contadores SET contado2 = 1 WHERE tiporegi='0'"
        Conn.Execute SQL
        
        
        'NUEVO. Para que los proveedores no se mezclen facturas. El contador
        If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
            I = Year(DateAdd("yyyy", 1, vParam.fechafin))
            If I > 0 Then
                If I > 2010 Then
                    SQL = "3"
                Else
                    SQL = "2"
                End If
                SQL = SQL & Mid(CStr(I), 4, 1)
                SQL = SQL & "00000"
                
                SQL = "UPDATE contadores SET contado2 = " & SQL & " WHERE tiporegi='1'"
                If Not EjecutaSQL(SQL) Then MsgBox "Error updateando contador proveeedores: " & vbCrLf & SQL, vbExclamation
            End If
        End If
    End If
    
    HacerElCierre = True
    
EHacerElCierre:
    If Err.Number <> 0 Then MuestraError Err.Number
    vParam.Leer
    Set RS = Nothing
End Function

Private Function SimulaCierreApertura() As Boolean


    SimulaCierreApertura = False
    Set RS = New ADODB.Recordset
    Conn.Execute "Delete from tmpCierre"  ' no hace falta codusu pq solo puede haber trabajando uno a la vez
    
    Label10.Caption = "Leyendo datos"
    Label10.Refresh
    If Not GeneraTmpCierre Then Exit Function
    
    'Esta grabado el fichero tmpCierre con los importes
    'Fijamos la pb3
    SQL = "Select count(*) from tmpcierre where importe<>0"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroRegistros = 0
    If Not RS.EOF Then NumeroRegistros = DBLet(RS.Fields(0), "N")
    RS.Close
    If NumeroRegistros = 0 Then Exit Function
    NumeroRegistros = NumeroRegistros * 2   'Apertura y cierre
    
    MaxAsiento = 0
    Label10.Caption = "Asiento cierre"
    Label10.Refresh
    
    CadenaLINAPU txtDiario(1).Text, vParam.fechafin, Text1(4).Text
    If Not GeneraLineasSimulacionCierre(True) Then Exit Function

    

    CadenaLINAPU txtDiario(2).Text, CDate(Text1(9).Text), Text1(6).Text
    If Not GeneraLineasSimulacionCierre(False) Then Exit Function
    SimulaCierreApertura = True

End Function



Private Function GeneraTmpCierre() As Boolean
Dim Importe As Currency
Dim vSQL As String
Dim B As Boolean
On Error GoTo EGeneraTmpCierre


    GeneraTmpCierre = False
    'ATENCION AÑOS PARTIDOS
    vSQL = " codmacta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
    
    
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Años son el mimo, luego en para mer, luego
         Cad = vSQL & " AND anopsald = " & Year(vParam.fechafin)
    Else
        'Fecha inicio y fin no estan en el mismo año natural
        Cad = "(" & vSQL & " AND anopsald = " & Year(vParam.fechaini) & " AND mespsald >=" & Month(vParam.fechaini) & ") OR "
        Cad = Cad & "(" & vSQL & " AND anopsald = " & Year(vParam.fechafin) & " AND mespsald <=" & Month(vParam.fechafin) & ")"
    End If
    
    Cad = " from hsaldos where " & Cad & " GROUP BY codmacta ORDER BY codmacta"
    
    'Contador
    SQL = "Select -1 * (sum(impmesde)-sum(impmesha)),codmacta " & Cad
    SQL = "INSERT INTO tmpCierre " & SQL
    Conn.Execute SQL
    
    
    'Y si es simulacion, enonces borramos las cuentas de perdidas ganancias
    If Opcion = 4 Then
        SQL = "Delete from tmpcierre where cta like '" & vParam.grupogto & "%'"
        SQL = SQL & " OR  cta like '" & vParam.grupovta & "%'"
        Conn.Execute SQL
        
            
        If vParam.grupoord <> "" Then
            SQL = "Delete from tmpcierre where cta like '" & vParam.grupoord & "%'"
            'Excepcion
            If Text1(10).Text <> "" Then SQL = SQL & " AND not (cta like '" & Text1(10).Text & "%')"
            Conn.Execute SQL
        End If
            
        'Si es gran empresa me cargo tb las 8% y 9%
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            SQL = "Delete from tmpcierre where cta like '8%' OR  cta like '9%'"
            Conn.Execute SQL
        End If
        
        'Comprobamos si existe el parametro
        Set miRsAux = New ADODB.Recordset
        SQL = "Select importe from tmpcierre where cta ='" & Text1(1).Text & "'"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        Importe = 0
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then
                SQL = "E"
                Importe = miRsAux.Fields(0)
            End If
        End If
        miRsAux.Close
        
        
        If SQL = "" Then
            'Metemos a la cta 129 de perdias y gancias las perdidas y gancias generadas
            SQL = "INSERT INTO tmpcierre (cta,importe) values ('" & Text1(1).Text & "'," & TransformaComasPuntos(CStr(ImportePyG)) & ")"
        Else
            ImportePyG = ImportePyG + Importe
            SQL = "UPDATE tmpcierre SET Importe= " & TransformaComasPuntos(CStr(ImportePyG)) & " WHERE cta='" & Text1(1).Text & "'"
        End If
        Conn.Execute SQL
        
        
        
        
        'Si es gran empresa, puede que haya saldado las cuentas 8 y 9.
        'Para ello haremos lo siguiente.
        'Iremos a zhistoapu y cogeremos los apuntes que esten relacionados
        'con la regularizacion8y9, es decir, o los que en numdocum pone un 2
        ' y/o numero asiento es el indicado en el txtbox de la regularizacion
        '
        'Agruparemos por cta, sum (importe), para saber cual sera el importe resultante
        'Updatearemos (o crearemos) la linea de apunte , al igual que con la 129
        If vParam.GranEmpresa Then
            Dim L As Collection
            
            SQL = "select codusu,codmacta,sum(timporteD) as d,sum(timporteH) as h from Usuarios.zhistoapu "
            SQL = SQL & " where codusu = " & vUsu.Codigo & " and numdocum=2 and codmacta <'8' group by 1,2"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Set L = New Collection
            While Not miRsAux.EOF
                
                Importe = DBLet(miRsAux!d, "N")
                If Not IsNull(miRsAux!H) Then Importe = Importe - miRsAux!H
                SQL = miRsAux!codmacta & "|" & CStr(Importe) & "|"
                L.Add SQL
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            'Ahora ya tengo todas las cuentas a actualizar/crear
            For I = 1 To L.Count
                Cad = RecuperaValor(L.Item(I), 1)
                ImporteTotal = CCur(RecuperaValor(L.Item(I), 2))
                SQL = "Select importe from tmpcierre where cta ='" & Cad & "'"
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = ""
                Importe = 0
                If Not miRsAux.EOF Then
                    If Not IsNull(miRsAux.Fields(0)) Then
                        SQL = "E"
                        Importe = miRsAux.Fields(0)
                    End If
                End If
                miRsAux.Close
                
                
                If SQL = "" Then
                    'Metemos a la cta 129 de perdias y gancias las perdidas y gancias generadas
                    SQL = "INSERT INTO tmpcierre (cta,importe) values ('" & Cad & "'," & TransformaComasPuntos(CStr(ImportePyG)) & ")"
                Else
                    ImporteTotal = ImporteTotal + Importe
                    SQL = "UPDATE tmpcierre SET Importe= " & TransformaComasPuntos(CStr(ImporteTotal)) & " WHERE cta='" & Cad & "'"
                End If
                Conn.Execute SQL
            Next I
        End If  'granempresa
        
        Set miRsAux = Nothing
    End If
    
    
    
    GeneraTmpCierre = True
    Exit Function
EGeneraTmpCierre:
    MuestraError Err.Number, "Genera TmpCierre"
End Function




Private Function GeneraAsientoCierre() As Boolean
    On Error GoTo EGeneraAsientoCierre

    GeneraAsientoCierre = False

    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES (" & txtDiario(1).Text
    Cad = SQL & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(4).Text & ",0,NULL,NULL)"
    Conn.Execute Cad

    
    CadenaLINAPU txtDiario(1).Text, vParam.fechafin, Text1(4).Text
    If Not GeneraLineasCierre Then Exit Function

    
    GeneraAsientoCierre = True
    Exit Function
EGeneraAsientoCierre:
    MuestraError Err.Number
End Function



Private Function GeneraLineasCierre() As Boolean
Dim Cont As Integer
Dim Importe As Currency
Dim Aux As String
    On Error GoTo EGeneraLineasCierre
    GeneraLineasCierre = False
    RS.Open "SELECT * from tmpCierre WHERE importe <>0 ORDER By Cta", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cont = 1
    
    ' linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
    ' codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada
    
    
    While Not RS.EOF
        Cad = SQL & "," & Cont & ",'" & RS.Fields(1) & "','',980,'" & Text2(2).Text & "',"
        Importe = RS.Fields(0)
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        Cad = Cad & Aux
        Cad = Cad & ",NULL,'contab',0)"
        
        I = Int(((Cont + MaxAsiento) / NumeroRegistros) * pb3.Max)
        pb3.Value = I
        
        Conn.Execute Cad
        
        'Siguiente
        Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.Close
    GeneraLineasCierre = True
    MaxAsiento = Cont - 1
    Exit Function
EGeneraLineasCierre:
    MuestraError Err.Number
End Function



Private Function GeneraAsientoApertura() As Boolean
    On Error GoTo EGeneraAsientoApertura

    GeneraAsientoApertura = False

    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES (" & txtDiario(2).Text
    Cad = SQL & ",'" & Format(Text1(9).Text, FormatoFecha) & "'," & Text1(6).Text & ",0,NULL,NULL)"
    Conn.Execute Cad
    
    CadenaLINAPU txtDiario(2).Text, CDate(Text1(9).Text), Text1(6).Text
    If Not GeneraLineasApertura Then Exit Function

    GeneraAsientoApertura = True
    Exit Function
EGeneraAsientoApertura:
    MuestraError Err.Number
End Function



Private Function GeneraLineasApertura() As Boolean
Dim Cont As Integer
Dim Importe As Currency
Dim Aux As String
    On Error GoTo EGeneraLineasApertura
    GeneraLineasApertura = False
    RS.Open "SELECT * from tmpCierre WHERE importe <>0 ORDER BY Cta ", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cont = 1
    
    ' linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
    ' codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada
    
    
    While Not RS.EOF
        Cad = SQL & "," & Cont & ",'" & RS.Fields(1) & "','',970,'" & Text2(3).Text & "',"
        Importe = RS.Fields(0)
        Importe = Importe * -1
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        Cad = Cad & Aux
        Cad = Cad & ",NULL,'contab',0)"
        
        I = Int(((Cont + MaxAsiento) / NumeroRegistros) * pb3.Max)
        pb3.Value = I
        
        Conn.Execute Cad
        
        'Siguiente
        Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.Close
    GeneraLineasApertura = True
    Exit Function
EGeneraLineasApertura:
    MuestraError Err.Number
End Function


Private Function ActualizarAsientoCierreApertura() As Boolean
Screen.MousePointer = vbHourglass

            ActualizarAsientoCierreApertura = False
            
            
            'CIERRE
            frmActualizar.OpcionActualizar = 1
            frmActualizar.NumAsiento = CLng(Text1(4).Text)
            frmActualizar.FechaAsiento = vParam.fechafin
            frmActualizar.NumDiari = CInt(txtDiario(1).Text)
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            Me.Refresh
            If Not AlgunAsientoActualizado Then Exit Function
            

            'Apertura
            frmActualizar.OpcionActualizar = 1
            frmActualizar.NumAsiento = CLng(Text1(6).Text)
            frmActualizar.FechaAsiento = CDate(Text1(9).Text)
            frmActualizar.NumDiari = CInt(txtDiario(2).Text)
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            Me.Refresh
            If Not AlgunAsientoActualizado Then Exit Function

            Me.Refresh
            ActualizarAsientoCierreApertura = True
End Function

'Comprobamos k ha fecha fin ejercicio anterior, no hay cierre
Private Function NoHayCierre(FechaCierre As Date) As Boolean

    On Error GoTo ENoHayCierre
    NoHayCierre = True
    
    'SQL = "Select numasien from hlinapu WHERE codconce=980 AND fechaent='" & Format(Label13.Caption, FormatoFecha) & "'"
    SQL = "Select numasien from hlinapu WHERE codconce=980 AND fechaent='" & Format(FechaCierre, FormatoFecha) & "'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then NoHayCierre = False
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
ENoHayCierre:
    MuestraError Err.Number, "Comprobar cierre anterior"
End Function

'Para el traspaso de las tablas hcabapu, hlinapu,.saldos y slados anal
Private Function YaHayDatosFechas(FechaCierre As Date) As Boolean
Dim F1 As Date
    YaHayDatosFechas = False
    
    Set RS = New ADODB.Recordset
    '1º Obtenemos la fecha mas baja del ejercicio a traspasar
    'SQL = "Select fechaent from hcabapu where fechaent<= '" & Format(Label13.Caption, FormatoFecha) & "' ORDER BY fechaent ASC"
    SQL = "Select min(fechaent) from hcabapu where fechaent<= '" & Format(FechaCierre, FormatoFecha) & "'"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = "01/01/2099"
    If Not RS.EOF Then Cad = Format(RS.Fields(0), "dd/mm/yyyy")
    RS.Close
    
    SQL = "Select numasien from hcabapu1 WHERE fechaent >='" & Format(Cad, FormatoFecha) & "'"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then YaHayDatosFechas = True
    End If
    RS.Close
    
    
    If Not YaHayDatosFechas Then
        SQL = "Select codmacta from hsaldos1 where "
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            SQL = SQL & " (anopsald > " & Year(CDate(Cad)) & ") "
        Else
            F1 = CDate(Cad)
            SQL = SQL & " ((anopsald = " & Year(F1) & "  and mespsald >=" & Month(CDate(F1)) & " ) OR ("
            F1 = DateAdd("yyyy", 1, F1)
            F1 = DateAdd("d", -1, F1)
            SQL = SQL & " anopsald = " & Year(F1) & "  and mespsald <" & Month(CDate(F1)) & "))"
        End If
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then YaHayDatosFechas = True
        End If
        RS.Close
    End If
         
         
          
    If Not YaHayDatosFechas Then
        
        SQL = "Select codmacta from hsaldosanal1 where "
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            SQL = SQL & " (anoccost > " & Year(CDate(Cad)) & ") "
        Else
            F1 = CDate(Cad)
            SQL = SQL & " ((anoccost = " & Year(F1) & "  and mesccost >=" & Month(CDate(F1)) & " ) OR ("
            F1 = DateAdd("yyyy", 1, F1)
            F1 = DateAdd("d", -1, F1)
            SQL = SQL & " anoccost = " & Year(F1) & "  and mesccost <" & Month(CDate(F1)) & "))"
        End If
        
        
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then YaHayDatosFechas = True
        End If
        RS.Close
    End If
         
    Set RS = Nothing

End Function


Private Function HacerTraspaso(FechaCierre As Date) As Boolean
    HacerTraspaso = False
    Label15.Caption = "Cabecera apuntes"
    Label15.Refresh
    If Not Traspaso(1, FechaCierre) Then Exit Function
    
    Label15.Caption = "Lineas apuntes"
    Label15.Refresh
    If Not Traspaso(2, FechaCierre) Then Exit Function
    
    Label15.Caption = "Saldos"
    Label15.Refresh
    If Not Traspaso(3, FechaCierre) Then Exit Function
    
    Label15.Caption = "Saldos analítica"
    Label15.Refresh
    If Not Traspaso(4, FechaCierre) Then Exit Function
    HacerTraspaso = True
End Function


Private Function Traspaso(Num As Integer, FechaCierre As Date) As Boolean

    On Error GoTo ETraspaso
    Traspaso = False
    Select Case Num
    Case 1
        'Cabecera de apuntes
        Cad = "fechaent<= '" & Format(FechaCierre, FormatoFecha) & "'"
        SQL = "INSERT INTO hcabapu1 SELECT * from hcabapu WHERE " & Cad
        Conn.Execute SQL
        SQL = "DELETE FROM hcabapu WHERE " & Cad
        Conn.Execute SQL
        
        
    Case 2
        'Lineas de apuntes
        Cad = "fechaent<= '" & Format(FechaCierre, FormatoFecha) & "'"
        SQL = "INSERT INTO hlinapu1 SELECT * from hlinapu WHERE " & Cad
        Conn.Execute SQL
        SQL = "DELETE FROM hlinapu WHERE " & Cad
        Conn.Execute SQL
    
    Case 3
        'Hsaldos
        SQL = FechaCierre
        Cad = " (anopsald < " & Year(CDate(SQL)) & ") OR "
        Cad = Cad & " (mespsald <=" & Month(CDate(SQL)) & " AND anopsald =" & Year(CDate(SQL)) & ")"
        
        SQL = "INSERT INTO hsaldos1 SELECT * from hsaldos WHERE "
        SQL = SQL & Cad
        Conn.Execute SQL
        
        'Borramos
        SQL = "DELETE FROM hsaldos WHERE " & Cad
        Conn.Execute SQL
        
        
    Case 4
        'SALDOS ANALITICA
        SQL = FechaCierre
        Cad = " (anoccost < " & Year(CDate(SQL)) & ") OR "
        Cad = Cad & " (mesccost <=" & Month(CDate(SQL)) & " AND anoccost =" & Year(CDate(SQL)) & ")"
        
        SQL = "INSERT INTO hsaldosanal1 SELECT * from hsaldosanal WHERE "
        SQL = SQL & Cad
        Conn.Execute SQL
        
        'Borramos
        SQL = "DELETE FROM hsaldosanal WHERE " & Cad
        Conn.Execute SQL
    
    End Select
    Traspaso = True
    Exit Function
ETraspaso:
    MuestraError Err.Number, "Trasapasando datos(" & Num & ")"
End Function




Private Function GeneraLineasSimulacionCierre(EsCierre As Boolean) As Boolean
Dim Cont As Integer
Dim Importe As Currency
Dim Aux As String
    On Error GoTo EGeneraLineasSimulacionCierre
    GeneraLineasSimulacionCierre = False
    Cad = "select tmpcierre.*,nommacta from tmpcierre,cuentas where tmpcierre.cta=cuentas.codmacta"
    Cad = Cad & " AND importe <>0 ORDER By Cta"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Cont = 1
    Cont = 2
    'linliapu, codmacta, nommacta, numdocum, ampconce, timporteD, "
    'codccost,timporteH) VALUES ("
    
    While Not RS.EOF
        
        Importe = RS.Fields(0)
        If EsCierre Then
            Cad = SQL & "," & Cont & ",'" & RS.Fields(1) & "','" & DevNombreSQL(RS!nommacta) & "','3','" & Text2(2).Text & "',"
        Else
            Cad = SQL & "," & Cont & ",'" & RS.Fields(1) & "','" & DevNombreSQL(RS!nommacta) & "','4','" & Text2(3).Text & "',"
            Importe = Importe * -1
        End If
        
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        Cad = Cad & Aux & ")"
    
        I = Int(((Cont + MaxAsiento) / (NumeroRegistros + 3)) * pb3.Max)
        pb3.Value = I
        
        Conn.Execute Cad
        
        'Siguiente
        Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.Close
    GeneraLineasSimulacionCierre = True
    MaxAsiento = Cont - 1
    Exit Function
EGeneraLineasSimulacionCierre:
    MuestraError Err.Number
End Function



Private Function ExisteAsientosDescerrar() As Boolean

'Tiene k existe el asiento de apertura del año siguient
    Cad = CadenaFechasActuralSiguiente(True)
    SQL = "Select numasien from hlinapu WHERE codconce=970"
    SQL = SQL & " AND " & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ExisteAsientosDescerrar = True
    Else
        ExisteAsientosDescerrar = False
    End If
    RS.Close
    Set RS = Nothing
End Function


Private Function ExistenApuntesEjercicioAnterior() As Boolean

'Tiene k existe el asiento de apertura del año siguient
    SQL = "Select count(*) from hlinapu WHERE "
    SQL = SQL & " Fechaent < '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ExistenApuntesEjercicioAnterior = False
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then ExistenApuntesEjercicioAnterior = True
    End If
    RS.Close
    Set RS = Nothing
End Function


Private Function ExistenApuntesEjercicioSiguiente() As Boolean

'Tiene k existe el asiento de apertura del año siguient
    SQL = "Select count(*) from hlinapu WHERE "
    SQL = SQL & " Fechaent > '" & Format(vParam.fechafin, FormatoFecha) & "'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ExistenApuntesEjercicioSiguiente = False
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then ExistenApuntesEjercicioSiguiente = True
    End If
    RS.Close
    Set RS = Nothing
End Function




Private Function HacerDescierre() As Boolean
Dim N As Long
Dim MaxAsien As Long

On Error GoTo EHacerDescierre
    HacerDescierre = False
    Screen.MousePointer = vbHourglass
    
    'Si ya hay un 960 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    'P y G
    MaxAsien = 0
    Label19.Caption = "Perdidas y ganancias"
    Label19.Refresh
    SQL = "Select numasien,fechaent,numdiari from hlinapu WHERE codconce=960 "
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "'"
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        Cad = RS!fechaent
        I = RS!NumDiari
        NumeroRegistros = RS!Numasien
        MaxAsien = NumeroRegistros
    End If
    RS.Close
    If Cad <> "" Then
        EliminarAsiento
    End If
    Me.Refresh
    espera 0.5
    Me.Refresh


    'Compruebo si hay REGULARIZACON 8 y 9
    If vParam.GranEmpresa Then
        Label19.Caption = "Regularización 8 y 9"
        Label19.Refresh
    End If
    SQL = "Select numasien,fechaent,numdiari from hlinapu WHERE codconce=961 "
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "'"
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        Cad = RS!fechaent
        I = RS!NumDiari
        NumeroRegistros = RS!Numasien
    End If
    RS.Close
    If Cad <> "" Then
        EliminarAsiento
    End If
    Me.Refresh
    espera 0.5
    Me.Refresh



    'Cierre
    'Si hay un 980  en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Label19.Caption = "Cierre"
    Label19.Refresh
    SQL = "Select numasien,fechaent,numdiari from hlinapu WHERE codconce=980"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        Cad = RS!fechaent
        I = RS!NumDiari
        NumeroRegistros = RS!Numasien
        If MaxAsien = 0 Then MaxAsien = NumeroRegistros
            
    End If
    RS.Close
    If Cad <> "" Then
        EliminarAsiento
        If Not AlgunAsientoActualizado Then Exit Function  'Error en ekiminar asiento
    End If
    Me.Refresh
    espera 0.5
    Me.Refresh
        

    'Si ya hay un  970 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Label19.Caption = "Apertura"
    Label19.Refresh
    SQL = "Select numasien,fechaent,numdiari from hlinapu WHERE codconce=970"
    Cad = Format(vParam.fechaini, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = Format(vParam.fechafin, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        Cad = RS!fechaent
        I = RS!NumDiari
        NumeroRegistros = RS!Numasien
    End If
    RS.Close
    If Cad <> "" Then
        EliminarAsiento
        If Not AlgunAsientoActualizado Then Exit Function  'Error en ekiminar asiento
    End If
    Me.Refresh
    espera 0.5
    Me.Refresh


    'Hay k bajar una año las fechas de parametros de INICIO y FIN ejerecicio
    'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
    'Reestablecemos las fechas
    'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
    Label19.Caption = "Contadores"
    Label19.Refresh
    I = -1
    Cad = Format(DateAdd("yyyy", I, vParam.fechaini), FormatoFecha)
    SQL = "UPDATE parametros SET fechaini= '" & Cad
    Cad = Format(DateAdd("yyyy", I, vParam.fechafin), FormatoFecha)
    SQL = SQL & "' , fechafin='" & Cad & "'"
    'Fechaactiva
    Cad = Format(DateAdd("yyyy", I, vParam.FechaActiva), FormatoFecha)
    SQL = SQL & " , FechaActiva='" & Cad & "'"
    
    SQL = SQL & " WHERE fechaini='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Conn.Execute SQL
    
    vParam.fechaini = DateAdd("yyyy", I, vParam.fechaini)
    vParam.fechafin = DateAdd("yyyy", I, vParam.fechafin)
    'Fecha activa
    vParam.FechaActiva = DateAdd("yyyy", I, vParam.FechaActiva)



    'No podemos borrar los contadores, entonces para cada
    'Reestablecemos los contadores
    'por lo tanto tendremos k UPDATEAR, o CREAR
    Set RS = New ADODB.Recordset
    SQL = "SELECT tiporegi, nomregis, contado1, contado2 from Contadoreshco WHERE anoregis = " & Year(vParam.fechaini)
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    While Not RS.EOF

        Cad = DevuelveDesdeBD("tiporegi", "Contadores", "tiporegi", RS.Fields(0), "T")
        If Cad = "" Then
            SQL = "INSERT INTO contadores (tiporegi, nomregis, contado1, contado2) VALUES ('" & RS.Fields(0)
            SQL = SQL & "','" & RS.Fields(1) & "'," & RS.Fields(2) & "," & RS.Fields(3) & ")"
        Else
                
            'SQL = "UPDATE Contadores SET contado1 = " & RS.Fields(2) & " , Contado2 = " & RS.Fields(3)
            SQL = "UPDATE Contadores SET contado2 = contado1 , Contado1 = " & RS.Fields(2)
            SQL = SQL & " WHERE tiporegi ='" & RS.Fields(0) & "'"
        End If
        EjecutaSQL SQL
        
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    'El contador de ejercicio actual es, MaxAsien
    If MaxAsien <> 0 Then
        'ESTO NO ES ASI Y NO LO VOY A TOCAR
        'SQL = "UPDATE Contadores SET contado1 = " & MaxAsien - 1 & " WHERE tiporegi ='0'"
        'EjecutaSQL SQL
    End If
    'Los borramos de hco
    SQL = "DELETE from Contadoreshco WHERE anoregis = " & Year(vParam.fechaini)
    EjecutaSQL SQL
    
    HacerDescierre = True
    Exit Function
EHacerDescierre:
    MuestraError Err.Number, "Proc. HacerDescierre"
End Function




Private Sub EliminarAsiento()
        Screen.MousePointer = vbHourglass
        frmActualizar.OpcionActualizar = 2  'Desactualizar para eliminar
        'frmActualizar.NUmSerie = C
        frmActualizar.NumAsiento = NumeroRegistros
        frmActualizar.FechaAsiento = CDate(Cad)
        frmActualizar.NumDiari = I
        frmActualizar.NUmSerie = ""
        AlgunAsientoActualizado = False
        frmActualizar.Show vbModal
End Sub


Private Function EliminarEjerciciosCerrados(F As Date) As Boolean

    On Error GoTo EEliminarEjerciciosCerrados
    EliminarEjerciciosCerrados = False
    
    Cad = "fechaent <='" & Format(F, FormatoFecha) & "'"
    
        
    'Eliminamos del hlinapu
    Label20.Caption = "Lineas apuntes 1"
    SQL = "DELETE from hlinapu1 where " & Cad
    Me.Refresh
    Conn.Execute SQL
    
    'Cabecera
    Label20.Caption = "Cabecera apuntes 1"
    SQL = "DELETE from hcabapu1 where " & Cad
    Me.Refresh
    Conn.Execute SQL
    
    
    'Hsaldos
    If Month(F) = 12 And Day(F) = 31 Then
        'Fin de año
        'Con solo poner año <= sirve
        I = 1
    Else
        I = 0
    End If
    
    
    
    'Saldos
    Label20.Caption = "Saldos 1"
    If I = 1 Then
        'Ultimo de año 31/12 del algo
        Cad = "anopsald<=" & Year(F)
    Else
        Cad = "(anopsald<" & Year(F)
        Cad = Cad & ") OR (anopsald =" & Year(F)
        Cad = Cad & " AND mespsald <=" & Month(F) & ")"
    End If
    SQL = "DELETE from hsaldos1 where " & Cad
    Conn.Execute SQL
    
    
    'Analitica
    Label20.Caption = "Saldos analitica 1"
    If I = 1 Then
        'Ultimo de año 31/12 del algo
        Cad = "anoccost<=" & Year(F)
    Else
        Cad = "(anoccost<" & Year(F)
        Cad = Cad & ") OR (anoccost =" & Year(F)
        Cad = Cad & " AND mesccost <=" & Month(F) & ")"
    End If
    
    SQL = "DELETE from hsaldosanal1 where " & Cad
    Conn.Execute SQL
    
    
    
    
    EliminarEjerciciosCerrados = True
    Exit Function
EEliminarEjerciciosCerrados:
    MuestraError Err.Number, Err.Description
End Function


Private Sub CargaComboTraspaso()
    Dim F As Date
    Dim F2 As Date
    

    cmbCierre.Clear
    Cad = "Select min(fechaent) from  hlinapu"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            F = RS.Fields(0)
            I = 1
        End If
    End If
    RS.Close
    Set RS = Nothing
    If I = 0 Then Exit Sub
    If F >= vParam.fechaini Then Exit Sub
    F2 = vParam.fechaini
    I = 1
    Do
        SQL = Format(DateAdd("d", -1, F2), "dd/mm/yyyy")
        I = Year(CDate(SQL))
        SQL = "Cierre: " & SQL
        F2 = DateAdd("yyyy", -1, F2)

        cmbCierre.AddItem SQL
        cmbCierre.ItemData(cmbCierre.NewIndex) = I
    Loop Until F2 <= F
    
End Sub


Private Sub PonerGrandesEmpresas()
Dim B As Boolean

    B = False
    If vParam.NuevoPlanContable Then B = vParam.GranEmpresa
    'Cuando sean autmocion en plan nuevo ya veremos como resolvemos esto
    'FALTA###
    
    
    GElabel1.visible = B
    Me.GELine.visible = B
    Label5(17).visible = B
    Label5(16).visible = B
    Label5(16).visible = B
    Label5(15).visible = B
    Label5(14).visible = False 'vEmpresa.GranEmpresa
    Label5(18).visible = B
    Text2(4).visible = B
    Me.Image1(3).visible = B
    Me.txtDiario(3).visible = B
    Me.txtDescDiario(3).visible = B
    Text1(13).visible = B
    Text1(14).visible = False 'vEmpresa.GranEmpresa
    Text1(11).visible = B
    Text1(12).visible = B
    
    
    If Not B Then
        'SOlo para la antigua
        '---------------------------------------------------------
        'Si tiene el otro grupo de perdidas y ganacias entonces
        'tenemos k solicitar la excepcion a digitos de tercer nivel
        'ofertando el de parametros
        B = (vParam.grupoord <> "") And (vParam.Automocion <> "")
        Label5(13).visible = B
        Text1(10).visible = B
        If B Then Text1(10).Text = vParam.Automocion
    End If
    
End Sub



Private Function ComprobarCierreCuentas8y9() As Boolean


    On Error GoTo EComprobarCierreCuentas8y9
    ComprobarCierreCuentas8y9 = False

    Conn.Execute "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo

    Set RS = New ADODB.Recordset
    Cad = "select " & vUsu.Codigo & ",codmacta,'T',sum(impmesde)-sum(impmesha) from hsaldos"
    Cad = Cad & " WHERE (codmacta like '8__' or codmacta like '9__') AND "
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Años son el mimo, luego en para mer, luego
         Cad = Cad & " anopsald = " & Year(vParam.fechafin)
    Else
        'Fecha inicio y fin no estan en el mismo año natural
        'ejemplo: and ((anopsald=2006 and mespsald>2) and (anopsald=2006 and mespsald<=3))
        Cad = Cad & "((anopsald = " & Year(vParam.fechaini) & " AND mespsald >=" & Month(vParam.fechaini) & ") AND "
        Cad = Cad & " (anopsald = " & Year(vParam.fechafin) & " AND mespsald <=" & Month(vParam.fechafin) & "))"
        
    End If
    Cad = Cad & " GROUP BY codmacta"
    'Insertamos en tmpcierre
    Cad = "INSERT INTO TMPCIERRE1 " & Cad
    Conn.Execute Cad
    
    'COJEREMOS TODAS LAS CUENTAS 8 y 9 a tres digitos y comprobaremos que en
    'la configuracion tienen puesto en el campo cuentaba
    '
    'Cruzamos tmpcierr1 con codusu = vusu y left join con cuentas
    'Veremos si hay null con lo cual esta mal y si no, updatearemos tmpcierre
    Cad = "select cta,nommacta,cuentaba from tmpcierre1,cuentas where codusu = " & vUsu.Codigo & " and cta = codmacta"
  
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    I = 0
    While Not RS.EOF
        If DBLet(RS!cuentaba, "T") = "" Then
            
            I = I + 1
            Cad = Cad & "     " & RS!Cta
            If (I Mod 5) = 0 Then Cad = Cad & vbCrLf
        
        Else
        
            'ASi, tanto para la simulacion, como para el cierre ya se contra que cuentas saldan las del 8 9
            Conn.Execute "UPDATE tmpcierre1 SET nomcta = '" & RS!cuentaba & "' WHERE codusu = " & vUsu.Codigo & " and cta = '" & RS!Cta & "'"
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    
    If I > 0 Then
        Cad = "Cuentas sin configurar el cierre: " & vbCrLf & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Set RS = Nothing
        Exit Function
    End If
    
    
    'OK tiene todas las cuentas configuradas
    Cad = "Select tmpcierre1.nomcta, cuentas.codmacta from tmpcierre1 left join cuentas on tmpcierre1.codusu=" & vUsu.Codigo & " AND nomcta = cuentas.codmacta"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    I = 0
    While Not RS.EOF
        
        If IsNull(RS!codmacta) Then
            I = I + 1
            Cad = "    " & RS!nomcta
            If (I Mod 5) = 0 Then Cad = Cad & vbCrLf
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    
    If I > 0 Then
        Cad = "Cuentas de cierre configurada, pero no existen: " & vbCrLf & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Set RS = Nothing
        Exit Function
    End If
    
            
                
            
    
    
    ComprobarCierreCuentas8y9 = True
    Set RS = Nothing
    
    
    Exit Function
EComprobarCierreCuentas8y9:
    Set RS = Nothing
    MuestraError Err.Number, "Comprobar Cierre Cuentas 8y9"
End Function

Private Function ASiento8y9() As Boolean
Dim Ok As Byte

    ASiento8y9 = False
    
        
    Ok = GenerarASiento8y9     '0.- MAL   '1,. Bien pero sin generar apunte(para que no contabilice)    2.- Bien
    
    If Opcion = 4 Then
        ASiento8y9 = (Ok > 0)
        Exit Function
    End If
    
    
    'Llamamos a actualizar el asiento, para pasarlo a hco
    If Ok = 2 Then
            Screen.MousePointer = vbHourglass
            frmActualizar.OpcionActualizar = 1
            frmActualizar.NumAsiento = CLng(Text1(11).Text)
            frmActualizar.FechaAsiento = vParam.fechafin
            frmActualizar.NumDiari = CInt(txtDiario(3).Text)
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            Me.Refresh
            If AlgunAsientoActualizado Then
                Ok = 1
            Else
                Ok = 0
            End If
    End If
        
        
    'Si entra mal hay que borrar los apuntes que pudieran haberse creado
    If Ok = 0 Then
        'Comun lineas y cabeceras
        SQL = " where numdiari=" & txtDiario(2).Text
        SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(11).Text & ""
        
        'Borramos por si acaso ha insertado lineas
        Cad = "Delete FROM linapu" & SQL
        Conn.Execute Cad
    
        'Borramos la cabcecera del apunte
        Cad = "DELETE FROM cabapu" & SQL
        Conn.Execute Cad
        
        Label3.Caption = ""
    
    End If
    
    ASiento8y9 = Ok > 0
    

End Function


Private Function GenerarASiento8y9() As Byte
Dim CuentaSaldo As String
Dim Importe As Currency
Dim ImpComprobacion As Currency
Dim RT As ADODB.Recordset
Dim Cont As Long

    On Error GoTo EASiento8y9
    GenerarASiento8y9 = 0
    
    'Generamos los  apuntes, sobre cabapu, y luego los actualizamos

    
    'Cogemos tmpcierre1 que tendra cta, cta saldada
    'Ejemlo   codusu, cta, saldar
    '                 910   1290003
    '                 920   1290003
    '                 830   1200002
    'Entonces cogeremos los saldos para estas cuentas y las iremos saldando
    'Cargaremos RT con los saldos a 3 digitos (como no ocupara mucho... NO problemo
    Set RS = New ADODB.Recordset
    
    SQL = "Select count(*) from tmpcierre1 where codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroRegistros = 0
    
    If Not RS.EOF Then NumeroRegistros = DBLet(RS.Fields(0), "N")
    RS.Close
    If NumeroRegistros = 0 Then
    
    
        'Automaticamente el numero de registro que se le iba asignar pasa al cierre
        Text1(4).Text = Text1(11).Text
    
    
    
        'OK
        Set RS = Nothing
        GenerarASiento8y9 = 1
        Exit Function
    End If
    
    If Opcion <> 4 Then
        SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES (" & txtDiario(3).Text
        Cad = SQL & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(11).Text & ",0,NULL,NULL)"
        
        Conn.Execute Cad
    End If
    
    
    
    
    NumeroRegistros = NumeroRegistros + 1 'Para que no desborde
    
    SQL = "Select tmpcierre1.* from tmpcierre1 where  codusu = " & vUsu.Codigo & " ORDER BY nomcta,cta"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CuentaSaldo = ""
    MaxAsiento = 1

    
    
    'Preparamos el SQL para la insercion de lineas de apunte
    'Montamos la cadena casi al completo
    CadenaLINAPU txtDiario(3).Text, vParam.fechafin, Text1(11).Text

    Set RT = New ADODB.Recordset
    
    While Not RS.EOF
        If CuentaSaldo <> RS!nomcta Then
            
                'SALDAMOS
            If CuentaSaldo <> "" Then SaldarCuenta8y9 CuentaSaldo
 
            CuentaSaldo = RS!nomcta   'Cuenta saldo
            ImporteTotal = 0
        End If
    
        
        
        'Progress y label
        Label3.Caption = RS!Cta & " - " & RS!nomcta
        Label3.Refresh
        Cont = Cont + 1
        I = Int((Cont / NumeroRegistros) * pb1.Max)
        pb1.Value = I
        
        
        
        'Selecciono todas las cuentas para el subgrupo de 3 digitos
        Cad = Mid(RS!Cta & "_______", 1, vEmpresa.DigitosUltimoNivel)
        Cad = " WHERE hsaldos.codmacta=cuentas.codmacta AND hsaldos.codmacta like '" & Cad & "' AND "
        Cad = "select hsaldos.codmacta,sum(impmesde)-sum(impmesha) as miImporte,nommacta from hsaldos,cuentas" & Cad
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            'Años son el mimo, luego en para mer, luego
             Cad = Cad & " anopsald = " & Year(vParam.fechafin)
        Else
            'Fecha inicio y fin no estan en el mismo año natural
            'ejemplo: and ((anopsald=2006 and mespsald>2) and (anopsald=2006 and mespsald<=3))
            Cad = Cad & "((anopsald = " & Year(vParam.fechaini) & " AND mespsald >=" & Month(vParam.fechaini) & ") AND "
            Cad = Cad & " (anopsald = " & Year(vParam.fechafin) & " AND mespsald <=" & Month(vParam.fechafin) & "))"
            
        End If
        Cad = Cad & " GROUP BY codmacta"
        ImpComprobacion = 0
        RT.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RT.EOF
            Importe = RT!miImporte  'importe
            
            If Importe <> 0 Then
                If Opcion = 4 Then
                    ' "linliapu, codmacta, nommacta, numdocum, ampconce "
                    Cad = SQL & "," & MaxAsiento & ",'" & RT!codmacta & "','" & DevNombreSQL(RT!nommacta) & "','2','" & Text2(4).Text & "',"
                Else
                    Cad = SQL & "," & MaxAsiento & ",'" & RT!codmacta & "','',961,'" & Text2(4).Text & "',"
                End If
                InsertarLineasDeAsientos Importe, "NULL"
            End If
        
            'Sig
            ImpComprobacion = ImpComprobacion + Importe
            MaxAsiento = MaxAsiento + 1
            RT.MoveNext
        Wend
        RT.Close
        
        
        Importe = RS!acumPerD   'Para combrobar que a ultimo nivel suma igual que a 3 digitos
        
        If ImpComprobacion <> Importe Then
            Cad = "Error obteniendo saldos. " & vbCrLf & "Subgrupo: " & RS!Cta & vbCrLf
            Cad = Cad & "Imp 3 digitos: " & Importe & vbCrLf & "Ultimo nivel: " & ImpComprobacion
            If Opcion <> 4 Then Cad = Cad & vbCrLf & vbCrLf & " No puede continuar con el cierre"
            MsgBox Cad, vbExclamation
            If Opcion <> 4 Then
                RS.Close
                Exit Function
            End If
        End If
            
            
            
        'Sigueinte subgrupoo
        RS.MoveNext
    Wend
            
    RS.Close
    Set RS = Nothing
    Set RT = Nothing
    If CuentaSaldo <> "" Then SaldarCuenta8y9 CuentaSaldo
    ImporteTotal = 0
    
    
    
    

    
    
    GenerarASiento8y9 = 2

    
    
   Exit Function
EASiento8y9:
    Set RS = Nothing
    MuestraError Err.Number, Err.Description
End Function


Private Sub SaldarCuenta8y9(LaCuenta As String)
Dim C As String
    ImporteTotal = ImporteTotal * -1 'Para k cuadre
    If ImporteTotal <> 0 Then
        If Opcion = 4 Then
            C = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", LaCuenta, "T")
            ' "linliapu, codmacta, nommacta, numdocum, ampconce "
            Cad = SQL & "," & MaxAsiento & ",'" & LaCuenta & "','" & DevNombreSQL(C) & "','2','" & Text2(4).Text & "',"
        Else
            Cad = SQL & "," & MaxAsiento & ",'" & LaCuenta & "','',961,'" & Text2(4).Text & "',"
        End If
        InsertarLineasDeAsientos ImporteTotal, "NULL"
        MaxAsiento = MaxAsiento + 1
    End If
End Sub



Private Sub PonerDatosTraerCierre()
    Screen.MousePointer = vbHourglass
    Label21(0).Caption = "Obteniendo datos"
    Set RS = New ADODB.Recordset
    Me.cmdTraer.Enabled = PrepararTraerCierre
    Set RS = Nothing
End Sub


Private Function PrepararTraerCierre() As Boolean
Dim F As Date
    On Error GoTo EPrepararTraerCierre
    
    Cad = "Select max(fechaent) from hlinapu1 "
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then CadenaDesdeOtroForm = RS.Fields(0)
    End If
    RS.Close
    
    
    If CadenaDesdeOtroForm = "" Then
        Label21(0).Caption = "Ningun dato en ejercicios cerrados"
        Exit Function
    End If
    F = CDate(CadenaDesdeOtroForm)
    
    If F >= vParam.fechaini Then
        Label21(0).Caption = "Fecha mayor que fecha inicio en parametros"
        Exit Function
    End If
    
    
    'Veo el minimo en introducicion o apudirec
    Cad = "Select min(fechaent) from hlinapu "
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then SQL = RS.Fields(0)
    End If
    RS.Close
    
    If SQL = "" Then SQL = "01/01/2100"
    F = CDate(SQL)
    Cad = "Select min(fechaent) from linapu "
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then If RS.Fields(0) < F Then F = RS.Fields(0)
    End If
    RS.Close
    
    
    
    If F < CDate(CadenaDesdeOtroForm) Then
        Label21(0).Caption = "Fecha en cerrados mayor que fecha en apuntes"
        Exit Function
    End If
    
    'La ultima fecha (en dia/mes) debe ser igual a la de cierre ejercicio
    If Format(CDate(CadenaDesdeOtroForm), "ddmm") <> Format(vParam.fechafin, "ddmm") Then
        Label21(0).Caption = "No coincide fecha cierre(mes/dia) con la de fin de ejercicios"
        Exit Function
    End If
    
    
    'Si llega aqui dejaremos que traiga de cerrados
    Label21(0).Caption = "Ultimo cierre en cerrados: " & CadenaDesdeOtroForm
    Label21(1).Tag = CadenaDesdeOtroForm
    PrepararTraerCierre = True
    
    Exit Function
EPrepararTraerCierre:
    MuestraError Err.Number
End Function



Private Function TraerDeCerrados() As Boolean
Dim F1 As Date
Dim F2 As Date
    
    On Error GoTo ETraerDeCerrados
    TraerDeCerrados = False
    CadenaDesdeOtroForm = Label21(1).Tag
    
    F2 = CDate(CadenaDesdeOtroForm)
    F1 = vParam.fechaini
    While F1 > F2
        F1 = DateAdd("yyyy", -1, F1)
    Wend
    Label21(0).Tag = "Fechas: " & F1 & "  -   " & F2 & vbCrLf & vbCrLf
        
    'Traer saldos
    I = 0
    Label21(0).Caption = Label21(0).Tag & "Saldos  (1/4)"
    Cad = "hsaldos"
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Año inicio igual a fin Ejercicios naturales
        SQL = "  from hsaldos1 where anopsald = " & Year(F1)
        AccionesTraer
        
    Else
        SQL = " from hsaldos1 where anopsald = " & Year(F1) & " and mespsald >= " & Month(F1)
        AccionesTraer
        
        SQL = " from hsaldos1 where anopsald = " & Year(F2) & " and mespsald < " & Month(F2)
        AccionesTraer
    End If
    
    
    
    'ANALITICA
    Label21(0).Caption = Label21(0).Tag & "Analitica  (2/4)"
    If Not vParam.autocoste Then
        espera 1.5
    Else
        Cad = "hsaldosanal"
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            'Año inicio igual a fin Ejercicios naturales
            SQL = "  from hsaldosanal1 where anoccost = " & Year(F1)
            AccionesTraer
        Else
            SQL = "  from hsaldos1 where anopsald = " & Year(F1) & " and mespsald >= " & Month(F1)
            AccionesTraer
    
            SQL = "  from hsaldos1 where anopsald = " & Year(F2) & " and mespsald < " & Month(F2)
            AccionesTraer
        End If
    End If
    
    
    'CABAPU
    Label21(0).Caption = Label21(0).Tag & "Apuntes C.  (3/4)"
    SQL = "  from hcabapu1 where fechaent >= '" & Format(F1, FormatoFecha) & "' AND fechaent <= '" & Format(F2, FormatoFecha) & "'"
    Cad = "hcabapu"
    AccionesTraer
    
    'LINAPU
    Label21(0).Caption = Label21(0).Tag & "Apuntes C.  (4/4)"
    SQL = "  from hlinapu1 where fechaent >= '" & Format(F1, FormatoFecha) & "' AND fechaent <= '" & Format(F2, FormatoFecha) & "'"
    Cad = "hlinapu"
    AccionesTraer
    
    
    TraerDeCerrados = True
    
    Exit Function
ETraerDeCerrados:
    MuestraError Err.Number
End Function



Private Sub AccionesTraer()
Dim C As String
    'Para cada accion llevara un INSERT y UN DELETE
    I = I + 1
    Label21(1).Caption = "Proceso " & I
    Label21(1).Refresh
    DoEvents
    C = "INSERT INTO " & Cad & " SELECT  * " & SQL
    If Not EjecutaSQL(C) Then GoTo AccTraer
    I = I + 1
    Label21(1).Caption = "Proceso " & I
    Label21(1).Refresh
    DoEvents
    C = "DELETE " & SQL
    If Not EjecutaSQL(C) Then GoTo AccTraer
  
    Exit Sub
AccTraer:
    MsgBox "Error ejecutando SQL: " & vbCrLf & C
End Sub
