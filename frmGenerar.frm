VERSION 5.00
Begin VB.Form frmGenerar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar presupuestos"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmGenerar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMasiva 
      Caption         =   "Generacion masiva"
      Height          =   255
      Left            =   3000
      TabIndex        =   83
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Frame FrMasiva 
      Height          =   6615
      Left            =   120
      TabIndex        =   64
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   1320
         TabIndex        =   79
         Top             =   2280
         Width           =   3735
         Begin VB.OptionButton optEjer 
            Caption         =   "Actual"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   81
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optEjer 
            Caption         =   "Siguiente"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   80
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.OptionButton optIncre 
         Caption         =   "Presupuesto anterior"
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   78
         Top             =   3720
         Width           =   2535
      End
      Begin VB.OptionButton optIncre 
         Caption         =   "Ejercicio"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   77
         Top             =   3720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtInc 
         Height          =   285
         Left            =   1560
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   3675
         Width           =   855
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtDesCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox txtDesCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   6240
         Width           =   7335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   76
         Top             =   3720
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incremento"
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
         TabIndex        =   74
         Top             =   3240
         Width           =   1005
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
         Index           =   0
         Left            =   240
         TabIndex        =   73
         Top             =   2400
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Generar datos cuentas presupuestarias"
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
         Index           =   36
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   71
         Top             =   1605
         Width           =   465
      End
      Begin VB.Image imgcta 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmGenerar.frx":030A
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   720
         TabIndex        =   68
         Top             =   1245
         Width           =   465
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
         TabIndex        =   67
         Top             =   960
         Width           =   600
      End
      Begin VB.Image imgcta 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmGenerar.frx":6B5C
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrameN 
      Height          =   6615
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDesCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   360
         Width           =   4335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ejercicio siguiente"
         Height          =   255
         Left            =   5880
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5880
         TabIndex        =   4
         Top             =   2130
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   5880
         TabIndex        =   5
         Top             =   2490
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   5880
         TabIndex        =   6
         Top             =   2850
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   5880
         TabIndex        =   7
         Top             =   3210
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   5880
         TabIndex        =   8
         Top             =   3570
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   9
         Top             =   3930
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   5880
         TabIndex        =   10
         Top             =   4290
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   5880
         TabIndex        =   11
         Top             =   4650
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   5880
         TabIndex        =   12
         Top             =   5010
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   5880
         TabIndex        =   13
         Top             =   5370
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   5880
         TabIndex        =   14
         Top             =   5730
         Width           =   1575
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   5880
         TabIndex        =   15
         Top             =   6090
         Width           =   1575
      End
      Begin VB.TextBox txtAnual 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtPorc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   615
      End
      Begin VB.Image imgcta 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmGenerar.frx":D3AE
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "ENERO"
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
         Left            =   360
         TabIndex        =   62
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "FEBRERO"
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
         Index           =   1
         Left            =   360
         TabIndex        =   61
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "MARZO"
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
         Left            =   360
         TabIndex        =   60
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "ABRIL"
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
         Index           =   3
         Left            =   360
         TabIndex        =   59
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "MAYO"
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
         Index           =   4
         Left            =   360
         TabIndex        =   58
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "JUNIO"
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
         Index           =   5
         Left            =   360
         TabIndex        =   57
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "JULIO"
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
         Index           =   6
         Left            =   360
         TabIndex        =   56
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "AGOSTO"
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
         Index           =   7
         Left            =   360
         TabIndex        =   55
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "SEPTIEMBRE"
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
         Index           =   8
         Left            =   360
         TabIndex        =   54
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "OCTUBRE"
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
         Index           =   9
         Left            =   360
         TabIndex        =   53
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "NOVIEMBRE"
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
         Index           =   10
         Left            =   360
         TabIndex        =   52
         Top             =   5760
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "DICIEMBRE"
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
         Index           =   11
         Left            =   360
         TabIndex        =   51
         Top             =   6120
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " "
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
         Index           =   12
         Left            =   2160
         TabIndex        =   50
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   15
         Left            =   2160
         TabIndex        =   49
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   16
         Left            =   2160
         TabIndex        =   48
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   17
         Left            =   2160
         TabIndex        =   47
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   18
         Left            =   2160
         TabIndex        =   46
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   2160
         TabIndex        =   45
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   20
         Left            =   2160
         TabIndex        =   44
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   21
         Left            =   2160
         TabIndex        =   43
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   22
         Left            =   2160
         TabIndex        =   42
         Top             =   5760
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   23
         Left            =   2160
         TabIndex        =   41
         Top             =   6120
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   24
         Left            =   4080
         TabIndex        =   40
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   25
         Left            =   4080
         TabIndex        =   39
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   26
         Left            =   4080
         TabIndex        =   38
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   27
         Left            =   4080
         TabIndex        =   37
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   28
         Left            =   4080
         TabIndex        =   36
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   29
         Left            =   4080
         TabIndex        =   35
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   30
         Left            =   4080
         TabIndex        =   34
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   31
         Left            =   4080
         TabIndex        =   33
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   32
         Left            =   4080
         TabIndex        =   32
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   33
         Left            =   4080
         TabIndex        =   31
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   34
         Left            =   4080
         TabIndex        =   30
         Top             =   5760
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   35
         Left            =   4080
         TabIndex        =   29
         Top             =   6120
         Width           =   1500
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7560
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   1920
         Y1              =   1680
         Y2              =   6480
      End
      Begin VB.Line Line3 
         X1              =   3840
         X2              =   3840
         Y1              =   1680
         Y2              =   6480
      End
      Begin VB.Line Line4 
         X1              =   5760
         X2              =   5760
         Y1              =   1680
         Y2              =   6480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "ANTERIOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "ACTUAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   26
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   6120
         TabIndex        =   25
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   13
         Left            =   2160
         TabIndex        =   24
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   14
         Left            =   2160
         TabIndex        =   23
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   0
         X2              =   7680
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Importe ANUAL"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Incremento %"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.CheckBox ChkEliminar 
      Caption         =   "Eliminar datos (si los tuviera)"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   6720
      Width           =   975
   End
End
Attribute VB_Name = "frmGenerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Opcion As Byte
    '0 - Normal
    '1.- Generacion masiva

Dim Modo As Byte
    '0   En blanco
    '1   Pidiendo datos anual/porcentual
    '2   Mostrando datos



Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1

Dim Cad As String
Dim RS As Recordset
Dim I As Integer
Dim VV As Currency


Private AntiguoText As String 'Para comprobar si ha cambiado cosas o no

Private Sub Check1_Click()
    'Ha cambiado actual seiguiente
End Sub

Private Sub Check1_GotFocus()
    AntiguoText = Check1.Value
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check1_LostFocus()
    If CStr(Check1.Value) <> AntiguoText Then PonerDatos2
End Sub

Private Sub HacerClick()
Dim Aux As String

On Error GoTo EINs
    'Comprobamos que los campos son correctos
    
    If txtCta(0).Text = "" Then
        MsgBox "Introduzca la cuenta ", vbExclamation
        Exit Sub
    End If
    
    For I = 0 To 11
        If txtN(I).Text <> "" Then
            If Not IsNumeric(txtN(I).Text) Then
                MsgBox "Los valores de los importes deben de ser numéricos", vbExclamation
                Exit Sub
            End If
        End If
    Next I
    'Llegados aqui esta todo bien. Luego haremos dos cosas
    I = Year(vParam.fechaini)
    If Check1.Value Then I = I + 1

    If ChkEliminar.Value = 1 Then
        Cad = "DELETE FROM presupuestos WHERE codmacta='" & txtCta(0).Text & "' AND anopresu = " & I
        Conn.Execute Cad
    End If
    
    
    Cad = "INSERT INTO presupuestos (codmacta, anopresu, mespresu, imppresu) VALUES ('"
    Cad = Cad & txtCta(0).Text & "'," & I & ","
    For I = 0 To 11
        If txtN(I).Text <> "" Then
            Aux = TransformaComasPuntos(ImporteFormateado(txtN(I).Text))
            Conn.Execute Cad & I + 1 & "," & Aux & ")"
        End If
    Next I
    
    'Llegados aqui dejamos que vuelva a poner valores para otras cuentas
    If MsgBox("Datos generados.     ¿Salir?", vbQuestion + vbYesNoCancel) = vbYes Then
        I = Year(vParam.fechaini)
        If Check1.Value Then I = I + 1
        CadenaDesdeOtroForm = " presupuestos.codmacta = '" & txtCta(0).Text & "' and anopresu = " & I
        Unload Me
    Else
        txtAnual.Text = ""
        txtPorc.Text = ""
    End If

    Exit Sub
EINs:
    MuestraError Err.Number, "Insertando nuevos valores"
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMasica_Click()
    
End Sub

Private Sub chkMasiva_Click()
    Opcion = Abs(Me.chkMasiva.Value)
    PonerFrames
 
    PonleFoco txtCta(Opcion)
    
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    If Opcion = 0 Then
        HacerClick
    Else
        If GeneracionMasiva Then Unload Me
        CadenaDesdeOtroForm = ""
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()

    Unload Me

End Sub


Private Sub PonerFrames()
    Me.FrameN.Visible = (Opcion = 0)
    FrMasiva.Visible = (Opcion = 1)
End Sub
Private Sub Form_Load()
    Opcion = 0
    Limpiar Me
    LimpiarLabels
    Label5.Caption = ""
    PonerFrames
    chkMasiva.Visible = vUsu.Nivel < 2   'Solo administradores
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(I).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescta(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgcta_Click(Index As Integer)
    I = Index
    AntiguoText = txtCta(Index).Text
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.Show vbModal
    Set frmC = Nothing
        
End Sub


Private Sub txtAnual_GotFocus()
    PonFoco txtAnual
    AntiguoText = txtAnual.Text
End Sub

Private Sub txtAnual_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAnual_LostFocus()
    txtAnual.Text = Trim(txtAnual.Text)
    If AntiguoText = txtAnual.Text Then Exit Sub
        
    Cad = ""
    If txtAnual.Text <> "" Then
        If Not IsNumeric(txtAnual.Text) Then
            MsgBox "Campo numerico", vbExclamation
            txtAnual.Text = ""
            PonleFoco txtAnual
        Else
            If InStr(1, txtAnual.Text, ",") > 0 Then
                VV = ImporteFormateado(txtAnual.Text)
            Else
                VV = CCur(TransformaPuntosComas(txtAnual.Text))
            End If
            VV = Round((VV / 12), 2)
            For I = 0 To 11
                txtN(I).Text = VV
            Next I
            Cad = "OK"
        End If
        If txtAnual.Text <> "" Then txtPorc.Text = ""
    End If
    If Cad = "" Then LimpiarCampos
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = AntiguoText Then Exit Sub
    If txtCta(Index).Text = "" Then
        txtDescta(Index).Text = ""
    Else
        CadenaDesdeOtroForm = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(CadenaDesdeOtroForm, Cad) Then
                txtCta(Index).Text = CadenaDesdeOtroForm
                txtDescta(Index).Text = Cad
        Else
            MsgBox Cad, vbExclamation
            txtDescta(Index).Text = Cad
        End If
        CadenaDesdeOtroForm = ""
    End If
    If Opcion = 0 Then PonerDatos2
End Sub
         


Private Sub txtInc_GotFocus()
    PonFoco txtInc
End Sub

Private Sub txtInc_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtInc_LostFocus()
    txtInc.Text = Trim(txtInc.Text)
    If txtInc.Text = "" Then Exit Sub
    
End Sub

Private Sub txtN_GotFocus(Index As Integer)
    PonFoco txtN(Index)
    
End Sub

Private Sub txtN_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtN_LostFocus(Index As Integer)
With txtN(Index)
    .Text = Trim(.Text)
    If .Text = "" Then Exit Sub
    If Not IsNumeric(.Text) Then
        MsgBox "Los importes deben de ser numéricos.", vbExclamation
        Exit Sub
    End If
    If InStr(1, .Text, ",") > 0 Then
        VV = ImporteFormateado(.Text)
    Else
        VV = CCur(TransformaPuntosComas(.Text))
    End If
    .Text = Format(VV, FormatoImporte)
End With
End Sub

Private Sub txtPorc_GotFocus()
    PonFoco txtPorc
    AntiguoText = txtPorc.Text
End Sub

Private Sub txtPorc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
For I = 0 To 11
     Me.txtN(I).Text = ""
Next I

End Sub

Private Sub PonerDatosCuenta()

End Sub

Private Sub PonerDatos2()

    If txtCta(0).Text = "" Then
        LimpiarLabels
        If txtPorc.Text <> "" Then LimpiarCampos
        
    Else
        'Pondremos los datos en los campos
        PonerValoresAnteriores
        'Los valores actuales
        PonerValoresActuales
        
    End If

End Sub



Private Sub PonerValoresAnteriores()
On Error GoTo EPonerValoresAnteriores

    Cad = "Select * from presupuestos where codmacta='" & txtCta(0).Text & "' AND"
    Cad = Cad & " anopresu = "
    I = Year(vParam.fechaini)
    'Ejercicio siguiente
    If Check1.Value = 1 Then I = I + 1
    
    'Los anteriores al periodo solicitado es menos uno
    I = I - 1
    Cad = Cad & I & ";"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        While Not RS.EOF
            'El mes viene en el campo "mespresu"
            'Entonces para el label , k van del 12 al 23
            I = RS!mespresu + 11
            Label2(I).Caption = Format(RS!imppresu, FormatoImporte)
            
            'Ponemos el futuro valor(NUEVO)
            If txtPorc.Text <> "" Then
                'Porcentual
                VV = Round(RS!imppresu * CCur(txtPorc.Text), 2)
                VV = VV / 100
                VV = VV + RS!imppresu
                I = RS!mespresu - 1
                txtN(I).Text = Format(VV, FormatoImporte)
            End If
            
            
            'Sig
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    
    
Exit Sub
EPonerValoresAnteriores:
    MuestraError Err.Number, "Poner datos anterior"
End Sub


Private Sub PonerValoresActuales()
On Error GoTo EPonerValoresActauales

    Cad = "Select * from presupuestos where codmacta='" & txtCta(0).Text & "' AND"
    Cad = Cad & " anopresu = "
    I = Year(vParam.fechaini)
    'Ejercicio siguiente
    If Check1.Value = 1 Then I = I + 1
    
    Cad = Cad & I & ";"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        While Not RS.EOF
            'El mes viene en el campo "mespresu"
            'Entonces para el label , k van del 24 al 35
            I = RS!mespresu + 23
            Label2(I).Caption = Format(RS!imppresu, FormatoImporte)
            
            'Sig
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    
    
Exit Sub
EPonerValoresActauales:
    MuestraError Err.Number, "Poner datos anterior"
End Sub

Private Sub LimpiarLabels()
    For I = 12 To 35
        Label2(I).Caption = ""
    Next I
    
End Sub

Private Sub txtPorc_LostFocus()
    txtPorc.Text = Trim(txtPorc.Text)
    If txtPorc.Text = AntiguoText Then Exit Sub
    
    If txtPorc.Text <> "" Then txtAnual.Text = ""
    PonerDatos2
    
End Sub


Private Function GeneracionMasiva() As Boolean
Dim SQL As String
Dim Incremento As Currency

    On Error GoTo EGeneracionMasiva
    GeneracionMasiva = False
    
    'Generacion masvia de datos presupuestarios
    If txtInc.Text = "" Then
        MsgBox "Indique el incremento(%) a aplicar", vbExclamation
        Exit Function
    End If
    
    If MsgBox("Desea continuar con la generacion de datos presupuestarios?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    
    Incremento = ImporteSinFormato(txtInc.Text) / 100
    Set RS = New ADODB.Recordset
    
    
    'Obtenedre el SQL de  las cuentas
    
    If ChkEliminar.Value = 1 Then
        FijarSQLTablaPresu False
        SQL = "DELETE FROM presupuestos WHERE " & Cad
        Conn.Execute SQL
    End If
    
    FijarSQLTablaPresu True
    If optIncre(1).Value Then
        'Cojera los datos del presupuesto anterior
        SQL = "Select codmacta, anopresu anyo , mespresu mes ,imppresu debe,0 haber FROM presupuestos WHERE " & Cad
    
    Else
        'Select sum desde hlinapu
        Cad = Replace(Cad, "anopresu", "anopsald")
        Cad = Replace(Cad, "mespresu", "mespsald")
        SQL = "Select codmacta,anopsald anyo,mespsald mes,impmesde debe,impmesha haber FROM hsaldos WHERE " & Cad
        'Añado codmacta ultimo nivel
        SQL = SQL & "   AND codmacta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
        
    End If
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = "INSERT INTO presupuestos (codmacta, anopresu, mespresu, imppresu) VALUES "
    I = 0
    SQL = ""

    While Not RS.EOF
        I = I + 1
        VV = DBLet(RS!Debe, "N") - DBLet(RS!Haber, "N")
        VV = Round((VV * Incremento) + VV, 2)  'importe
        Cad = ",('" & RS!Codmacta & "'," & RS!Anyo & "," & RS!Mes & "," & TransformaComasPuntos(CStr(VV)) & ")"
        SQL = SQL & Cad
        If (I Mod 25) = 0 Then
             SQL = CadenaDesdeOtroForm & Mid(SQL, 2)  'QUITO la PRIMERa coma
            Conn.Execute SQL
            SQL = ""
        End If
        RS.MoveNext
    Wend
    RS.Close
    If SQL <> "" Then
        SQL = CadenaDesdeOtroForm & Mid(Cad, 2)  'QUITO la PRIMERa coma
         Conn.Execute SQL
    End If
    MsgBox "Proceso finalizado", vbInformation
    GeneracionMasiva = True
EGeneracionMasiva:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Set RS = Nothing
End Function


Private Sub FijarSQLTablaPresu(PeriodoAnterior As Boolean)
    Cad = ""
    'Ejercicio
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Año natural
        I = Year(vParam.fechaini)
        If optEjer(1).Value Then I = I + 1
        If PeriodoAnterior Then I = I - 1
        Cad = "anopresu = " & I
    Else
        I = Year(vParam.fechaini)
        If optEjer(1).Value Then I = I + 1
        If PeriodoAnterior Then I = I - 1
        
        Cad = "( anopresu = " & I & " AND mespresu >= " & Month(vParam.fechaini) & ") AND ("
        Cad = Cad & " ( anopresu = " & I + 1 & " AND mespresu <= " & Month(vParam.fechafin) & ") AND ("
        
    End If
    
    If txtCta(1).Text <> "" Then Cad = Cad & " AND codmacta >= '" & txtCta(1).Text & "'"
    If txtCta(2).Text <> "" Then Cad = Cad & " AND codmacta <= '" & txtCta(2).Text & "'"
    
End Sub
