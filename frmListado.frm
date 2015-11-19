VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   9000
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameImpresionTarjetas 
      Height          =   3645
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   6915
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5070
         TabIndex        =   108
         Top             =   2850
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepImpTarjetas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   106
         Top             =   2850
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   104
         Top             =   1770
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   103
         Top             =   1410
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   20
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   102
         Text            =   "Text5"
         Top             =   1410
         Width           =   3315
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   21
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "Text5"
         Top             =   1785
         Width           =   3315
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   105
         Text            =   "Combo1"
         Top             =   2310
         Width           =   2475
      End
      Begin VB.Label Label6 
         Caption         =   "Impresion de Tarjetas"
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
         Left            =   600
         TabIndex        =   112
         Top             =   450
         Width           =   5145
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   111
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   960
         TabIndex        =   110
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   21
         Left            =   630
         TabIndex        =   109
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Impresión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   107
         Top             =   2310
         Width           =   1275
      End
   End
   Begin VB.Frame FrameClientes 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1920
         TabIndex        =   98
         Text            =   "Combo1"
         Top             =   3210
         Width           =   1995
      End
      Begin VB.CommandButton cmdSubir 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton cmdBajar 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2715
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2715
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   1635
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1650
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdAceptar2 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5910
         TabIndex        =   5
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   6990
         TabIndex        =   6
         Top             =   3705
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   0
         Left            =   6360
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   7
         Left            =   630
         TabIndex        =   99
         Top             =   3180
         Width           =   840
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1560
         MouseIcon       =   "frmListado.frx":0620
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmListado.frx":0772
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1560
         MouseIcon       =   "frmListado.frx":08C4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1560
         MouseIcon       =   "frmListado.frx":0A16
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   46
         Left            =   6360
         TabIndex        =   21
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   20
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   19
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   18
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lbltitulo2 
         Caption         =   "Informe de Clientes"
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
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Colectivo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   15
         Top             =   1635
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameFacRectif 
      Height          =   5235
      Left            =   30
      TabIndex        =   44
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox Check2 
         Caption         =   "Recuperar Albaranes"
         Height          =   255
         Left            =   4560
         TabIndex        =   97
         Top             =   2100
         Width           =   2055
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   53
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Text            =   "dd/mm/yyyy"
         Top             =   2070
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text5"
         Top             =   1260
         Width           =   2745
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   52
         Top             =   1260
         Width           =   915
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   51
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Text            =   "dd/mm/yyyy"
         Top             =   1260
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1050
         MaxLength       =   7
         TabIndex        =   50
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   540
         MaxLength       =   3
         TabIndex        =   49
         Top             =   1260
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text5"
         Top             =   2865
         Width           =   5025
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   540
         MaxLength       =   6
         TabIndex        =   54
         Top             =   2865
         Width           =   915
      End
      Begin VB.CommandButton cmdAceptarFacRect 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4470
         TabIndex        =   56
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5670
         TabIndex        =   57
         Top             =   4500
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   645
         Index           =   87
         Left            =   540
         MaxLength       =   72
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   3570
         Width           =   6105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Factura Rectificativa"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   540
         TabIndex        =   63
         Top             =   2100
         Width           =   1965
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   13
         Left            =   2640
         Picture         =   "frmListado.frx":0B68
         ToolTipText     =   "Buscar fecha"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   2610
         Picture         =   "frmListado.frx":0BF3
         ToolTipText     =   "Buscar fecha"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   2940
         TabIndex        =   62
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   1920
         TabIndex        =   60
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   59
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   540
         TabIndex        =   58
         Top             =   1050
         Width           =   360
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1650
         MouseIcon       =   "frmListado.frx":0C7E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Cliente"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   540
         TabIndex        =   48
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Factura a Rectificar"
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
         Left            =   510
         TabIndex        =   46
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   82
         Left            =   600
         TabIndex        =   45
         Top             =   2520
         Width           =   480
      End
   End
   Begin VB.Frame FrameArticulos 
      Height          =   4455
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8595
      Begin VB.CheckBox Check1 
         Caption         =   "Resumen"
         Height          =   285
         Left            =   720
         TabIndex        =   96
         Top             =   3360
         Width           =   2145
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   34
         Top             =   3735
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   33
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1275
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   1275
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   32
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2355
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   2355
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":0DD0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":10DA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   1
         Left            =   6360
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   43
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   42
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   9
         Left            =   600
         TabIndex        =   41
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Artículos"
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
         Left            =   600
         TabIndex        =   40
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   39
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   38
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   37
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   6360
         TabIndex        =   36
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmListado.frx":13E4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmListado.frx":1536
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmListado.frx":1688
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar articulo"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmListado.frx":17DA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   2400
         Width           =   240
      End
   End
   Begin VB.Frame FrameProveedores 
      Height          =   4065
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   8595
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   6900
         TabIndex        =   89
         Top             =   3165
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   5820
         TabIndex        =   88
         Top             =   3150
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   87
         Top             =   1740
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   86
         Top             =   1410
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Text5"
         Top             =   1770
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text5"
         Top             =   1410
         Width           =   3015
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":192C
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListado.frx":1C36
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   2
         Left            =   6360
         TabIndex        =   90
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Proveedores"
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
         Left            =   600
         TabIndex        =   95
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   990
         TabIndex        =   94
         Top             =   1410
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   990
         TabIndex        =   93
         Top             =   1725
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   18
         Left            =   630
         TabIndex        =   92
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   6360
         TabIndex        =   91
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1590
         MouseIcon       =   "frmListado.frx":1F40
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1590
         MouseIcon       =   "frmListado.frx":2092
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1410
         Width           =   240
      End
   End
   Begin VB.Frame FrameGeneracionTurnos 
      Height          =   3375
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   5745
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   76
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1335
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1845
         MaxLength       =   1
         TabIndex        =   78
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1935
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4275
         TabIndex        =   75
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3090
         TabIndex        =   74
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   80
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Turno"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   79
         Top             =   1965
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   17
         Left            =   1530
         Picture         =   "frmListado.frx":21E4
         ToolTipText     =   "Buscar fecha"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Generación Masiva Turno"
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
         Index           =   1
         Left            =   585
         TabIndex        =   77
         Top             =   585
         Width           =   4455
      End
   End
   Begin VB.Frame FrameTarjetas 
      Height          =   3375
      Left            =   45
      TabIndex        =   64
      Top             =   0
      Width           =   5745
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   2190
         TabIndex        =   67
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   3375
         TabIndex        =   68
         Top             =   2700
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   66
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1935
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   65
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1575
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Impresión de Tarjetas"
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
         Left            =   585
         TabIndex        =   72
         Top             =   585
         Width           =   4455
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   15
         Left            =   1530
         Picture         =   "frmListado.frx":226F
         ToolTipText     =   "Buscar fecha"
         Top             =   1935
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   14
         Left            =   1530
         Picture         =   "frmListado.frx":22FA
         ToolTipText     =   "Buscar fecha"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   71
         Top             =   1965
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   70
         Top             =   1605
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   69
         Top             =   1365
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 10 .- Listado de Clientes
    ' 11 .- Listados de Articulos
    ' 12 .- Factura Rectificativa
    ' 13 .- Listado de Tarjetas de Gasolinera
    ' 14 .- Generacion Masiva de Turnos para Pobla del Duc
    ' 15 .- Listado de Proveedores
    
    ' 16 .- Impresion de tarjetas de clientes
    
    
Public Socio As String
Public TARJETA As String

    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(CLIENTE As String, observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmCol As frmManCoope 'Colectivos
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmFam As frmManFamia  'Familias
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic  'Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmPro As frmManProve  'Proveedores
Attribute frmPro.VB_VarHelpID = -1


Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check2_Click()
    imgBuscar(8).Enabled = (Check2.Value = 0)
    txtCodigo(8).Enabled = (Check2.Value = 0)
    If Check2.Value = 1 Then
        txtCodigo(8).Text = ""
        txtNombre(8).Text = ""
    End If
End Sub

Private Sub CmdAcepImpTarjetas_Click()
Dim sql As String
Dim rs As ADODB.Recordset
Dim nomdocu As String

Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
    
    If txtCodigo(20).Text = "" Or txtCodigo(21).Text = "" Then
        MsgBox "Debe introducir valor para el desde/hasta tarjeta. Reintroduzca.", vbExclamation
        PonerFoco txtCodigo(20)
        Exit Sub
    End If
    If Combo2.ListIndex = -1 Then
        MsgBox "Debe introducir un tipo de impresión. Revise.", vbExclamation
        PonerFocoCmb Combo2
        Exit Sub
    End If

    sql = "select ccc.fichrpt01, ccc.fichrpt02, ccc.fichrpt03, ccc.fichrpt04, ttt.tiptarje, ttt.codsocio, ttt.numtarje "
    sql = sql & " from scoope ccc, ssocio sss, starje ttt "
    sql = sql & " where ttt.codsocio = sss.codsocio "
    sql = sql & " and ttt.numtarje >= " & DBSet(txtCodigo(20).Text, "N")
    sql = sql & " and ttt.numtarje <= " & DBSet(txtCodigo(21).Text, "N")
    sql = sql & " and ccc.codcoope = sss.codcoope "
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not rs.EOF
        If rs.Fields(4).Value = 2 Then
            nomdocu = rs.Fields(3)
        Else
            Select Case Combo2.ListIndex
                Case 0 ' solo anverso
                    nomdocu = rs.Fields(0)
                Case 1 ' anverso + banda magnetica
                    nomdocu = rs.Fields(1)
                Case 2 ' solo banda magnetica
                    nomdocu = rs.Fields(2)
            End Select
        End If
    
    
        frmImprimir.NombreRPT = nomdocu
        
        ActivaTicket
        
        With frmVisReport
            .FormulaSeleccion = "{starje.codsocio}=" & DBSet(rs!codsocio, "N") & " and {starje.numtarje}= " & DBSet(rs!Numtarje, "N")
            .SoloImprimir = True
            .OtrosParametros = ""
            .NumeroParametros = 1
            .MostrarTree = False
            .Informe = App.path & "\informes\" & nomdocu
            .InfConta = False
            .ConSubInforme = False
            .SubInformeConta = ""
            .Opcion = 0
            .ExportarPDF = False
            .Show vbModal
        End With
        
        DesactivaTicket
        
        rs.MoveNext
    Wend

End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
    
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    Select Case Index
       Case 0 'Frame Informe de articulos
            '======== FORMULA  ====================================
            'D/H Familia
            cDesde = Trim(txtCodigo(6).Text)
            cHasta = Trim(txtCodigo(7).Text)
            nDesde = txtNombre(6).Text
            nHasta = txtNombre(7).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codfamia}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFamilia= """) Then Exit Sub
            End If
            
            'D/H Articulo
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codartic}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo= """) Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            numOp = PonerGrupo(1, ListView1(1).SelectedItem.Text)
            cadNombreRPT = "rManArtic.rpt"
            cadTitulo = "Listado de Artículos"
            
            If Me.Check1.Value = 1 Then
                cadNombreRPT = "rManArticResum.rpt"
                cadTitulo = "Listado de Artículos Resumido"
            End If
            
    Case 1 'listado de tarjetas de gasolinera
        'D/H Fecha
        cDesde = Trim(txtCodigo(14).Text)
        cHasta = Trim(txtCodigo(15).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{imptarjetas.fecha}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        End If
        
    
        cadTitulo = "Impresión Listado de Tarjetas "
        cadNombreRPT = "rListTarjetas.rpt"
        
    Case 2 ' generacion masiva de turnos
         If DatosOkTurnos Then
            If GeneracionTurnos Then
                frmTurnosContador.DeConsulta = False
                frmTurnosContador.Fecha = txtCodigo(17).Text
                frmTurnosContador.Turno = txtCodigo(16).Text
                frmTurnosContador.Show vbModal
            End If
            DesBloqueoManual ("GENTUR")
            cmdCancel_Click (4)
            Exit Sub
         Else
            Exit Sub
         End If
         
    Case 3 ' informe de proveedores
            'D/H Proveedor
            cDesde = Trim(txtCodigo(18).Text)
            cHasta = Trim(txtCodigo(19).Text)
            nDesde = txtNombre(18).Text
            nHasta = txtNombre(19).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & Tabla & ".codprove}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProveedor= """) Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            numOp = PonerGrupo(2, ListView1(2).SelectedItem.Text)
            cadNombreRPT = "rManProveedor.rpt"
            cadTitulo = "Listado de Proveedores"
    
    End Select
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla, cadSelect) Then
        LlamarImprimir
    End If

End Sub

'Frame Informe Clientes

Private Sub cmdAceptar2_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Colectivo
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    nDesde = txtNombre(2).Text
    nHasta = txtNombre(3).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHColectivo= """) Then Exit Sub
    End If
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        If OpcionListado = 10 Then
            Codigo = "{" & Tabla & ".codsocio}"
        ElseIf OpcionListado = 14 Then
            Codigo = "{" & Tabla & ".gruprove}"
        End If
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    '[Monica]06/02/2013: Introducimos el poder seleccionar que tipo de socios vamos a listar
    Select Case Combo1.ListIndex
        Case 0
            
        Case 1
            If Not AnyadirAFormula(cadSelect, "(ssocio.fechabaj is null)") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "(isnull({ssocio.fechabaj}))") Then Exit Sub
        Case 2
            If Not AnyadirAFormula(cadSelect, "not ssocio.fechabaj is null ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "not isnull({ssocio.fechabaj})") Then Exit Sub
    End Select
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
'    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
'    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
    numOp = PonerGrupo(1, ListView1(0).SelectedItem.Text)

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla, cadSelect) Then
        cadNombreRPT = "rManClien.rpt"
        cadTitulo = "Listado de Clientes"
        LlamarImprimir
    End If
End Sub

Private Sub cmdAceptarFacRect_Click()
    If DatosOk Then
'        RaiseEvent RectificarFactura(txtCodigo(8).Text, txtCodigo(87))
'        cmdCancel_Click (1)
        If CrearFacturaRectificativa(txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text, txtCodigo(87).Text, txtCodigo(8).Text, txtCodigo(13).Text, Check2.Value) = 0 Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            cmdCancel_Click (1)
        End If
    End If
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView1
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSubir_Click()
    SubirItemList Me.ListView1
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 10 ' Listado de Clientes
                PonerFoco txtCodigo(2)
                
                Combo1.ListIndex = 0
                
            Case 11 ' Listado de Articulos
                PonerFoco txtCodigo(6)
            
            Case 12 ' factura rectificativa
                '[Monica]18/01/2013: devolver albaranes al mantenimiento
                Check2.Value = 1
                txtCodigo(13).Text = Format(Now, "dd/mm/yyyy")
                PonerFoco txtCodigo(9)
                
            Case 13 ' Listado de tarjetas de gasolinera
                PonerFoco txtCodigo(14)
            
            Case 14 ' Generacion masiva de turno
                PonerFoco txtCodigo(17)
            
            Case 15 ' Listado de proveedores
                PonerFoco txtCodigo(18)
            
            Case 16 ' impresion de tarjetas
                PonerFoco txtCodigo(20)
                Combo2.ListIndex = 1
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Set List = New Collection
    For h = 24 To 27
        List.Add h
    Next h
    For h = 1 To 10
        List.Add h
    Next h
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
    
'    For h = 1 To List.Count
'        Me.imgBuscar(List.item(h)).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next h
' ### [Monica] 09/11/2006    he sustituido el anterior
    For h = 0 To imgBuscar.Count - 1
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
     
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameClientes.visible = False
    FrameArticulos.visible = False
    FrameFacRectif.visible = False
    FrameTarjetas.visible = False
    FrameGeneracionTurnos.visible = False
    FrameProveedores.visible = False
    FrameImpresionTarjetas.visible = False
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 10 '10: Listado de Clientes
        FrameClienteVisible True, h, w
        CargarListViewOrden (0)
        Me.lbltitulo2.Caption = "Informe de Clientes"
        Me.Label2(3).Caption = "Cliente"
        indFrame = 2
        Tabla = "ssocio"
        CargaCombo
    
    Case 11 ' Listado de Articulos
        FrameArticuloVisible True, h, w
        CargarListViewOrden (1)
        Me.lbltitulo2.Caption = "Informe de Artículos"
        Me.Label2(3).Caption = "Artículos"
        indFrame = 0
        Tabla = "sartic"
        
    Case 12 ' factura rectificativa
        FrameFacRectifVisible True, h, w
        indFrame = 1
        Tabla = "schfac"
'        PonerValoresFactura
    Case 13 ' Listado de tarjetas
        FrameTarjetaVisible True, h, w
        indFrame = 3
        Tabla = "imptarjetas"
    Case 14 ' generacion automatica de turnos
        FrameGeneracionTurnosVisible True, h, w
        indFrame = 3
        Tabla = "sturno"
        txtCodigo(17).Text = Format(Now, "dd/mm/yyyy")
    
    Case 15 ' informe de proveedores
        FrameProveedoresVisible True, h, w
        CargarListViewOrden (2)
        indFrame = 0
        Tabla = "sartic"
    
    Case 16 ' impresion de tarjetas
        FrameImpresionTarjetasVisible True, h, w
        indFrame = 0
        Tabla = "ssocio"
        txtCodigo(20).Text = Format(TARJETA, "000000")
        txtCodigo(21).Text = Format(TARJETA, "000000")
        txtNombre(20).Text = DevuelveValor("select nomtarje from starje where numtarje = " & DBSet(txtCodigo(20).Text, "N"))
        txtNombre(21).Text = txtNombre(20).Text
        CargaCombo2
    
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(11).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Para listados básicos
'            Select Case OpcionListado
'                Case 1 'Listado Actividades
'                    AbrirFrmColectivos (Index)
'            End Select
            AbrirFrmArticulos (Index)
            
        Case 2, 3 'COLECTIVOS
            AbrirFrmColectivos (Index)
            
        Case 4, 5, 8 'CLIENTES
            AbrirFrmClientes (Index)
            
        Case 6, 7
            AbrirFrmFamilias (Index)
            
        Case 9, 10 'proveedores
            AbrirFrmProveedores (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub imgFec_Click(Index As Integer)
Dim indice As Integer
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    indice = Index
        
    imgFec(11).Tag = indice 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(indice).Text <> "" Then frmC.NovaData = txtCodigo(indice).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(11).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 2 'colectivo desde
            Case 3: KEYBusqueda KeyAscii, 3 'colectivo hasta
            Case 4: KEYBusqueda KeyAscii, 4 'cliente desde
            Case 5: KEYBusqueda KeyAscii, 5 'cliente hasta
            Case 8: KEYBusqueda KeyAscii, 8 'cliente de la factura rectificativa
            Case 11: KEYFecha KeyAscii, 11 'fecha factura a rectificar
            Case 13: KEYFecha KeyAscii, 13 'fecha factura rectificativa
            Case 14: KEYFecha KeyAscii, 14 'fecha desde
            Case 15: KEYFecha KeyAscii, 15 'fecha hasta
            
            Case 17: KEYFecha KeyAscii, 17 'fecha turno
            
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim TipoMov As String

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'ARTICULOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3 'COLECTIVOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
        
        Case 4, 5, 8 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "ssocio", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
            If Index = 8 Then ' en la factura rectificativa el nuevo cliente ha de existir
                If txtCodigo(8).Text <> "" And txtNombre(8).Text = "" Then
                    MsgBox "El cliente introducido no existe. Si introduce número de cliente éste debe existir.", vbExclamation
                    PonerFoco txtCodigo(8)
                End If
            End If
        Case 6, 7 'FAMILIA
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sfamia", "nomfamia", "codfamia", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
            
        Case 11 'FECHAS de la factura a rectificar
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            If txtCodigo(9).Text <> "" And txtCodigo(10).Text <> "" And txtCodigo(11).Text <> "" Then
                PonerValoresFactura
            End If
            
        Case 14, 15 'fechas
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
            
        Case 9 'SERIE de la factura a rectificar
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
            If txtCodigo(9).Text <> "" Then
                TipoMov = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(txtCodigo(9).Text, "T"))
                If TipoMov = "FAT" Then
                    MsgBox "Esta Serie de Facturas es de TPV, no se permite hacer la rectificativa. Reintroduzca.", vbExclamation
                    txtCodigo(9).Text = ""
                    PonerFoco txtCodigo(9)
                End If
            End If
            If txtCodigo(9).Text <> "" And txtCodigo(10).Text <> "" And txtCodigo(11).Text <> "" Then
                PonerValoresFactura
            End If
        
        Case 10 'FACTURAS de la factura a rectificar
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            If txtCodigo(9).Text <> "" And txtCodigo(10).Text <> "" And txtCodigo(11).Text <> "" Then
                PonerValoresFactura
            End If
        
        Case 13 'FECHAS de la factura rectificativa
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 17 'FECHAS del turno para la generacion masiva
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    Conexion = cPTours    'Conexión a BD: Ariges
'    Select Case OpcionListado
'        Case 7 'Traspaso de Almacenes
'            cad = cad & "Nº Trasp|scatra|codtrasp|N|0000000|40·Almacen Origen|scatra|almaorig|N|000|20·Almacen Destino|scatra|almadest|N|000|20·Fecha|scatra|fechatra|F||20·"
'            Tabla = "scatra"
'            titulo = "Traspaso Almacenes"
'        Case 8 'Movimientos de Almacen
'            cad = cad & "Nº Movim.|scamov|codmovim|N|0000000|40·Almacen|scamov|codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
'            Tabla = "scamov"
'            titulo = "Movimientos Almacen"
'        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
'                   '12: Inventario Articulos
'                   '14:Actualizar Diferencias de Stock Inventariado
'                   '16: Listado Valoracion stock inventariado
'            cad = cad & "Código|sartic|codartic|T||30·Denominacion|sartic|nomartic|T||70·"
'            Tabla = "sartic"
'            titulo = "Articulos"
'    End Select
          
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub FrameClienteVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de clientes
    Me.FrameClientes.visible = visible
    If visible = True Then
        Me.FrameClientes.Top = -90
        Me.FrameClientes.Left = 0
        Me.FrameClientes.Height = 4820
        Me.FrameClientes.Width = 8600
        w = Me.FrameClientes.Width
        h = Me.FrameClientes.Height
    End If
End Sub

Private Sub FrameProveedoresVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de proveedores
    Me.FrameProveedores.visible = visible
    If visible = True Then
        Me.FrameProveedores.Top = -90
        Me.FrameProveedores.Left = 0
        Me.FrameProveedores.Height = 4065
        Me.FrameProveedores.Width = 8600
        w = Me.FrameProveedores.Width
        h = Me.FrameProveedores.Height
    End If
End Sub


Private Sub FrameImpresionTarjetasVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de proveedores
    Me.FrameImpresionTarjetas.visible = visible
    If visible = True Then
        Me.FrameImpresionTarjetas.Top = -90
        Me.FrameImpresionTarjetas.Left = 0
        Me.FrameImpresionTarjetas.Height = 3645
        Me.FrameImpresionTarjetas.Width = 6915
        w = Me.FrameImpresionTarjetas.Width
        h = Me.FrameImpresionTarjetas.Height
    End If
End Sub


Private Sub FrameArticuloVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de clientes
    Me.FrameArticulos.visible = visible
    If visible = True Then
        Me.FrameArticulos.Top = -90
        Me.FrameArticulos.Left = 0
        Me.FrameArticulos.Height = 4820
        Me.FrameArticulos.Width = 8600
        w = Me.FrameArticulos.Width
        h = Me.FrameArticulos.Height
    End If
End Sub

Private Sub FrameTarjetaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de clientes
    Me.FrameTarjetas.visible = visible
    If visible = True Then
        Me.FrameTarjetas.Top = -90
        Me.FrameTarjetas.Left = 0
        Me.FrameTarjetas.Height = 3375
        Me.FrameTarjetas.Width = 5745
        w = Me.FrameTarjetas.Width
        h = Me.FrameTarjetas.Height
    End If
End Sub

Private Sub FrameGeneracionTurnosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de clientes
    Me.FrameGeneracionTurnos.visible = visible
    If visible = True Then
        Me.FrameGeneracionTurnos.Top = -90
        Me.FrameGeneracionTurnos.Left = 0
        Me.FrameGeneracionTurnos.Height = 3375
        Me.FrameGeneracionTurnos.Width = 5745
        w = Me.FrameGeneracionTurnos.Width
        h = Me.FrameGeneracionTurnos.Height
    End If
End Sub


Private Sub FrameFacRectifVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de clientes
    Me.FrameFacRectif.visible = visible
    If visible = True Then
        Me.FrameFacRectif.Top = -90
        Me.FrameFacRectif.Left = 0
        Me.FrameFacRectif.Height = 5235
        Me.FrameFacRectif.Width = 7155
        w = Me.FrameFacRectif.Width
        h = Me.FrameFacRectif.Height
    End If
End Sub


Private Sub CargarListViewOrden(Index As Integer)
Dim ItmX As ListItem

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear
    ListView1(Index).ColumnHeaders.Add , , "Campo", 1390

    Select Case Index
        Case 0
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Colectivo"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Tipo Cliente"
        Case 1
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Familia"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Artículo"
        Case 2
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Código"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Alfabético"
    End Select
        
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomcampo As String

    campo = "pGroup" & numGrupo & "="
    nomcampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0

    Select Case cadgrupo
        Case "Colectivo"
            cadParam = cadParam & campo & "{" & Tabla & ".codcoope}" & "|"
            cadParam = cadParam & nomcampo & " {" & "scoope" & ".nomcoope}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Tipo Cliente""" & "|"
            numParam = numParam + 3
            
        Case "Tipo Cliente"
            cadParam = cadParam & campo & "{" & Tabla & ".tipsocio}" & "|"
            cadParam = cadParam & nomcampo & " {" & "tiposoci" & ".nomtipso}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Colectivo""" & "|"
            numParam = numParam + 3

        Case "Artículo"
            cadParam = cadParam & campo & "{" & Tabla & ".codartic}" & "|"
            cadParam = cadParam & nomcampo & " {" & "sartic" & ".nomartic}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Artículo""" & "|"
            numParam = numParam + 3
            
        Case "Familia"
            cadParam = cadParam & campo & "{" & Tabla & ".codfamia}" & "|"
            cadParam = cadParam & nomcampo & " {" & "sfamia" & ".nomfamia}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Familia""" & "|"
            numParam = numParam + 3

        Case "Código"
            cadParam = cadParam & "pOrden={proveedor.codprove}|"
            numParam = numParam + 1
            
        Case "Alfabético"
            cadParam = cadParam & "pOrden={proveedor.nomprove}|"
            numParam = numParam + 1
            
    End Select

End Function

Private Sub AbrirFrmColectivos(indice As Integer)
    indCodigo = indice
    Set frmCol = New frmManCoope
    frmCol.DatosADevolverBusqueda = "0|1|"
    frmCol.DeConsulta = True
    frmCol.CodigoActual = txtCodigo(indCodigo)
    frmCol.Show vbModal
    Set frmCol = Nothing
End Sub

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmcli = New frmManClien
    frmcli.DatosADevolverBusqueda = "0|1|"
    frmcli.DeConsulta = True
    frmcli.CodigoActual = txtCodigo(indCodigo)
    frmcli.Show vbModal
    Set frmcli = Nothing
End Sub
 
Private Sub AbrirFrmProveedores(indice As Integer)
    indCodigo = indice
    Set frmPro = New frmManProve
    frmPro.DatosADevolverBusqueda = "0|1|"
    frmPro.DeConsulta = True
    frmPro.CodigoActual = txtCodigo(indCodigo)
    frmPro.Show vbModal
    Set frmPro = Nothing
End Sub
 
 
 
Private Sub AbrirFrmFamilias(indice As Integer)
    indCodigo = indice
    Set frmFam = New frmManFamia
    frmFam.DatosADevolverBusqueda = "0|1|"
    frmFam.DeConsulta = True
    frmFam.CodigoActual = txtCodigo(indCodigo)
    frmFam.Show vbModal
    Set frmcli = Nothing
End Sub
 
Private Sub AbrirFrmArticulos(indice As Integer)
    indCodigo = indice
    Set frmArt = New frmManArtic
    frmArt.DatosADevolverBusqueda = "0|1|"
    frmArt.DeConsulta = True
    frmArt.CodigoActual = txtCodigo(indCodigo)
    frmArt.Show vbModal
    Set frmcli = Nothing
End Sub


Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = OpcionListado
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub

Private Sub PonerValoresFactura()
Dim intconta As String
Dim cad As String
'    txtCodigo(9).Text = RecuperaValor(CadTag, 1)
'    txtCodigo(10).Text = RecuperaValor(CadTag, 2)
'    txtCodigo(11).Text = RecuperaValor(CadTag, 3)
'    txtCodigo(12).Text = RecuperaValor(CadTag, 4)
'    txtNombre(9).Text = RecuperaValor(CadTag, 5)
'    Contabilizada = RecuperaValor(CadTag, 6)
     intconta = "intconta"
     txtCodigo(12).Text = ""
     txtCodigo(12).Text = DevuelveDesdeBDNew(cPTours, "schfac", "codsocio", "letraser", txtCodigo(9).Text, "T", intconta, "numfactu", txtCodigo(10).Text, "N", "fecfactu", txtCodigo(11).Text, "F")
     If txtCodigo(12).Text <> "" Then
        txtNombre(9).Text = PonerNombreDeCod(txtCodigo(12), "ssocio", "nomsocio", "codsocio", "N")
        Contabilizada = CInt(intconta)
     Else
        cad = "No existe la factura. Reintroduzca. " & vbCrLf & vbCrLf
        cad = cad & "   Serie: " & txtCodigo(9).Text & vbCrLf
        cad = cad & "   Factura: " & txtCodigo(10).Text & vbCrLf
        cad = cad & "   Fecha: " & txtCodigo(11).Text & vbCrLf
        cad = cad & vbCrLf
        MsgBox cad, vbExclamation
        txtCodigo(9).Text = ""
        txtCodigo(10).Text = ""
        txtCodigo(11).Text = ""
        PonerFoco txtCodigo(9)
     End If
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim sql As String
Dim sql2 As String
Dim vClien As CSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim TipoMov As String


    b = True
    If txtCodigo(9).Text = "" Or txtCodigo(10).Text = "" Or txtCodigo(11).Text = "" Then
        MsgBox "Debe introducir la letra de serie, el número de factura y la fecha de factura para localizar la factura a rectificar", vbExclamation
        b = False
    End If
    If b And vParamAplic.Cooperativa = 2 Then
        If txtCodigo(8).Text = "" Then
            MsgBox "Debe introducir el cliente. Reintroduzca.", vbExclamation
            b = False
        Else
            ' obtenemos la cooperativa del anterior cliente y del nuevo pq tienen que coincidir
            ' anterior cliente
            sql = ""
            sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codcoope", "codsocio", txtCodigo(12).Text, "N")
            ' nuevo cliente
            sql2 = ""
            sql2 = DevuelveDesdeBDNew(cPTours, "ssocio", "codcoope", "codsocio", txtCodigo(8).Text, "N")
            If sql <> sql2 Then
                MsgBox "El nuevo cliente debe pertenecer al mismo colectivo que el cliente de la factura a rectificar. Reintroduzca.", vbExclamation
                b = False
            End If
        End If
    End If
    
'    If b And Contabilizada = 1 And vParamAplic.NumeroConta <> 0 And txtCodigo(8).Text <> "" Then 'comprobamos que la cuenta contable del nuevo cliente existe
'        Set vClien = New CSocio
'        If vClien.LeerDatos(txtCodigo(8).Text) Then
'            sql = ""
'            sql = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", vClien.CuentaConta, "T")
'            If sql = "" Then
'                MsgBox "La cuenta contable del nuevo cliente no existe. Revise", vbExclamation
'                b = False
'            End If
'        End If
'    End If

' añadido
'    b = True
    
    '[Monica]18/01/2013: no se permite rectificar una factura de TPV
    If b Then
        TipoMov = DevuelveValor("select codtipom from stipom where letraser = " & DBSet(txtCodigo(9).Text, "T"))
        If TipoMov = "FAT" Then
            MsgBox "Este Factura es de TPV, no se permite hacer la factura rectificativa", vbExclamation
            b = False
        End If
    End If

    If b Then
        If ConTarjetaProfesional(txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text) Then
            MsgBox "Este Factura tiene alguna tarjeta profesional, no se permite hacer la factura rectificativa", vbExclamation
            b = False
        Else
            If txtCodigo(13).Text = "" Then
                MsgBox "Debe introducir obligatoriamente una Fecha de Facturación.", vbExclamation
                b = False
                PonerFoco txtCodigo(13)
            Else
                    If Not FechaDentroPeriodoContable(CDate(txtCodigo(13).Text)) Then
                        Mens = "La Fecha de Facturación no es del ejercicio actual ni siguiente. Reintroduzca."
                        MsgBox Mens, vbExclamation
                        b = False
                        PonerFoco txtCodigo(13)
                    Else
                        'VRS:2.0.1(0)
                        If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(13).Text)) Then
                            Mens = "  La Fecha de Facturación es inferior a la última liquidación de Iva. " & vbCrLf & vbCrLf
                            ' unicamente si el usuario es root el proceso continuará
                            If vSesion.Nivel > 0 Then
                                Mens = Mens & "  El proceso no continuará."
                                MsgBox Mens, vbExclamation
                                b = False
                                PonerFoco txtCodigo(13)
                            Else
                                Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                                If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                    b = False
                                    PonerFoco txtCodigo(13)
                                End If
                            End If
                        End If
                        ' la fecha de factura no debe ser inferior a la ultima factura de la serie
                        numser = "letraser"
                        numfactu = ""
                        numfactu = DevuelveDesdeBDNew(cPTours, "stipom", "contador", "codtipom", "FAG", "T", numser)
                        If numfactu <> "" Then
                            If FechaFacturaInferiorUltimaFacturaSerieHco(CDate(txtCodigo(13).Text), CLng(numfactu), numser, 0) Then
                                Mens = "La Fecha de Factura es inferior a la última factura de la serie." & vbCrLf & vbCrLf
                                Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                                If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                    b = False
                                    PonerFoco txtCodigo(13)
                                End If
                            End If
                        End If
                    End If
            End If
        End If
    End If
    DatosOk = b


' end añadido
    If b And txtCodigo(87).Text = "" Then
        MsgBox "Para rectificar una factura ha de introducir obligatoriamente un motivo. Reintroduzca", vbExclamation
        b = False
    End If
    DatosOk = b


' una factura no puede estar rectificada más de una vez
    If b Then
        sql = "select count(*) from schfac where rectif_letraser = " & DBSet(txtCodigo(9).Text, "T") & " and rectif_numfactu = " & DBSet(txtCodigo(10).Text, "N")
        sql = sql & " and rectif_fecfactu = " & DBSet(txtCodigo(11).Text, "F")
        If DevuelveValor(sql) >= 1 Then
            MsgBox "Esta factura ya ha sido rectificada. Revise.", vbExclamation
            b = False
            PonerFoco txtCodigo(9)
        End If
    End If
    DatosOk = b


End Function

Private Function DatosOkTurnos() As Boolean
Dim b As Boolean
Dim sql As String
Dim sql2 As String
Dim vClien As CSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date

    b = True
    If txtCodigo(17).Text = "" Then
        MsgBox "Debe introducir la fecha. Reintroduzca.", vbExclamation
        PonerFoco txtCodigo(17)
        b = False
    End If
    
    If b And txtCodigo(16).Text = "" Then
        MsgBox "Debe introducir el turno. Reintroduzca.", vbExclamation
        PonerFoco txtCodigo(16)
        b = False
    End If
    
    ' comprobamos que la fecha / turno no este introducida ya
    If b Then
        sql = "select count(*) from sturno where codturno = " & DBSet(txtCodigo(16).Text, "N")
        sql = sql & " and fechatur = " & DBSet(txtCodigo(17).Text, "F") & " and tiporegi = 0 "
        If TotalRegistros(sql) <> 0 Then
            MsgBox "Existe el turno creado de contadores. Revise.", vbExclamation
            PonerFoco txtCodigo(17)
            b = False
        End If
    End If
    
    
    
    DatosOkTurnos = b

End Function

Private Function ConTarjetaProfesional(letraser As String, numfactu As String, fecfactu As String) As Boolean
Dim sql As String
Dim rs As ADODB.Recordset

    sql = "select count(*) from slhfac, starje where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
    sql = sql & " and fecfactu = " & DBSet(fecfactu, "F") & " and starje.tiptarje = 2 and slhfac.numtarje = starje.numtarje "
    
    ConTarjetaProfesional = (TotalRegistros(sql) <> 0)

End Function

Private Function GeneracionTurnos() As Boolean
Dim vSQL As String
Dim SqlAux As String
Dim sql As String
Dim sql2 As String
Dim TurnoAux As Integer
Dim FechaAux As Date
Dim rs As ADODB.Recordset
    
Dim FechaMax As Date
Dim TurnoMax As Integer
    
    On Error GoTo eGeneracionTurnos
    
    GeneracionTurnos = False
    
    
    sql = "GENTUR" 'GENeneracion de TURnos
    
    'Bloquear para que nadie mas pueda generar automaticamente
    DesBloqueoManual (sql)
    If Not BloqueoManual(sql, "1") Then
        MsgBox "No se puede generar automáticamente turnos. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    sql = "insert into sturno (fechatur,codturno,numlinea,tiporegi,numtanqu,nummangu,codartic,tipocred,litrosve,"
    sql = sql & "importel,containi,contafin) values "
    
    TurnoAux = CCur(txtCodigo(16).Text)
    FechaAux = CDate(txtCodigo(17).Text)
    
    'FechaAux y TurnoAux es lo que le corresponderia (turno anterior en gasolinera automática)
    If TurnoAux = 1 Then
        FechaAux = FechaAux - 1
        TurnoAux = 2
    Else
        TurnoAux = 1
    End If
    
    sql2 = "select max(fechatur) from sturno where fechatur <= " & DBSet(FechaAux, "F")
    sql2 = sql2 & " and tiporegi = 0 " ' sólo seleccionamos los registros que sean contadores
    
    FechaMax = DevuelveValor(sql2)
    
    If FechaMax = "0:00:00" Then
        MsgBox "No existe un turno anterior de contadores. Debe crearlo a mano.", vbExclamation
        Exit Function
    End If
    
    
    If FechaMax < FechaAux Then
        sql2 = "select max(codturno) from sturno where fechatur = " & DBSet(FechaMax, "F")
        sql2 = sql2 & " and tiporegi = 0 " ' sólo seleccionamos los registros que sean contadores
        TurnoMax = DevuelveValor(sql2)
    Else
        sql2 = "select max(codturno) from sturno where fechatur = " & DBSet(FechaMax, "F")
        sql2 = sql2 & " and tiporegi = 0 " ' sólo seleccionamos los registros que sean contadores
        TurnoMax = DevuelveValor(sql2)
    End If
    
    SqlAux = ""
    
    sql2 = "select * from sturno where fechatur = " & DBSet(FechaMax, "F")
    sql2 = sql2 & " and codturno = " & DBSet(TurnoMax, "N")
    sql2 = sql2 & " and tiporegi = 0 " ' sólo seleccionamos los registros que sean contadores
    
    Set rs = New ADODB.Recordset
    rs.Open sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not rs.EOF
        SqlAux = SqlAux & "(" & DBSet(txtCodigo(17).Text, "F") & "," & DBSet(txtCodigo(16).Text, "N") & ","
        SqlAux = SqlAux & DBSet(rs!NumLinea, "N") & ","
        SqlAux = SqlAux & "0," ' tipo registro = contadores
        SqlAux = SqlAux & DBSet(rs!numtanqu, "N") & ","
        SqlAux = SqlAux & DBSet(rs!nummangu, "N") & ","
        SqlAux = SqlAux & DBSet(rs!codArtic, "N") & ","
        SqlAux = SqlAux & DBSet(rs!tipocred, "N") & ","
        SqlAux = SqlAux & "0,0," ' litrosve, importe
        SqlAux = SqlAux & DBSet(rs!ContaFin, "N") & ","
        SqlAux = SqlAux & "0),"
        
        rs.MoveNext
    Wend
    
    If SqlAux <> "" Then
        SqlAux = Mid(SqlAux, 1, Len(SqlAux) - 1)
        Conn.Execute sql & SqlAux
    End If

    GeneracionTurnos = True
    Exit Function


eGeneracionTurnos:
    MuestraError Err.Number, "Generacion Turnos", Err.Description
End Function



Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Combo1.AddItem "Todos"
    Combo1.ItemData(Combo1.NewIndex) = 0
    
    Combo1.AddItem "De Alta"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "De Baja"
    Combo1.ItemData(Combo1.NewIndex) = 2
End Sub


' IMPRESION DE TICKETS DESDE AQUI

Private Sub CargaCombo2()
    Combo2.Clear
    'Conceptos
    Combo2.AddItem "Sólo Anverso"
    Combo2.ItemData(Combo2.NewIndex) = 0
    
    Combo2.AddItem "Anverso + Banda Magnética"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    Combo2.AddItem "Sólo Banda Magnética"
    Combo2.ItemData(Combo2.NewIndex) = 2
End Sub



Private Sub ActivaTicket()
    ImpresoraDefecto = Printer.DeviceName
    XPDefaultPrinter vParamAplic.ImpresoraTarjetas
End Sub

Private Sub DesactivaTicket()
    XPDefaultPrinter ImpresoraDefecto
End Sub


'---------------- Procesos para cambio de impresora por defecto ------------------
Private Sub XPDefaultPrinter(PrinterName As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim r As Long
    ' Get the printer information for the currently selected
    ' printer in the list. The information is taken from the
    ' WIN.INI file.
    Buffer = Space(1024)
    r = GetProfileString("PrinterPorts", PrinterName, "", _
        Buffer, Len(Buffer))

    ' Parse the driver name and port name out of the buffer
    GetDriverAndPort Buffer, DriverName, PrinterPort

       If DriverName <> "" And PrinterPort <> "" Then
           SetDefaultPrinter PrinterName, DriverName, PrinterPort
       End If
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim L As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub
'------------------ Fin de los procesos relacionados con el cambio de impresora ----





