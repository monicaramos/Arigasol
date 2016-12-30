VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación por Cliente"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6555
   Icon            =   "frmFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   5685
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   6285
      Begin VB.CheckBox ChkContado 
         Caption         =   "Contado"
         Height          =   195
         Left            =   4230
         TabIndex        =   57
         Top             =   2100
         Width           =   1545
      End
      Begin VB.CheckBox ChkInterna 
         Caption         =   "Interna"
         Height          =   195
         Left            =   4230
         TabIndex        =   56
         Top             =   2460
         Width           =   1545
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1245
         Left            =   390
         TabIndex        =   45
         Top             =   3600
         Width           =   5625
         Begin VB.Frame Frame7 
            Caption         =   "Tipo Facturación"
            ForeColor       =   &H00972E0B&
            Height          =   585
            Left            =   30
            TabIndex        =   52
            Top             =   30
            Width           =   5235
            Begin VB.OptionButton Option1 
               Caption         =   "Gasóleo B"
               Height          =   255
               Index           =   7
               Left            =   1590
               TabIndex        =   55
               Top             =   240
               Width           =   1035
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Gasolinas"
               Height          =   255
               Index           =   6
               Left            =   150
               TabIndex        =   54
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Resto de Productos"
               Height          =   255
               Index           =   5
               Left            =   3210
               TabIndex        =   53
               Top             =   240
               Width           =   1725
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   705
            Left            =   -180
            TabIndex        =   46
            Top             =   450
            Visible         =   0   'False
            Width           =   5655
            Begin VB.TextBox txtCodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   11
               Left            =   1500
               MaxLength       =   6
               TabIndex        =   49
               Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
               Top             =   450
               Width           =   735
            End
            Begin VB.TextBox txtNombre 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   2310
               Locked          =   -1  'True
               TabIndex        =   48
               Text            =   "Text5"
               Top             =   450
               Width           =   3135
            End
            Begin VB.TextBox txtCodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   10
               Left            =   1500
               MaxLength       =   10
               TabIndex        =   47
               Tag             =   "Código Postal|T|S|||clientes|codposta|||"
               Top             =   120
               Width           =   1050
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Banco Propio"
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
               Index           =   10
               Left            =   210
               TabIndex        =   51
               Top             =   450
               Width           =   930
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   2
               Left            =   1230
               MouseIcon       =   "frmFacturas.frx":000C
               MousePointer    =   4  'Icon
               ToolTipText     =   "Buscar cliente"
               Top             =   450
               Width           =   240
            End
            Begin VB.Image imgFec 
               Height          =   240
               Index           =   0
               Left            =   1230
               Picture         =   "frmFacturas.frx":015E
               ToolTipText     =   "Buscar fecha"
               Top             =   120
               Width           =   240
            End
            Begin VB.Label Label4 
               Caption         =   "Fecha Vto"
               ForeColor       =   &H00972E0B&
               Height          =   255
               Index           =   9
               Left            =   180
               TabIndex        =   50
               Top             =   120
               Width           =   885
            End
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   600
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   600
         Width           =   990
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   3285
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   2910
         Width           =   3165
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   8
         Top             =   3285
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2910
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2430
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2070
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4785
         TabIndex        =   10
         Top             =   5220
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   5220
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1170
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1515
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1170
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1515
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   4860
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1155
         Left            =   480
         TabIndex        =   31
         Top             =   3660
         Width           =   5625
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   705
            Left            =   -180
            TabIndex        =   36
            Top             =   450
            Visible         =   0   'False
            Width           =   5655
            Begin VB.TextBox txtCodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   9
               Left            =   1500
               MaxLength       =   10
               TabIndex        =   37
               Tag             =   "Código Postal|T|S|||clientes|codposta|||"
               Top             =   30
               Width           =   1050
            End
            Begin VB.TextBox txtNombre 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   285
               Index           =   8
               Left            =   2310
               Locked          =   -1  'True
               TabIndex        =   39
               Text            =   "Text5"
               Top             =   360
               Width           =   3135
            End
            Begin VB.TextBox txtCodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   8
               Left            =   1500
               MaxLength       =   6
               TabIndex        =   38
               Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label4 
               Caption         =   "Fecha Vto"
               ForeColor       =   &H00972E0B&
               Height          =   255
               Index           =   6
               Left            =   180
               TabIndex        =   41
               Top             =   30
               Width           =   885
            End
            Begin VB.Image imgFec 
               Height          =   240
               Index           =   9
               Left            =   1230
               Picture         =   "frmFacturas.frx":01E9
               ToolTipText     =   "Buscar fecha"
               Top             =   30
               Width           =   240
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   8
               Left            =   1230
               MouseIcon       =   "frmFacturas.frx":0274
               MousePointer    =   4  'Icon
               ToolTipText     =   "Buscar cliente"
               Top             =   360
               Width           =   240
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Banco Propio"
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
               Index           =   5
               Left            =   210
               TabIndex        =   40
               Top             =   360
               Width           =   930
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   405
            Left            =   2670
            TabIndex        =   33
            Top             =   60
            Visible         =   0   'False
            Width           =   2505
            Begin VB.TextBox txtCodigo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   1440
               MaxLength       =   7
               TabIndex        =   34
               Top             =   60
               Width           =   830
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Factura Inicial"
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
               Left            =   240
               TabIndex        =   35
               Top             =   60
               Width           =   1005
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Facturación Cepsa"
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Facturación"
         ForeColor       =   &H00972E0B&
         Height          =   585
         Left            =   420
         TabIndex        =   26
         Top             =   3600
         Width           =   5325
         Begin VB.OptionButton Option1 
            Caption         =   "Interna"
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   42
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Clientes/Tarjetas"
            Height          =   255
            Index           =   2
            Left            =   2490
            TabIndex        =   29
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Clientes"
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   28
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Tarjetas"
            Height          =   255
            Index           =   1
            Left            =   1410
            TabIndex        =   27
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
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
         Index           =   8
         Left            =   2880
         TabIndex        =   44
         Top             =   360
         Width           =   960
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
         Left            =   4200
         TabIndex        =   43
         Top             =   360
         Width           =   840
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1380
         Picture         =   "frmFacturas.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   25
         Top             =   390
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1395
         MouseIcon       =   "frmFacturas.frx":0451
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1410
         MouseIcon       =   "frmFacturas.frx":05A3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Colectivo"
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
         Index           =   2
         Left            =   420
         TabIndex        =   24
         Top             =   2670
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   810
         TabIndex        =   23
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   22
         Top             =   2910
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   16
         Left            =   420
         TabIndex        =   19
         Top             =   1830
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   810
         TabIndex        =   18
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   810
         TabIndex        =   17
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1380
         Picture         =   "frmFacturas.frx":06F5
         ToolTipText     =   "Buscar fecha"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1380
         Picture         =   "frmFacturas.frx":0780
         ToolTipText     =   "Buscar fecha"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   780
         TabIndex        =   16
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   780
         TabIndex        =   15
         Top             =   1515
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   11
         Left            =   420
         TabIndex        =   14
         Top             =   900
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1380
         MouseIcon       =   "frmFacturas.frx":080B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1380
         MouseIcon       =   "frmFacturas.frx":095D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1530
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmCol As frmManCoope 'Colectivo
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBpr As frmManBanco 'Banco Propio
Attribute frmBpr.VB_VarHelpID = -1

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
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check1_Click()
Dim vCont As CContador

    Frame2.visible = (Check1.Value = 1)
    Frame3.visible = (Check1.Value = 1)
    If Check1.Value = 0 Then
        txtCodigo(7).Text = ""
        txtCodigo(8).Text = ""
        txtCodigo(9).Text = ""
    Else
        Set vCont = New CContador
        If vCont.LeerContador("FAC") Then
            txtCodigo(7).Text = vCont.Contador
        End If
        Set vCont = Nothing
    End If
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    cmdAceptar1_Click
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAceptar1_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim SQL As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Tipo As Byte
Dim NRegs As Integer
Dim NumError As Long
Dim db As BaseDatos
Dim TipoClien As String

Dim Sql5 As String


    If Not DatosOk Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    InicializarVbles
    MensError = ""
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H Colectivo
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{ssocio.codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHColec= """) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTabla = tabla & " INNER JOIN ssocio ON " & tabla & ".codsocio=ssocio.codsocio "
    
    SQL = "select count(*) "
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & "from ((scaalb inner join ssocio on scaalb.codsocio = ssocio.codsocio) "
        SQL = SQL & " inner join scoope on ssocio.codcoope = scoope.codcoope) "
        SQL = SQL & " inner join sartic on scaalb.codartic = sartic.codartic "
        
'[Monica]30/06/2014: ya no hay facturacion cepsa en la pobla del duc
'--        If Check1.Value = 1 Then
'            Sql = Sql & " and sartic.tipogaso > 0"
'        Else
'            Sql = Sql & " and sartic.tipogaso = 0"
'        End If
'++
        If Option1(5).Value Then SQL = SQL & " and sartic.tipogaso = 0"
        If Option1(6).Value Then SQL = SQL & " and sartic.tipogaso in (1,2,4)"
        If Option1(7).Value Then SQL = SQL & " and sartic.tipogaso = 3"
'++
    Else
        SQL = SQL & "from ((scaalb inner join ssocio on scaalb.codsocio = ssocio.codsocio) "
        SQL = SQL & " inner join scoope on ssocio.codcoope = scoope.codcoope) "
        
        '[Monica]19/06/2013: si es alzira hemos de diferenciar entre gasoleo B, domiciliado y los demas articulos
        If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2 Then
            SQL = SQL & " inner join sartic on scaalb.codartic = sartic.codartic "

        End If
        
    End If
'[Monica]07/03/2012: lo ponemos debajo pq la comprobacion de si esta traspasado lo de TPV es independiente de
'                    la cooperativa
'    If txtCodigo(4).Text <> "" Then sql = sql & " and ssocio.codcoope >= " & DBSet(txtCodigo(4).Text, "N")
'    If txtCodigo(5).Text <> "" Then sql = sql & " and ssocio.codcoope <= " & DBSet(txtCodigo(5).Text, "N")
    
    Tipo = 1
    If Option1(0).Value Then Tipo = 1
    If Option1(1).Value Then Tipo = 0
    If Option1(2).Value Then Tipo = 2
    
    '[Monica]04/03/2011: añadido el tema de facturacion interna (tipo 3)
    If Option1(3).Value Then Tipo = 3
    
    '[Monica]11/04/2016: interna de pobla del duc
    If Me.ChkInterna.Value Then Tipo = 3
    
    '[Monica]19/06/2013: añadido el tema de facturacion de gasoleo B
    If Tipo = 0 And Combo2.ListIndex > 0 Then Tipo = 4
    
    sql2 = SQL & " where scaalb.codforpa <> 98 "
    '[Monica]07/03/2012: lo dejo aqui.
    If txtCodigo(4).Text <> "" Then sql2 = sql2 & " and ssocio.codcoope >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then sql2 = sql2 & " and ssocio.codcoope <= " & DBSet(txtCodigo(5).Text, "N")
    
    If Tipo < 2 Or Tipo = 3 Then ' 04/03/2011: tipo de facturacion interna
        sql2 = sql2 & " and scoope.tipfactu = " & DBSet(Tipo, "N")
    Else
       'VRS:2.0.2(1) añadida nueva opción 2 facturacion ajena
       ' SQL = SQL & ")"
        sql2 = sql2 & " and (scoope.tipfactu = 0 or scoope.tipfactu = 1) "
    End If
    
    '[Monica]19/06/2013: añadimos la condicion
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2 Then
        Select Case Combo2.ListIndex
            Case 0
                sql2 = sql2 & " and not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                         "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3)"
            Case 1
                sql2 = sql2 & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0)"
            Case 2
                sql2 = sql2 & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1)"
        End Select
    End If
    
    SQL = SQL & " where scaalb.numfactu = 0 and scaalb.codforpa <>98 "
    
    If Tipo < 2 Or Tipo = 3 Then ' 04/03/2011: tipo de facturacion interna
        SQL = SQL & " and scoope.tipfactu = " & DBSet(Tipo, "N")
    Else
       'VRS:2.0.2(1) añadida nueva opción 2 facturacion ajena
       ' SQL = SQL & ")"
        SQL = SQL & " and (scoope.tipfactu = 0 or scoope.tipfactu = 1) "
    End If
    
    '[Monica]19/06/2013: añadimos la condicion
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2 Then
        Select Case Combo2.ListIndex
            Case 0
                SQL = SQL & " and not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                         "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3)"
            Case 1
                SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0)"
            Case 2
                SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1)"
        End Select
    End If
    
    Sql4 = SQL
    
    '[Monica]07/03/2012: lo dejo aqui.
    If txtCodigo(4).Text <> "" Then SQL = SQL & " and ssocio.codcoope >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then SQL = SQL & " and ssocio.codcoope <= " & DBSet(txtCodigo(5).Text, "N")
    
    
    Sql3 = " scaalb.codforpa <>98 "
    
    If Tipo < 2 Or Tipo = 3 Then ' 04/03/2011: tipo de facturacion interna
        Sql3 = Sql3 & " and scoope.tipfactu = " & DBSet(Tipo, "N")
    Else
       'VRS:2.0.2(1) añadida nueva opción 2 facturacion ajena
       ' SQL = SQL & ")"
        Sql3 = Sql3 & " and (scoope.tipfactu = 0 or scoope.tipfactu = 1) "
    End If
    
    '[Monica]19/06/2013: añadimos la condicion
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2 Then
        Select Case Combo2.ListIndex
            Case 0
                Sql3 = Sql3 & " and not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                         "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3)"
            Case 1
                Sql3 = Sql3 & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0)"
            Case 2
                Sql3 = Sql3 & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1)"
        End Select
    End If
    
    '[Monica]07/03/2012: dejo el enlace en las condiciones.
    If txtCodigo(4).Text <> "" Then Sql3 = Sql3 & " and ssocio.codcoope >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then Sql3 = Sql3 & " and ssocio.codcoope <= " & DBSet(txtCodigo(5).Text, "N")
    
    If txtCodigo(2).Text <> "" Then
        SQL = SQL & " and scaalb.fecalbar >= " & DBSet(txtCodigo(2).Text, "F") & " "
        sql2 = sql2 & " and scaalb.fecalbar >= " & DBSet(txtCodigo(2).Text, "F") & " "
        Sql3 = Sql3 & " and scaalb.fecalbar >= " & DBSet(txtCodigo(2).Text, "F") & " "
        Sql4 = Sql4 & " and scaalb.fecalbar >= " & DBSet(txtCodigo(2).Text, "F") & " "
    End If
    
    If txtCodigo(3).Text <> "" Then
        SQL = SQL & " and scaalb.fecalbar <= " & DBSet(txtCodigo(3).Text, "F") & " "
        sql2 = sql2 & " and scaalb.fecalbar <= " & DBSet(txtCodigo(3).Text, "F") & " "
        Sql3 = Sql3 & " and scaalb.fecalbar <= " & DBSet(txtCodigo(3).Text, "F") & " "
        Sql4 = Sql4 & " and scaalb.fecalbar <= " & DBSet(txtCodigo(3).Text, "F") & " "
    End If
    If txtCodigo(0).Text <> "" Then
        SQL = SQL & " and scaalb.codsocio >= " & DBSet(txtCodigo(0).Text, "N")
        sql2 = sql2 & " and scaalb.codsocio >= " & DBSet(txtCodigo(0).Text, "N")
        Sql3 = Sql3 & " and scaalb.codsocio >= " & DBSet(txtCodigo(0).Text, "N")
    End If
    If txtCodigo(1).Text <> "" Then
        SQL = SQL & " and scaalb.codsocio <= " & DBSet(txtCodigo(1).Text, "N")
        sql2 = sql2 & " and scaalb.codsocio <= " & DBSet(txtCodigo(1).Text, "N")
        Sql3 = Sql3 & " and scaalb.codsocio <= " & DBSet(txtCodigo(1).Text, "N")
    End If
    
    '[Monica]29/12/2016: para el caso de Ribarroja cogemos las entradas de contado o no
    If vParamAplic.Cooperativa = 5 Then
        If Me.ChkContado.Value = 1 Then ' contados
            SQL = SQL & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa = 0) "
            sql2 = sql2 & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa = 0) "
            Sql3 = Sql3 & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa = 0) "
            Sql4 = Sql4 & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa = 0) "
        Else ' no contados
            SQL = SQL & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa <> 0) "
            sql2 = sql2 & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa <> 0) "
            Sql3 = Sql3 & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa <> 0) "
            Sql4 = Sql4 & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa <> 0) "
        End If
    End If
    
    
    
    '[Monica]18/01/2013: condicion de tipo de socios
    Select Case Combo1.ListIndex
        Case 0
            TipoClien = "0"
        Case 1 ' clientes con bonificacion especial
            TipoClien = "1"
            
            SQL = SQL & " and ssocio.bonifesp = 1"
            sql2 = sql2 & " and ssocio.bonifesp = 1"
            Sql3 = Sql3 & " and ssocio.bonifesp = 1"
            Sql4 = Sql4 & " and ssocio.bonifesp = 1"
        
        Case 2 ' clientes sin bonificacion especial
            TipoClien = "2"
            
            SQL = SQL & " and ssocio.bonifesp = 0"
            sql2 = sql2 & " and ssocio.bonifesp = 0"
            Sql3 = Sql3 & " and ssocio.bonifesp = 0"
            Sql4 = Sql4 & " and ssocio.bonifesp = 0"
    End Select
    
    NRegs = TotalRegistros(SQL)
    
    If NRegs <> 0 Then
        '090908:comprobamos que existan todas las tarjetas en starjet
        If Not TarjetasInexistentes(sql2) Then
        
          '[Monica]04/03/2011: si hay articulos y la facturacion no es interna
          '                    añadida la condicion : and Tipo <> 3
          '[Monica]27/01/2010 comprobamos que todos los articulos a facturar tienen iva
          If Not ArticulosConIva(Sql3) And Tipo <> 3 Then
              MsgBox "Hay artículos sin Código de Iva. Revise.", vbExclamation
              Exit Sub
          End If
        
          '[Monica]25/01/2013: Comprobamos que no hay ningun albaran de importe superior a 2500
          If vParamAplic.Cooperativa = 1 And vParamAplic.LimiteFra <> 0 Then
              Sql5 = SQL & " and scaalb.importel > " & DBSet(vParamAplic.LimiteFra, "N") & " and ssocio.esdevarios = 0 and scaalb.codforpa in (select codforpa from sforpa where tipforpa = 0) "
              If TotalRegistros(Sql5) > 0 Then
                 MsgBox "Hay albaranes con importe superior a " & vParamAplic.LimiteFra & " euros. Revise.", vbExclamation
                 Exit Sub
              End If
          End If
        
          ' comprobamos si hay registros pendientes para pasar a tpv
          If Not PendientePasarTPV(Replace(Sql4, "scaalb.numfactu = 0", "scaalb.numfactu <> 0"), Tipo) Then
              If Not PendienteCierresTurno(Trim(txtCodigo(2).Text), Trim(txtCodigo(3).Text)) Then
                    On Error GoTo eError
                    Pb1.visible = True
                    CargarProgres Pb1, NRegs
                    
                    Set db = New BaseDatos
                    db.abrir vSesion.CadenaConexion, "root", "aritel"
                    db.Tipo = "MYSQL"
                    db.AbrirTrans
                    If Check1.Value = 1 Then ConnConta.BeginTrans
                    ' ### [Monica] 31/05/2007
                    ' previo cambiamos la forma de pago de aquellos albaranes que no sean
                    ' de efectivo cuyo cliente tenga facturafp = si (tenga que facturar con
                    ' la fp de la ficha del cliente
                    NumError = 0
                    NumError = Prefacturacion(db, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text, TipoClien)
                    
                    
                    ' Facturacion por cliente ambas o interna                                           '[Monica]15/07/2013: o es normal o es de gasoleo bonif interna
                    If (Option1(0).Value Or Option1(2).Value Or Option1(3).Value) And NumError = 0 And (Combo2.ListIndex = 0 Or (Combo2.ListIndex > 0 And Option1(3).Value)) Then
                        '[Monica]30/06/2014: pobla del duc pasa a no tener facturacion cepsa
                        '--
                        'If Check1.Value = 1 Then
                        '    NumError = FacturacionCepsa(db, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text, CDate(txtCodigo(6).Text), 1, Pb1, txtCodigo(9).Text, txtCodigo(8).Text)
                        '++
                        '[Monica]06/08/2014: faltaba la condicion de que sea la cooperativa de Pobla del Duc
                        If (Option1(5).Value Or Option1(6).Value Or Option1(7).Value) And vParamAplic.Cooperativa = 4 Then
                            Dim TipoArt As Integer
                            If Option1(5).Value Then TipoArt = 0
                            If Option1(6).Value Then TipoArt = 1
                            If Option1(7).Value Then TipoArt = 2
                            '[Monica]11/04/2016: facturas internas para el caso de pobla del duc
                            Dim tipoF As Byte
                            tipoF = 1
                            If Me.ChkInterna.Value Then tipoF = 3                                                                                                                                 ' antes tipof era 1 fijo
                            NumError = Facturacion(db, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text, CDate(txtCodigo(6).Text), tipoF, Pb1, TipoClien, 0, TipoArt)
                        '++
                        Else
                            ' Interna
                            If Option1(3).Value Then
                                NumError = Facturacion(db, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text, CDate(txtCodigo(6).Text), 3, Pb1, TipoClien, Combo2.ListIndex, , (Me.ChkContado.Value = 1)) '[Monica]15/07/2013:antes 0
                            Else
                                NumError = Facturacion(db, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text, CDate(txtCodigo(6).Text), 1, Pb1, TipoClien, 0, , (Me.ChkContado.Value = 1))
                            End If
                        End If
                    End If
                       
                    ' facturacion por tarjeta
                    If (Option1(1).Value Or Option1(2).Value) And NumError = 0 Then
                       NumError = Facturacion(db, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text, CDate(txtCodigo(6).Text), 0, Pb1, TipoClien, Combo2.ListIndex, , (Me.ChkContado.Value = 1))
                    End If
              Else
                 Exit Sub
              End If
          Else
            Exit Sub
          End If
       Else 'no existen todas las tarjetas
           Exit Sub
       End If
    Else
        MsgBox "No hay registros a procesar.", vbExclamation
        Exit Sub
    End If
'    db.CommitTrans
'    MsgBox "Proceso finalizado correctamente", vbExclamation
'    Set db = Nothing
'    Pb1.visible = False
'    Exit Sub

eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de facturación. Llame a soporte." & vbCrLf & vbCrLf & _
                    MensError
        db.RollbackTrans
        If Check1.Value = 1 Then ConnConta.RollbackTrans
        Set db = Nothing
        Pb1.visible = False
    Else
        db.CommitTrans
        If Check1.Value = 1 Then ConnConta.CommitTrans
        MsgBox "Proceso finalizado correctamente", vbExclamation
        Set db = Nothing
        Pb1.visible = False
        cmdCancel_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo2_LostFocus()
    Select Case Combo2.ListIndex
        Case 0
            Frame4.visible = True
            Frame4.Enabled = True
'[Monica]15/07/2013: añadido esto
            Option1(0).Enabled = True
            Option1(2).Enabled = True
            Option1(0).Value = True
        Case 1
'            Frame4.Enabled = False
'            Option1(1).Value = True
'[Monica]15/07/2013: cambiado por esto
            Frame4.visible = True
            Frame4.Enabled = True
            Option1(0).Enabled = False
            Option1(2).Enabled = False
            Option1(1).Value = True
        Case 2
'            Frame4.Enabled = False
'            Option1(1).Value = True
'[Monica]15/07/2013: cambiado por esto
            Frame4.visible = True
            Frame4.Enabled = True
            Option1(0).Enabled = False
            Option1(2).Enabled = False
            Option1(1).Value = True
    End Select

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonerFoco txtCodigo(6)
        
        Me.Combo1.ListIndex = 0
        Me.Combo2.ListIndex = 0
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "scaalb"
    
    Pb1.visible = False
    
    CargaCombo
    
    Combo2.Enabled = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2)
    Combo2.visible = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2)
    Label4(8).visible = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2)
    
    Frame4.visible = (vParamAplic.Cooperativa <> 4)
    Frame4.Enabled = (vParamAplic.Cooperativa <> 4)
    
    '[Monica]30/06/2014: cambiamos la facturacion de pobla del duc, ya no hay facturacion cepsa
    Frame1.visible = False '(vParamAplic.Cooperativa = 4)
    Frame1.Enabled = False '(vParamAplic.Cooperativa = 4)
    
    Frame5.visible = (vParamAplic.Cooperativa = 4)
    Frame5.Enabled = (vParamAplic.Cooperativa = 4)
    
    '[Monica]11/04/2016:
    ChkInterna.Enabled = (vParamAplic.Cooperativa = 4)
    ChkInterna.visible = (vParamAplic.Cooperativa = 4)
    
    
    '[Monica]18/01/2013: Obligamos a Ribarroja a introducir el colectivo
    Label4(0).visible = (vParamAplic.Cooperativa <> 5)
    Label4(1).visible = (vParamAplic.Cooperativa <> 5)
    imgBuscar(5).Enabled = (vParamAplic.Cooperativa <> 5)
    imgBuscar(5).visible = (vParamAplic.Cooperativa <> 5)
    txtCodigo(5).Enabled = (vParamAplic.Cooperativa <> 5)
    txtCodigo(5).visible = (vParamAplic.Cooperativa <> 5)
    txtNombre(5).visible = (vParamAplic.Cooperativa <> 5)
    
    '[Monica]29/12/2016: Ribarroja Factura con distinto contador dependiendo de si es o no contado
    Me.ChkContado.Enabled = (vParamAplic.Cooperativa = 5)
    Me.ChkContado.visible = (vParamAplic.Cooperativa = 5)
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmBpr_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("FACTURAC")
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFec_Click(Index As Integer)
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
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLIENTE
            AbrirfrmClientes (Index)
        
        Case 4, 5 'COLECTIVO
            AbrirFrmColectivo (Index)
        
        Case 8 ' BANCO PROPIO
            AbrirFrmBancoPropio (Index)

    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.Caption = "Facturación por Cliente"
        Case 1
            Me.Caption = "Facturación por Tarjeta"
        Case 2
            Me.Caption = "Facturación por Cliente y por Tarjeta"
    End Select
    
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'   KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 4: KEYBusqueda KeyAscii, 4 'colectivo desde
            Case 5: KEYBusqueda KeyAscii, 5 'colectivo hasta
            Case 6: KEYFecha KeyAscii, 6 'fecha factura
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
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

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "ssocio", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3, 6, 7, 9 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 4, 5 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 8 ' BANCO PROPIO
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sbanco", "nombanco", "codbanpr", "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "El Banco introducido no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(Index)
                Else
                    cmdAceptar.SetFocus
                End If
            End If
    
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtCodigo(6).Text = Format(Now, "dd/mm/yyyy")
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirfrmClientes(indice As Integer)
    indCodigo = indice
    Set frmcli = New frmManClien
    frmcli.DatosADevolverBusqueda = "0|1|"
    frmcli.DeConsulta = True
    frmcli.CodigoActual = txtCodigo(indCodigo)
    frmcli.Show vbModal
    Set frmcli = Nothing
End Sub

Private Sub AbrirFrmColectivo(indice As Integer)
    indCodigo = indice
    Set frmCol = New frmManCoope
    frmCol.DatosADevolverBusqueda = "0|1|"
    frmCol.DeConsulta = True
    frmCol.CodigoActual = txtCodigo(indCodigo)
    frmCol.Show vbModal
    Set frmCol = Nothing
End Sub

Private Sub AbrirFrmBancoPropio(indice As Integer)
    indCodigo = indice
    Set frmBpr = New frmManBanco
    frmBpr.DatosADevolverBusqueda = "0|1|"
    frmBpr.DeConsulta = True
    frmBpr.CodigoActual = txtCodigo(indCodigo)
    frmBpr.Show vbModal
    Set frmBpr = Nothing
End Sub



Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
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
        .Opcion = ""
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

Private Function PendientePasarTPV(SQL As String, Tipo As Byte) As Boolean
'Dim sql As String
Dim cadMen As String

    PendientePasarTPV = False
'    sql = "select count(*) from scaalb, ssocio, scoope where " & Formula & " and numfactu <> 0 and " & _
'          " scaalb.codsocio = ssocio.codsocio and ssocio.codcoope = scoope.codcoope "
    
    If Tipo <> 2 Then
        SQL = SQL & " and scoope.tipfactu = " & DBSet(Tipo, "N")
    Else 'VRS:2.0.2(1) añadida nueva opción
        SQL = SQL & " and (scoope.tipfactu = 0 or scoope.tipfactu = 1)"
    End If
    
    If (RegistrosAListar(SQL) <> 0) Then
        cadMen = "Hay registros pendientes de Traspaso a TPV." & vbCrLf & vbCrLf & _
                 "Debe realizar este proceso previamente." & vbCrLf & vbCrLf
        MsgBox cadMen, vbExclamation
        PendientePasarTPV = True
    End If
End Function

Private Function PendienteCierresTurno(DesFec As String, HasFec As String) As Boolean
Dim SQL As String
Dim cadMen As String

    PendienteCierresTurno = False
    
'   [Monica] 27/01/2010 si es castelduc no cierra turno y pobla del duc tampoco
    If vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 4 Then Exit Function
    
    
    
    SQL = "select count(*) from srecau where intconta = 0 "
    If DesFec <> "" Then SQL = SQL & " and fechatur >= " & DBSet(CDate(DesFec), "F") & " "
    If HasFec <> "" Then SQL = SQL & " and fechatur <= " & DBSet(CDate(HasFec), "F") & " "

    If (RegistrosAListar(SQL) <> 0) Then
        cadMen = "Quedan cierres de Turno por contabilizar. Revise." & vbCrLf & vbCrLf
        MsgBox cadMen, vbExclamation
        PendienteCierresTurno = True
    End If
    
End Function

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim gdgp As GestorDeclaracionesGasoleoProf
Dim Fecha As Date
    b = True

    '[Monica]18/01/2013: en Ribarroja obligamos a que me introduzcan un colectivo y que exista
    If vParamAplic.Cooperativa = 5 Then
        If txtCodigo(4).Text = "" Then
            MsgBox "Debe introducir obligatoriamente un colectivo. Revise.", vbExclamation
            PonerFoco txtCodigo(4)
            DatosOk = False
            Exit Function
        Else
            SQL = DevuelveValor("select count(*) from scoope where codcoope = " & DBSet(txtCodigo(4).Text, "N"))
            If SQL = "0" Then
                MsgBox "No existe el colectivo. Revise.", vbExclamation
                PonerFoco txtCodigo(4)
                DatosOk = False
                Exit Function
            Else
                txtCodigo(5).Text = txtCodigo(4).Text
            End If
        End If
    End If
    
    If txtCodigo(6).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Facturación.", vbExclamation
        b = False
        PonerFoco txtCodigo(6)
    Else
        ' Cooperativa de la Pobla del Duc
        If vParamAplic.Cooperativa = 4 Then
            If Check1.Value = 1 Then
                If txtCodigo(7).Text = "" Then
                    Mens = "Debe introducir el primer Nro.Factura de Cepsa. Revise"
                    MsgBox Mens, vbExclamation
                    b = False
                    PonerFoco txtCodigo(7)
                End If
                If b Then
                    If txtCodigo(9).Text = "" Then
                        Mens = "Debe introducir la Fecha de Vencimiento del pago. Revise"
                        MsgBox Mens, vbExclamation
                        b = False
                        PonerFoco txtCodigo(9)
                    End If
                End If
                If b Then
                    If txtCodigo(8).Text = "" Then
                        Mens = "Debe introducir el Banco del pago. Revise"
                        MsgBox Mens, vbExclamation
                        b = False
                        PonerFoco txtCodigo(8)
                    End If
                End If
            Else
                If Not FechaDentroPeriodoContable(CDate(txtCodigo(6).Text)) Then
                    Mens = "La Fecha de Facturación no es del ejercicio actual ni siguiente. Reintroduzca."
                    MsgBox Mens, vbExclamation
                    b = False
                    PonerFoco txtCodigo(6)
                Else
                    'VRS:2.0.1(0)
                    If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(6).Text)) Then
                        Mens = "  La Fecha de Facturación es inferior a la última liquidación de Iva. " & vbCrLf & vbCrLf
                        ' unicamente si el usuario es root el proceso continuará
                        If vSesion.Nivel > 0 Then
                            Mens = Mens & "  El proceso no continuará."
                            MsgBox Mens, vbExclamation
                            b = False
                            PonerFoco txtCodigo(6)
                        Else
                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                b = False
                                PonerFoco txtCodigo(6)
                            End If
                        End If
                    End If
                    ' la fecha de factura no debe ser inferior a la ultima factura de la serie
                    numser = "letraser"
                    numfactu = ""
                    numfactu = DevuelveDesdeBDNew(cPTours, "stipom", "contador", "codtipom", "FAG", "T", numser)
                    If numfactu <> "" Then
                        If FechaFacturaInferiorUltimaFacturaSerieHco(CDate(txtCodigo(6).Text), CLng(numfactu), numser, 0) Then
                            Mens = "La Fecha de Factura es inferior a la última factura de la serie." & vbCrLf & vbCrLf
                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                b = False
                                PonerFoco txtCodigo(6)
                            End If
                        End If
                    End If
                End If
            End If
        Else
        ' cooperativa distinta a la Pobla del DUC
        
            Set gdgp = New GestorDeclaracionesGasoleoProf
            If txtCodigo(3).Text <> "" Then
                Fecha = CDate(txtCodigo(3).Text)
            Else
                Fecha = Now
            End If
            
            If gdgp.quedaPorDeclarar(Conn, Fecha) Then
                Mens = "Quedan albaranes de gasóleo profesional pendientes de declarar. Revise"
                MsgBox Mens, vbExclamation
                b = False
                PonerFoco txtCodigo(6)
            Else
                If Not FechaDentroPeriodoContable(CDate(txtCodigo(6).Text)) Then
                    Mens = "La Fecha de Facturación no es del ejercicio actual ni siguiente. Reintroduzca."
                    MsgBox Mens, vbExclamation
                    b = False
                    PonerFoco txtCodigo(6)
                Else
                    'VRS:2.0.1(0)
                    If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(6).Text)) Then
                        Mens = "  La Fecha de Facturación es inferior a la última liquidación de Iva. " & vbCrLf & vbCrLf
                        ' unicamente si el usuario es root el proceso continuará
                        If vSesion.Nivel > 0 Then
                            Mens = Mens & "  El proceso no continuará."
                            MsgBox Mens, vbExclamation
                            b = False
                            PonerFoco txtCodigo(6)
                        Else
                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                b = False
                                PonerFoco txtCodigo(6)
                            End If
                        End If
                    End If
                    ' la fecha de factura no debe ser inferior a la ultima factura de la serie
                    numser = "letraser"
                    numfactu = ""
                    numfactu = DevuelveDesdeBDNew(cPTours, "stipom", "contador", "codtipom", "FAG", "T", numser)
                    If numfactu <> "" Then
                        If FechaFacturaInferiorUltimaFacturaSerieHco(CDate(txtCodigo(6).Text), CLng(numfactu), numser, 0) Then
                            Mens = "La Fecha de Factura es inferior a la última factura de la serie." & vbCrLf & vbCrLf
                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                                b = False
                                PonerFoco txtCodigo(6)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    'comprobamos que haya articulo descuento en la tabla de parámetros.
    If vParamAplic.ArticDto = 0 Then
        MsgBox "Debe introducir un articulo de descuento en la tabla de parámetros. Revise.", vbExclamation
        b = False
        PonerFocoBtn CmdCancel
    Else
        'comprobamos que el articulo de descuento existe
        SQL = ""
        SQL = DevuelveDesdeBD("nomartic", "sartic", "codartic", vParamAplic.ArticDto, "N")
        If SQL = "" Then
            MsgBox "El artículo descuento de la tabla de parámetros no existe. Revise.", vbExclamation
            b = False
            PonerFocoBtn CmdCancel
        End If
    End If

    DatosOk = b
End Function


'Private Function TarjetasInexistentes(Sql As String) As Boolean
'Dim cadMen As String
'
'    TarjetasInexistentes = False
'
'    Sql = Sql & " and not (scaalb.codsocio, scaalb.numtarje) in (select codsocio, numtarje from starje) "
'
'    If (RegistrosAListar(Sql) <> 0) Then
'        cadMen = "Hay cargas en las que no es correcta la tarjeta para el socio." & vbCrLf & vbCrLf & _
'                 "Revise en el mantenimiento de albaranes." & vbCrLf & vbCrLf
'        MsgBox cadMen, vbExclamation
'        TarjetasInexistentes = True
'    End If
'End Function

Private Function ArticulosConIva(vSQL As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim sql2 As String
Dim b As Boolean


    ArticulosConIva = False

    SQL = "select distinct codigiva from sartic, scaalb, ssocio, scoope where sartic.codartic = scaalb.codartic "
    SQL = SQL & " and scaalb.codsocio = ssocio.codsocio and ssocio.codcoope = scoope.codcoope "
    If vSQL <> "" Then SQL = SQL & " and " & vSQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    
    While Not Rs.EOF And b
        sql2 = ""
        sql2 = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", DBLet(Rs!CodigIVA, "N"), "N")
        If sql2 = "" Then b = False
        
        Rs.MoveNext
    Wend
    
    ArticulosConIva = b

End Function
    
    

Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Combo1.AddItem "Todos"
    Combo1.ItemData(Combo1.NewIndex) = 0
    
    Combo1.AddItem "Con Bonif.Esp."
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "Sin Bonif."
    Combo1.ItemData(Combo1.NewIndex) = 2


    Combo2.Clear
    'Tipos de facturacion
    Combo2.AddItem "Normal"
    Combo2.ItemData(Combo2.NewIndex) = 0
    
    Combo2.AddItem "Gas.B"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    ' 19/06/2013: solo en el caso de que sea Alzira hay gasoleo B domiciliado ( de momento )
    If vParamAplic.Cooperativa = 1 Then
        Combo2.AddItem "Gas.B Dom."
        Combo2.ItemData(Combo2.NewIndex) = 2
    End If

End Sub

