VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmManClien2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarjetas"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   6930
   Icon            =   "frmManClien2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDatosEmpleado 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   90
      TabIndex        =   27
      Top             =   390
      Width           =   6705
      Begin VB.Frame Frame4 
         Caption         =   "Códigos DIR"
         ForeColor       =   &H00972E0B&
         Height          =   1560
         Left            =   225
         TabIndex        =   52
         Top             =   3510
         Width           =   6285
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   20
            Left            =   1755
            MaxLength       =   255
            TabIndex        =   16
            Tag             =   "Oficina Contable|T|S|||starje|oficinacontable||N|"
            Text            =   "Text1"
            Top             =   855
            Width           =   4230
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   19
            Left            =   1755
            MaxLength       =   255
            TabIndex        =   15
            Tag             =   "Unidad Tramitadora|T|S|||starje|unidadtramitadora||N|"
            Text            =   "Text1"
            Top             =   540
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   18
            Left            =   1755
            MaxLength       =   255
            TabIndex        =   14
            Tag             =   "Organo gestor|T|S|||starje|organogestor||N|"
            Top             =   225
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   21
            Left            =   1755
            MaxLength       =   255
            TabIndex        =   17
            Tag             =   "Organo proponente|T|S|||starje|orgproponente||N|"
            Text            =   "Text1"
            Top             =   1170
            Width           =   4230
         End
         Begin VB.Label Label1 
            Caption         =   "Oficina Contable"
            Height          =   195
            Index           =   67
            Left            =   135
            TabIndex        =   56
            Top             =   900
            Width           =   2235
         End
         Begin VB.Label Label1 
            Caption         =   "Unidad Tramitadora"
            Height          =   195
            Index           =   65
            Left            =   135
            TabIndex        =   55
            Top             =   585
            Width           =   2565
         End
         Begin VB.Label Label1 
            Caption         =   "Órgano Gestor "
            Height          =   255
            Index           =   63
            Left            =   135
            TabIndex        =   54
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Órgano Proponente"
            Height          =   285
            Index           =   64
            Left            =   135
            TabIndex        =   53
            Top             =   1215
            Width           =   2235
         End
      End
      Begin VB.Frame Frame2 
         Height          =   885
         Index           =   0
         Left            =   210
         TabIndex        =   36
         Top             =   -30
         Width           =   6255
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   48
            Top             =   390
            Width           =   4725
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   6
            TabIndex        =   1
            Tag             =   "Código de cliente|N|N|1|999999|starje|codsocio|000000|S|"
            Top             =   400
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   37
            Top             =   195
            Width           =   675
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5295
         TabIndex        =   26
         Top             =   8355
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   5310
         TabIndex        =   40
         Top             =   8325
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4125
         TabIndex        =   25
         Top             =   8355
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Index           =   1
         Left            =   210
         TabIndex        =   38
         Top             =   8205
         Width           =   2385
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
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
            Left            =   45
            TabIndex        =   39
            Top             =   210
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos básicos"
         ForeColor       =   &H00972E0B&
         Height          =   2490
         Index           =   1
         Left            =   210
         TabIndex        =   32
         Top             =   930
         Width           =   6285
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   945
            MaxLength       =   6
            TabIndex        =   13
            Tag             =   "Código departamento|N|N|0|9999|starje|coddepar|0000||"
            Top             =   2070
            Width           =   825
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   1800
            TabIndex        =   57
            Top             =   2070
            Width           =   4050
         End
         Begin VB.ComboBox cmbAux 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            ItemData        =   "frmManClien2.frx":000C
            Left            =   4230
            List            =   "frmManClien2.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "Estado|N|N|||starje|estado|||"
            Top             =   1680
            Width           =   1665
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   17
            Left            =   2700
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Fecha alta|F|S|||starje|fecalta|dd/mm/yyyy||"
            Top             =   1710
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   960
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Matricula|T|S|||starje|matricul|||"
            Top             =   1710
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   4200
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Número de Cuenta|T|S|||starje|cuentaba|0000000000||"
            Top             =   1080
            Width           =   1755
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Dígito de Control|T|S|||starje|digcontr|00||"
            Top             =   1080
            Width           =   465
         End
         Begin VB.ComboBox cmbAux 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            ItemData        =   "frmManClien2.frx":0010
            Left            =   4260
            List            =   "frmManClien2.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "Tipo Tarjeta|N|N|||starje|tiptarje|||"
            Top             =   270
            Width           =   1665
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   960
            MaxLength       =   9
            TabIndex        =   2
            Tag             =   "Tarjeta|N|N|1|99999999|starje|numtarje|00000000||"
            Top             =   330
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   960
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "Titular|T|S|||starje|nomtarje|||"
            Top             =   720
            Width           =   4995
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "Banco|T|S|0|9999|starje|codbanco|0000||"
            Top             =   1080
            Width           =   825
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   960
            MaxLength       =   6
            TabIndex        =   5
            Tag             =   "IBAN|T|S|||starje|iban|||"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   2730
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "Oficina|T|S|0|9999|starje|codsucur|0000||"
            Top             =   1080
            Width           =   765
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   675
            ToolTipText     =   "Buscar Departamento"
            Top             =   2070
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Dpto"
            Height          =   255
            Index           =   14
            Left            =   180
            TabIndex        =   58
            Top             =   2070
            Width           =   450
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   15
            Left            =   3510
            Picture         =   "frmManClien2.frx":0014
            ToolTipText     =   "Buscar fecha"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Estado"
            Height          =   255
            Index           =   11
            Left            =   4230
            TabIndex        =   51
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Alta"
            Height          =   255
            Index           =   1
            Left            =   2700
            TabIndex        =   50
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Matrícula"
            Height          =   255
            Index           =   4
            Left            =   960
            TabIndex        =   42
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Tarjeta"
            Height          =   255
            Index           =   5
            Left            =   3210
            TabIndex        =   41
            Top             =   300
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Tarjeta"
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   35
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   2
            Left            =   195
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   33
            Top             =   1080
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos impresión Tarjeta"
         ForeColor       =   &H00972E0B&
         Height          =   3000
         Index           =   2
         Left            =   225
         TabIndex        =   28
         Top             =   5175
         Width           =   6285
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   16
            Left            =   990
            MaxLength       =   50
            TabIndex        =   24
            Tag             =   "Banda 3|T|S|||starje|banda3|||"
            Top             =   2580
            Width           =   4965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   15
            Left            =   990
            MaxLength       =   50
            TabIndex        =   23
            Tag             =   "Banda 2|T|S|||starje|banda2|||"
            Top             =   2190
            Width           =   4965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   990
            MaxLength       =   50
            TabIndex        =   22
            Tag             =   "Banda 1|T|S|||starje|banda1|||"
            Top             =   1830
            Width           =   4965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   990
            MaxLength       =   50
            TabIndex        =   20
            Tag             =   "Linea 3|T|S|||starje|linea3|||"
            Top             =   1050
            Width           =   4965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   13
            Left            =   990
            MaxLength       =   50
            TabIndex        =   21
            Tag             =   "Cod.Barras|T|S|||starje|codbarras|||"
            Top             =   1440
            Width           =   4965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   990
            MaxLength       =   50
            TabIndex        =   18
            Tag             =   "Linea 1|T|S|||starje|linea1|||"
            Top             =   300
            Width           =   4965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   990
            MaxLength       =   50
            TabIndex        =   19
            Tag             =   "Linea 2|T|S|||starje|linea2|||"
            Top             =   690
            Width           =   4965
         End
         Begin VB.Label Label1 
            Caption         =   "Banda 3"
            Height          =   255
            Index           =   10
            Left            =   150
            TabIndex        =   46
            Top             =   2580
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Banda 2"
            Height          =   255
            Index           =   9
            Left            =   150
            TabIndex        =   45
            Top             =   2190
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Banda 1"
            Height          =   255
            Index           =   8
            Left            =   150
            TabIndex        =   44
            Top             =   1830
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Cód.Barras"
            Height          =   255
            Index           =   6
            Left            =   150
            TabIndex        =   43
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Línea 2"
            Height          =   255
            Index           =   12
            Left            =   150
            TabIndex        =   31
            Top             =   705
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Línea 1"
            Height          =   255
            Index           =   13
            Left            =   150
            TabIndex        =   30
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Línea 3 "
            Height          =   255
            Index           =   16
            Left            =   150
            TabIndex        =   29
            Top             =   1095
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   47
         Tag             =   "Número de línea|N|N|1|9999|starje|numlinea|0000|S|"
         Top             =   420
         Width           =   885
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Enabled         =   0   'False
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5670
         TabIndex        =   49
         Top             =   30
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManClien2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO   +-+-
' +-+- Fecha: 23/05/06 +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public DeConsulta As Boolean
Public CodigoActual As String

Public ModoExt As Integer
Public socio As String
Public NumLin As String


Private HaDevueltoDatos As Boolean
Private CadenaSelect As String
Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmDep As frmManDpto
Attribute frmDep.VB_VarHelpID = -1


Dim Modo As Byte
'-------------- MODOS ---------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'------------------------------------------------
Dim FormatoCod As String 'formato del campo código
Dim NomTabla As String
Dim Ordenacion As String
Dim PrimeraVez As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos
Dim SQL As String
Dim miRsAux As ADODB.Recordset
Dim I As Integer


Private Sub PonerModo(vModo)
Dim b As Boolean
Dim Numreg As Byte

    On Error GoTo EPonerModo
    
    Modo = vModo
    If Modo = 2 Then
        lblIndicador.Caption = PonerContRegistros(Me.adodc1)
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    
    b = (Modo = 2)
    
    '=======================================
    'Poner Flechas de desplazamiento visibles
    Numreg = 1
    If Not Me.adodc1.Recordset.EOF Then
        If adodc1.Recordset.RecordCount > 1 Then Numreg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    
     '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    cmdRegresar.visible = Not b
    
    
    BloquearText1 Me, Modo
    BloquearCmb Me.cmbAux(0), Modo = 0 Or Modo = 2
    BloquearCmb Me.cmbAux(1), Modo = 0 Or Modo = 2
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = Not b
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos 'Pone el Maxlength de los campos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner modo.", Err.Description
End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2) Or Modo = 0
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
    Me.mnImprimir.Enabled = b

End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    
    LimpiarCampos 'Vacía los TextBox
    CadB = ""
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    '******************** canviar taula i camp **************************
    
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    'PosarDescripcions
    Text1(0).Text = socio
    Text1(1).Text = NumLin
    
    '[Monica]20/06/2014: ofertamos el maximo numero de tarjeta
    Text1(2).Text = Format(SugerirCodigoSiguienteStr("starje", "numtarje", "numtarje < 20000"), "00000000")
    
    text2(0).Text = NuevoCodigo
    Text1(8).Text = NuevoCodigo
    
    
    Text1(3).Text = text2(0).Text
    Text1(17).Text = Format(Now, "dd/mm/yyyy")
    cmbAux(0).ListIndex = 1
    cmbAux(1).ListIndex = 0
    
    
    SQL = "select iban, codbanco, codsucur, digcontr, cuentaba from ssocio where codsocio = " & DBSet(Text1(0).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text1(4).Text = DBLet(Rs!IBAN, "T")
        Text1(5).Text = Format(DBLet(Rs!codbanco, "N"), "0000")
        Text1(6).Text = Format(DBLet(Rs!codsucur, "N"), "0000")
        Text1(7).Text = DBLet(Rs!digcontr, "T")
        Text1(11).Text = DBLet(Rs!cuentaba, "T")
    End If
    
    Set Rs = Nothing
    
    PonerFoco Text1(2)
    ' ********************************************************************
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    LimpiarCampos 'Limpia los Text1
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NomTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String

        'Llamamos a al form
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 20, "Cód.")
        Cad = Cad & ParaGrid(Text1(1), 80, "Nombre")
        
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NomTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Trabajadores"
            frmB.vSelElem = 0

            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Me.adodc1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(1)
            End If
        End If
        ' *************************************************************************
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Me.adodc1.RecordSource = CadenaConsulta
    adodc1.Refresh
    If adodc1.Recordset.RecordCount <= 0 Then
        If CadB = "" Then
            MsgBox "No hay ningún registro en la tabla " & NomTabla, vbInformation
'            Screen.MousePointer = vbDefault
'            Exit Sub
        Else
            If Modo = 1 Then MsgBox "Ningún registro encontrado para el criterio de búsqueda.", vbInformation
            PonerFoco Text1(indice)
        End If
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        adodc1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonBuscar()
   If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0)
'        PosicionarCombo Combo1(0), 754
        Text1(0).BackColor = vbYellow
    End If
End Sub

Private Sub BotonModificar()
    
    PonerModo 4
   
    'Como es modificar
    ' *** primer control que no siga clau primaria ***
    PonerFoco Text1(2)
    ' ************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonEliminar()
Dim SQL As String

    On Error GoTo EEliminar
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'if EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el Trabajador?"
    SQL = SQL & vbCrLf & "Código: " & Text1(0).Text
    SQL = SQL & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        SQL = "Delete from " & NomTabla & " where codtraba=" & adodc1.Recordset!CodTraba
        Conn.Execute SQL
        
        If SituarDataTrasEliminar(adodc1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()

    Select Case Modo
         Case 1  'BUSQUEDA
            HacerBusqueda
    
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    TerminaBloquear
                    Unload Me
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    Unload Me
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            If Me.adodc1.Recordset.EOF Then
                PonerModo 0
                Unload Me
                
            Else
                PonerModo 2
                PonerCampos
            End If
            PonerFoco Text1(0)

        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    Cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        
        Modo = ModoExt
        Select Case Modo
            Case 0
                DatosADevolverBusqueda = "ZZ"
                PonerModo Modo
            Case 3
                mnNuevo_Click
            Case 4
                mnModificar_Click
        End Select
    End If

End Sub

Private Sub Form_Load()
    
    PrimeraVez = True
    
    ' ICONITOS DE LA BARRA
    btnPrimero = 15 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        '14 y 15 separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With

    Me.Toolbar1.visible = False
    Me.Toolbar1.Enabled = False
    
    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    

    CargaCombo

    LimpiarCampos   'Limpia los campos TextBox
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    

    '****************** canviar la consulta *********************************+
    NomTabla = "starje"
    Ordenacion = " ORDER BY numlinea"
    CadenaConsulta = "select * from " & NomTabla
    
    Me.adodc1.ConnectionString = Conn
    Me.adodc1.RecordSource = CadenaConsulta & " where codsocio= " & DBSet(socio, "N") & " and numlinea=" & DBSet(NumLin, "N")
    Me.adodc1.Refresh
    
    CadB = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        ' *** canviar o llevar el WHERE ***
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub imgMail_Click(Index As Integer)
    If Index = 0 Then
        If Text1(9).Text <> "" Then
            LanzaMailGnral Text1(9).Text
        End If
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(CByte(imgFec(15).Tag) + 2).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmDep_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(22).Text = RecuperaValor(CadenaSeleccion, 1)
        text2(22).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
     Select Case Index
        Case 0 'departamentos
            Set frmDep = New frmManDpto
            frmDep.DatosADevolverBusqueda = "0|1|"
            frmDep.CodigoActual = Text1(22).Text
            frmDep.Show vbModal
            Set frmDep = Nothing
            PonerFoco Text1(22)
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(15).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(17).Text <> "" Then frmC.NovaData = Text1(17).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(15).Tag) + 2) '<===
    ' ********************************************

End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub



Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    
    TerminaBloquear
    
    PonerModo 2
    text2(0).Text = NuevoCodigo
    PonerCampos
    
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    indice = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 17: KEYFecha KeyAscii, 15 'fecha alta
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Nuevo As Boolean
Dim cadMen As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 2
            If Text1(2).Text <> "" Then
                Text1(14).Text = "P000000208" & Format(Text1(2).Text, "00000") & "^0000" ' banda1
                Text1(15).Text = "9724000030" & Format(Text1(2).Text, "000000") ' banda2
            End If
        
        Case 3, 12 ' Nomtarje y matricula
            Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 5, 6 ' banco y sucursal con 4 digitos
            Text1(Index).Text = Format(Text1(Index).Text, "0000")
    
        Case 11 ' cuenta de banco con 10 digitos
            PonerFormatoEntero Text1(Index)
            
        Case 17 ' fecha de alta
            PonerFormatoFecha Text1(Index)
            
        '[Monica]03/05/2019: departamento al que pertenece la tarjeta
        Case 22 ' departamento
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index) = PonerNombreDeCod(Text1(Index), "departamento", "nomdepar", "coddepar", "N")
                If text2(Index).Text = "" Then
                    cadMen = "No existe el Departamento: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmDep = New frmManDpto
                        frmDep.DatosADevolverBusqueda = "0|1|"
                        frmDep.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmDep.Show vbModal
                        Set frmDep = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
        
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                BotonBuscar
        Case 3
                BotonVerTodos
        Case 6
                BotonAnyadir
        Case 7
                mnModificar_Click
        Case 8
                BotonEliminar
        Case 13 'Salir
                mnSalir_Click
                
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Me.adodc1.Recordset.EOF Then Exit Sub
    DesplazamientoData adodc1, Index
    PonerCampos
End Sub

Private Sub PonerCampos()

    If adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Me.adodc1
    
    text2(22).Text = PonerNombreDeCod(Text1(22), "departamento", "nomdepar")
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = PonerContRegistros(Me.adodc1)
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cta As String
Dim cadMen As String
Dim BuscaChekc As String


    b = CompForm(Me)
    If Not b Then Exit Function
    
    ' en el caso de que la tarjeta sea profesional la matricula es obligatoria
    If cmbAux(0).ListIndex = 2 And Text1(12).Text = "" Then
        MsgBox "Si el tipo de tarjeta es Profesional, debe introducir la matrícula", vbExclamation
        PonerFoco Text1(12)
        b = False
    End If
    If b Then
        '[Monica]22/11/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If Text1(11).Text = "" Or Text1(5).Text = "" Or Text1(6).Text = "" Or Text1(7).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            Text1(11).Text = ""
            Text1(4).Text = ""
            Text1(5).Text = ""
            Text1(6).Text = ""
            Text1(7).Text = ""
        Else
            cta = Format(Text1(5).Text, "0000") & Format(Text1(6).Text, "0000") & Format(Text1(7).Text, "00") & Format(Text1(11).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El cliente no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del cliente no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(4)
                    b = False
                End If
            Else
                BuscaChekc = ""
                If Me.Text1(4).Text <> "" Then BuscaChekc = Mid(Text1(4).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(4).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(4).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(4).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(4).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(4)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
'    imgBuscar_Click (indice)
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda(Me, , False)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' ******** Si la clau primaria no es Text1(0), canviar-ho ***********
        PonerFoco Text1(1)
        ' *******************************************************************
    End If
End Sub

Private Sub LimpiarCampos()

    On Error Resume Next

    limpiar Me
'   Me.Combo1(0).ListIndex = -1
    
    ' ****************************************************
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codtraba=" & Text1(0).Text & ")"
    If SituarData(Me.adodc1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Sub printNou()
    
    With frmImprimir2
        .cadTabla2 = "straba"
        .Informe2 = "rManTraba.rpt"
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(adodc1, Me)
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa=" & DBSet(vEmpresa.nomEmpre, "T") & "|" & "'|pOrden={straba.codtraba}|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False

        .Show vbModal
    End With
End Sub

Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    cmbAux(0).Clear
    
    cmbAux(0).AddItem "Bonificado"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 1
    cmbAux(0).AddItem "Normal"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 0
    cmbAux(0).AddItem "Profesional"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 2


    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    cmbAux(1).Clear
    
    cmbAux(1).AddItem "Activa"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 0
    cmbAux(1).AddItem "Inactiva"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 1
    cmbAux(1).AddItem "Perdida"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 2
    cmbAux(1).AddItem "Pendiente Entrega"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 3
    cmbAux(1).AddItem "Bloqueada"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 4


End Sub
